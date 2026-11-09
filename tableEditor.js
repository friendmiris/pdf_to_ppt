import * as pdfjsLib from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

// A table region is intentionally user-confirmed. Rasterized PDF tables have no
// reliable row/column semantics; OCR alone cannot infer merged cells safely.
export function initTableEditor({ getFiles, onSave }) {
  const root = document.querySelector('#tableEditor');
  const status = root.querySelector('#tableStatus');
  const fileSelect = root.querySelector('#tableFile');
  const pageInput = root.querySelector('#tablePage');
  const rowsInput = root.querySelector('#tableRows');
  const colsInput = root.querySelector('#tableCols');
  const canvas = root.querySelector('#tableCanvas');
  const overlay = root.querySelector('#tableOverlay');
  const grid = root.querySelector('#tableGrid');
  const ctx = canvas.getContext('2d');
  let pdfDoc = null;
  let pdfTask = null;
  let activeFile = null;
  let currentPage = 1;
  let region = null;
  let dragStart = null;
  let cellTexts = [];
  let imageCanvas = null;
  let busy = false;

  const error = message => { status.textContent = message; };
  function refreshFiles() {
    const files = getFiles();
    const selected = fileSelect.value;
    fileSelect.replaceChildren();
    files.forEach((file, i) => { const opt = document.createElement('option'); opt.value = String(i); opt.textContent = file.name; fileSelect.append(opt); });
    if (files.length && selected && Number(selected) < files.length) fileSelect.value = selected;
    root.querySelector('#tableLoad').disabled = files.length === 0;
  }
  function drawOverlay() {
    overlay.replaceChildren();
    if (!region) return;
    const marker = document.createElement('div');
    marker.className = 'table-selection';
    marker.style.left = `${region.x * 100}%`;
    marker.style.top = `${region.y * 100}%`;
    marker.style.width = `${region.w * 100}%`;
    marker.style.height = `${region.h * 100}%`;
    const rows = Math.max(1, Number(rowsInput.value) || 1);
    const cols = Math.max(1, Number(colsInput.value) || 1);
    marker.style.backgroundSize = `${100/cols}% ${100/rows}%`;
    for (let i=1; i<cols; i++) {
      const line=document.createElement('div'); line.className='table-guide vertical'; line.style.left=`${i/cols*100}%`; marker.append(line);
    }
    for (let i=1; i<rows; i++) {
      const line=document.createElement('div'); line.className='table-guide horizontal'; line.style.top=`${i/rows*100}%`; marker.append(line);
    }
    overlay.append(marker);
  }
  function buildGrid(values = []) {
    const rows = Math.max(1, Math.min(30, Number(rowsInput.value) || 2));
    const cols = Math.max(1, Math.min(12, Number(colsInput.value) || 2));
    cellTexts = Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => values[r]?.[c] || ''));
    grid.style.gridTemplateColumns = `repeat(${cols}, minmax(120px, 1fr))`;
    grid.replaceChildren();
    for(let r=0; r<rows; r++) for(let c=0; c<cols; c++) {
      const input=document.createElement('textarea'); input.rows=2; input.value=cellTexts[r][c];
      input.setAttribute('aria-label', `${r+1}행 ${c+1}열`);
      input.placeholder=`${r+1}행 ${c+1}열`;
      input.addEventListener('input', () => { cellTexts[r][c]=input.value; });
      grid.append(input);
    }
    drawOverlay();
  }
  function normalizedPosition(event) {
    const bounds = overlay.getBoundingClientRect();
    return { x: Math.max(0,Math.min(1,(event.clientX-bounds.left)/bounds.width)), y: Math.max(0,Math.min(1,(event.clientY-bounds.top)/bounds.height)) };
  }
  overlay.addEventListener('pointerdown', event => {
    if (!imageCanvas) return;
    dragStart=normalizedPosition(event); overlay.setPointerCapture(event.pointerId);
    region={...dragStart,w:0,h:0}; drawOverlay();
  });
  overlay.addEventListener('pointermove', event => {
    if (!dragStart) return;
    const pos=normalizedPosition(event);
    region={x:Math.min(pos.x,dragStart.x), y:Math.min(pos.y,dragStart.y), w:Math.abs(pos.x-dragStart.x), h:Math.abs(pos.y-dragStart.y)};
    drawOverlay();
  });
  overlay.addEventListener('pointerup', () => {
    dragStart=null;
    if(region && (region.w<0.025 || region.h<0.025)) { region=null; drawOverlay(); error('표의 왼쪽 위에서 오른쪽 아래까지 드래그하세요.'); }
    else if(region) error('표 영역을 선택했습니다. 행과 열 개수를 확인하고 OCR을 누르세요.');
  });
  async function loadPage() {
    if(busy) return;
    const files=getFiles(); const file=files[Number(fileSelect.value)];
    if(!file) return error('먼저 PDF 파일을 선택해 주세요.');
    busy=true;
    try {
      if(activeFile!==file) {
        await pdfDoc?.destroy(); pdfDoc=null;
        const data=new Uint8Array(await file.arrayBuffer());
        pdfTask=pdfjsLib.getDocument({data}); pdfDoc=await pdfTask.promise; activeFile=file;
      }
      currentPage=Math.floor(Number(pageInput.value) || 1);
      if(currentPage<1 || currentPage>pdfDoc.numPages) throw Error(`페이지는 1~${pdfDoc.numPages} 사이여야 합니다.`);
      const page=await pdfDoc.getPage(currentPage);
      const base=page.getViewport({scale:1});
      const viewport=page.getViewport({scale:Math.min(1.5,1600/base.width)});
      canvas.width=Math.floor(viewport.width); canvas.height=Math.floor(viewport.height);
      await page.render({canvasContext:ctx,viewport}).promise;
      imageCanvas=canvas;
      overlay.style.width='100%';
      page.cleanup(); region=null; drawOverlay(); buildGrid();
      error(`${file.name} — ${currentPage}/${pdfDoc.numPages}페이지 · 표 영역을 드래그해 선택하세요.`);
    } catch(err) { error(err.message || '페이지 로딩 실패'); }
    finally {busy=false;}
  }
  async function recognize() {
    if(!region) return error('먼저 그림 위에서 표 영역을 드래그해 주세요.');
    if(busy) return;
    busy=true; error('OCR 인식 중… 첫 실행은 한국어 언어 데이터 다운로드가 필요합니다.');
    let worker;
    try {
      // Lazy-load OCR so ordinary PDF conversion stays lightweight.
      const mod=await import('tesseract.js');
      worker=await mod.createWorker('kor+eng',1);
      const result=await worker.recognize(canvas, { rectangle: {
        left: Math.floor(region.x*canvas.width), top:Math.floor(region.y*canvas.height),
        width:Math.floor(region.w*canvas.width), height:Math.floor(region.h*canvas.height)
      } }, { blocks:true });
      const rows=cellTexts.length, cols=cellTexts[0]?.length || 1;
      const values=Array.from({length:rows},()=>Array(cols).fill(''));
      const words=[];
      function gather(node) {
        if(!node) return;
        if(Array.isArray(node)) return node.forEach(gather);
        if(node.bbox && node.text && !node.words && !node.lines && !node.paragraphs && !node.blocks) words.push(node);
        else for(const key of ['blocks','paragraphs','lines','words']) if(node[key]) gather(node[key]);
      }
      gather(result.data.blocks);
      if(!words.length && result.data.text) {
        // OCR provider may not return word boxes. Keep text for manual placement.
        values[0][0]=result.data.text.trim();
      } else {
        const left=region.x*canvas.width, top=region.y*canvas.height;
        words.sort((a,b)=>a.bbox.y0-b.bbox.y0 || a.bbox.x0-b.bbox.x0);
        for(const word of words) {
          const x=((word.bbox.x0+word.bbox.x1)/2-left)/(region.w*canvas.width);
          const y=((word.bbox.y0+word.bbox.y1)/2-top)/(region.h*canvas.height);
          if(x<0||x>1||y<0||y>1) continue;
          const row=Math.min(rows-1, Math.floor(y*rows));
          const col=Math.min(cols-1, Math.floor(x*cols));
          values[row][col] += (values[row][col]?' ':'')+word.text;
        }
      }
      buildGrid(values);
      error('OCR 결과를 셀별로 확인하고 틀린 내용을 수정한 다음 저장하세요.');
    } catch(err) { error(`OCR 실패: ${err.message}. 셀에 글씨를 직접 입력해도 저장할 수 있습니다.`); }
    finally { await worker?.terminate(); busy=false; }
  }
  root.querySelector('#tableLoad').addEventListener('click',loadPage);
  root.querySelector('#tableOCR').addEventListener('click',recognize);
  root.querySelector('#tableSave').addEventListener('click', () => {
    if(!region || !activeFile) return error('표 영역을 먼저 선택해 주세요.');
    const files=getFiles(); const idx=files.indexOf(activeFile);
    if(idx<0) return error('선택한 PDF 파일을 찾을 수 없습니다.');
    const cells=cellTexts.map(row=>row.slice());
    onSave(idx,currentPage,{...region,cells,rows:cells.length,cols:cells[0].length});
    error(`${activeFile.name} ${currentPage}페이지의 ${cells.length}행 × ${cells[0].length}열 표 저장 완료. 변환 시작을 누르세요.`);
  });
  [rowsInput,colsInput].forEach(el=>el.addEventListener('change',()=>buildGrid(cellTexts)));
  refreshFiles();
  return { refreshFiles, reset: () => { activeFile=null; pdfDoc?.destroy(); pdfDoc=null; region=null; grid.replaceChildren(); canvas.width=0; canvas.height=0; drawOverlay(); refreshFiles(); } };
}
