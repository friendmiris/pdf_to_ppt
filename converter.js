import * as pdfjsLib from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import pptxgen from 'pptxgenjs';
import JSZip from 'jszip';

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

const SLIDE_WIDTH = 13.333333;
const SLIDE_HEIGHT = 7.5;
const RESOLUTION = 1.25;
export const LIMITS = Object.freeze({ maxFiles: 30, maxPagesPerFile: 200, maxTotalPages: 1000, maxTotalBytes: 500 * 1024 * 1024, maxFileBytes: 200 * 1024 * 1024, batchFiles: 5, mergedPartPages: 500 });
const MAX_FILES = LIMITS.maxFiles;
const MAX_PAGES = LIMITS.maxPagesPerFile;
const DEFAULT_FONT = 'Arial';

function bytesToDataUrl(canvas) {
  return canvas.toDataURL('image/jpeg', 0.87);
}

function makePresentation(title) {
  const pptx = new pptxgen();
  pptx.defineLayout({ name: 'CUSTOM_WIDE', width: SLIDE_WIDTH, height: SLIDE_HEIGHT });
  pptx.layout = 'CUSTOM_WIDE';
  pptx.author = 'SlideCraft PDF to PPT';
  pptx.subject = 'Editable PowerPoint converted from PDF';
  pptx.title = title;
  pptx.lang = 'ko-KR';
  pptx.theme = { headFontFace: DEFAULT_FONT, bodyFontFace: DEFAULT_FONT, lang: 'ko-KR' };
  return pptx;
}

function sanitizeName(name) {
  return (name.replace(/\.pdf$/i, '').replace(/[\\/:*?"<>|\x00-\x1F]/g, '_').trim() || 'document').slice(0, 90);
}

function report(progress, detail) {
  if (progress) progress(detail);
}

function getCanvasContextWithoutText(realContext) {
  // PDF.js uses native canvas glyph drawing for ordinary fonts. By ignoring
  // fillText/strokeText, we retain PDF pictures and vector art as background
  // while exporting selectable PDF text as editable PPT text boxes.
  // Outline glyphs/Type3 fonts may remain rasterized as part of the background.
  return new Proxy(realContext, {
    get(target, prop) {
      if (prop === 'fillText' || prop === 'strokeText') return () => {};
      const value = Reflect.get(target, prop, target);
      return typeof value === 'function' ? value.bind(target) : value;
    },
    set(target, prop, value) {
      return Reflect.set(target, prop, value, target);
    }
  });
}

async function renderBackground(page, viewport) {
  const renderScale = Math.min(RESOLUTION, 2100 / Math.max(viewport.width, viewport.height));
  const renderViewport = page.getViewport({ scale: renderScale });
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(renderViewport.width));
  canvas.height = Math.max(1, Math.round(renderViewport.height));
  const ctx = canvas.getContext('2d', { alpha: false });
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  try {
    await page.render({ canvasContext: getCanvasContextWithoutText(ctx), viewport: renderViewport }).promise;
  } catch (error) {
    // Some PDF.js versions do not accept proxied canvas contexts.
    // In that case keep all visual content rather than return a blank slide.
    console.warn('Background without text rendering unavailable; using full page image.', error);
    await page.render({ canvasContext: ctx, viewport: renderViewport }).promise;
  }
  const data = bytesToDataUrl(canvas);
  canvas.width = 0;
  canvas.height = 0;
  return data;
}

function extractText(item, viewport, textContent) {
  const [xScale, ySkew, xSkew, yScale, x, y] = pdfjsLib.Util.transform(viewport.transform, item.transform);
  const fontHeight = Math.max(1, Math.hypot(xScale, ySkew));
  const style = textContent.styles[item.fontName] || {};
  const ascent = typeof style.ascent === 'number' ? style.ascent : 0.85;
  const descent = typeof style.descent === 'number' ? style.descent : -0.15;
  // Baseline in PDF coordinates -> approximate top of editable text box.
  const top = y - fontHeight * (ascent || 0.85);
  const width = Math.max(1, item.width * (Math.hypot(viewport.transform[0], viewport.transform[1]) || 1));
  const height = Math.max(fontHeight, fontHeight * (ascent - descent));
  const fontFace = style.fontFamily && !/sans-serif|serif|monospace/i.test(style.fontFamily)
    ? style.fontFamily.split(',')[0].trim().replace(/["']/g, '') : DEFAULT_FONT;
  return { text: item.str, x, top, width, height, fontHeight, fontFace };
}

function addEditableTables(slide, data, regions, sx, sy, scale) {
  let created = 0;
  for (const table of (regions || [])) {
    if (!table?.cells?.length || !table.w || !table.h) continue;
    const rows=table.cells.length, cols=table.cells[0]?.length || 0;
    if (!cols) continue;
    const x=sx+table.x*data.viewport.width*scale;
    const y=sy+table.y*data.viewport.height*scale;
    const w=table.w*data.viewport.width*scale;
    const h=table.h*data.viewport.height*scale;
    const rowItems=table.cells.map((row,r)=>row.map((value,c)=>({
      text:String(value||''),
      options:{
        fill:{color: r===0?'135C42':c===0?'C5E9DC':'FFFFFF'},
        color:r===0?'FFFFFF':'26443B',
        fontFace:'Malgun Gothic',fontSize:Math.max(6, Math.min(18, h/rows*15)),
        bold:r===0||c===0,margin:0.06,valign:'mid',
        border:{type:'solid',color:'7DAF9C',pt:0.65},breakLine:false
      }
    })));
    slide.addTable(rowItems,{
      x,y,w,h,colW:Array(cols).fill(w/cols),rowH:Array(rows).fill(h/rows),
      margin:0.06,border:{type:'solid',color:'7DAF9C',pt:0.65},
      autoFit:false,verbose:false
    });
    created++;
  }
  return created;
}
function isInsideTable(block, regions, viewport) {
  const centerX=(block.x+block.width*0.5)/viewport.width;
  const centerY=(block.top+block.height*0.5)/viewport.height;
  return (regions||[]).some(t=>centerX>=t.x && centerX<=t.x+t.w && centerY>=t.y && centerY<=t.y+t.h);
}
function placePage(pptx, data, originalName, pageIndex, pageCount, options = {}) {
  const slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };
  const { width, height } = data.viewport;
  const scale = Math.min(SLIDE_WIDTH / width, SLIDE_HEIGHT / height);
  const contentWidth = width * scale;
  const contentHeight = height * scale;
  const offsetX = (SLIDE_WIDTH - contentWidth) / 2;
  const offsetY = (SLIDE_HEIGHT - contentHeight) / 2;
  slide.addImage({
    data: data.background,
    x: offsetX, y: offsetY, w: contentWidth, h: contentHeight,
    altText: `Visual background of ${originalName}, page ${pageIndex + 1}`
  });
  let editableCount = 0;
  for (const block of data.blocks) {
    if (!block.text || !block.text.trim() || isInsideTable(block, options.tables, data.viewport)) continue;
    const boxX = offsetX + block.x * scale;
    const boxY = offsetY + block.top * scale;
    const boxW = Math.max(0.06, block.width * scale * 1.26 + 0.065);
    const boxH = Math.max(0.06, block.height * scale * 1.30);
    if (boxX + boxW < 0 || boxY + boxH < 0 || boxX > SLIDE_WIDTH || boxY > SLIDE_HEIGHT) continue;
    try {
      slide.addText(block.text, {
        x: Math.max(0, boxX), y: Math.max(0, boxY),
        w: Math.min(SLIDE_WIDTH - Math.max(0, boxX), boxW),
        h: Math.min(SLIDE_HEIGHT - Math.max(0, boxY), boxH),
        fontFace: block.fontFace,
        fontSize: Math.max(2, Math.min(96, block.fontHeight * scale * 72 * 0.84)),
        color: '222222',
        margin: 0,
        breakLine: false,
        valign: 'mid',
        bold: false,
        transparency: 0,
        isTextBox: true,
        lineSpacingMultiple: 1.0,
        rotate: 0
      });
      editableCount++;
    } catch (error) {
      console.warn('Skipped unsupported PDF text block', error);
    }
  }
  const tableCount = addEditableTables(slide,data,options.tables,offsetX,offsetY,scale);
  slide.addNotes(`원본: ${originalName} / ${pageIndex + 1} of ${pageCount}\n텍스트 요소: ${editableCount}개 · 편집형 표: ${tableCount}개. PDF 벡터/사진은 배경 이미지입니다.`);
  return editableCount;
}

function uniqueNames(files) {
  const seen = new Map();
  return files.map(file => {
    const base = sanitizeName(file.name);
    const counter = (seen.get(base.toLowerCase()) || 0) + 1;
    seen.set(base.toLowerCase(), counter);
    return counter > 1 ? `${base}_${counter}` : base;
  });
}

async function openPdf(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const task = pdfjsLib.getDocument({ data: bytes, useSystemFonts: true });
  return { doc: await task.promise, task };
}

async function preflight(files, onProgress) {
  const accepted = [];
  const failures = [];
  let pageTotal = 0;
  for (let index = 0; index < files.length; index++) {
    const file = files[index];
    let doc;
    try {
      if (!/\.pdf$/i.test(file.name) || file.size === 0) throw Error('올바른 PDF 파일이 아닙니다.');
      if (file.size > LIMITS.maxFileBytes) throw Error('파일당 200MB 제한 초과');
      ({ doc } = await openPdf(file));
      if (doc.numPages > MAX_PAGES) throw Error(`파일당 ${MAX_PAGES}페이지 제한 초과`);
      if (pageTotal + doc.numPages > LIMITS.maxTotalPages) throw Error(`전체 ${LIMITS.maxTotalPages}페이지 제한 초과`);
      pageTotal += doc.numPages;
      accepted.push({ file, pages: doc.numPages, stem: sanitizeName(file.name) });
    } catch (error) {
      failures.push({ name: file.name, reason: error.message || 'PDF 분석 실패' });
    } finally {
      if (doc) await doc.destroy();
    }
    report(onProgress, { phase: 'preflight', percent: Math.round((index + 1) / files.length * 10), file: file.name, fileIndex: index, fileCount: files.length });
  }
  return { accepted, failures, pageTotal };
}

export async function convertPdfFiles(files, onProgress = () => {}, tableEdits = {}) {
  if (!files?.length) throw Error('PDF 파일을 먼저 선택해 주세요.');
  if (files.length > MAX_FILES) throw Error(`PDF는 최대 ${MAX_FILES}개까지만 선택할 수 있습니다.`);
  const totalBytes = files.reduce((sum, file) => sum + file.size, 0);
  if (totalBytes > LIMITS.maxTotalBytes) throw Error('한 번의 작업에서 PDF 총 크기는 500MB 이하여야 합니다.');

  const { accepted, failures, pageTotal } = await preflight(files, onProgress);
  if (!accepted.length) throw Error(`변환 가능한 PDF가 없습니다. ${failures.map(f => `${f.name}: ${f.reason}`).join(' / ')}`);
  const names = uniqueNames(accepted.map(item => item.file));
  const zip = new JSZip();
  const mergedParts = [];
  let merged = makePresentation('전체_통합_1');
  let mergedCount = 0;
  let mergedPartNumber = 1;
  let completedPages = 0;
  let completedFiles = 0;
  let editableTextBoxes = 0;
  let editableTables = 0;

  async function flushMerged() {
    if (!mergedCount) return;
    report(onProgress, { phase: 'merging', percent: Math.min(95, 10 + Math.floor(completedPages / Math.max(1, pageTotal) * 80)) });
    const blob = await merged.write({ outputType: 'blob', compression: true });
    const name = `전체_통합_${String(mergedPartNumber).padStart(2, '0')}.pptx`;
    zip.file(`통합_PPT/${name}`, blob);
    mergedParts.push({ name, blob, pages: mergedCount });
    merged = null;
    mergedCount = 0;
    mergedPartNumber++;
  }

  // 5개 단위로 작업을 나누되, 브라우저 메모리를 아끼기 위해 PDF는 항상 하나씩 처리합니다.
  for (let batch = 0; batch < accepted.length; batch += LIMITS.batchFiles) {
    const group = accepted.slice(batch, batch + LIMITS.batchFiles);
    for (let localIndex = 0; localIndex < group.length; localIndex++) {
      const idx = batch + localIndex;
      const { file, pages } = group[localIndex];
      let doc;
      try {
        ({ doc } = await openPdf(file));
        const individual = makePresentation(names[idx]);
        let produced = 0;
        for (let pageNumber = 1; pageNumber <= pages; pageNumber++) {
          let page;
          try {
            page = await doc.getPage(pageNumber);
            const viewport = page.getViewport({ scale: 1 });
            const textContent = await page.getTextContent({ includeMarkedContent: false });
            const blocks = textContent.items
              .filter(item => typeof item.str === 'string' && item.str.trim().length)
              .map(item => extractText(item, viewport, textContent));
            const background = await renderBackground(page, viewport);
            const data = { viewport: { width: viewport.width, height: viewport.height }, blocks, background };
            const pageTables = tableEdits[idx]?.[pageNumber] || [];
            const count = placePage(individual, data, file.name, pageNumber - 1, pages, {tables:pageTables});
            if (!merged) merged = makePresentation(`전체_통합_${mergedPartNumber}`);
            placePage(merged, data, file.name, pageNumber - 1, pages, {tables:pageTables});
            editableTextBoxes += count;
            editableTables += pageTables.length;
            produced++;
            mergedCount++;
            completedPages++;
            if (mergedCount >= LIMITS.mergedPartPages) await flushMerged();
          } catch (error) {
            failures.push({ name: `${file.name} / ${pageNumber}페이지`, reason: error.message || '페이지 변환 실패' });
          } finally {
            if (page) page.cleanup();
          }
          report(onProgress, {
            phase: 'parsing', percent: Math.min(90, 10 + Math.floor((completedPages + failures.length) / Math.max(1, pageTotal) * 80)),
            file: file.name, fileIndex: idx, fileCount: accepted.length, page: pageNumber, pageCount: pages,
            batch: Math.floor(batch / LIMITS.batchFiles) + 1,
            batches: Math.ceil(accepted.length / LIMITS.batchFiles),
            completedFiles, failures: failures.length
          });
          // Yield the browser event loop so progress and controls can repaint.
          if (pageNumber % 3 === 0) await new Promise(resolve => setTimeout(resolve, 0));
        }
        if (!produced) throw Error('변환 가능한 페이지가 없습니다.');
        report(onProgress, { phase: 'writing', file: file.name, percent: Math.min(93, 10 + Math.floor(completedPages / Math.max(1, pageTotal) * 80)) });
        zip.file(`개별_PPT/${names[idx]}.pptx`, await individual.write({ outputType: 'blob', compression: true }));
        completedFiles++;
      } catch (error) {
        failures.push({ name: file.name, reason: error.message || '파일 변환 실패' });
      } finally {
        if (doc) await doc.destroy();
      }
    }
  }
  await flushMerged();
  if (!completedPages || !completedFiles) throw Error('변환이 모두 실패했습니다. PDF 형식과 암호 설정을 확인해 주세요.');
  // All output parts can be separately downloaded if the ZIP is too large for a browser.
  zip.file('변환_결과.txt', `SlideCraft 변환 결과\n성공한 원본 PDF: ${completedFiles}개\n완료 슬라이드: ${completedPages}장\n편집 가능한 텍스트 상자: ${editableTextBoxes}개\n편집 가능한 표: ${editableTables}개\n통합 PPT: ${mergedParts.length}개 (최대 ${LIMITS.mergedPartPages}장씩)\n실패 또는 건너뛴 항목: ${failures.length}건\n\n${failures.map(f => `${f.name}: ${f.reason}`).join('\n')}\n\n선택·교정한 표만 편집 가능한 PPT 표로 생성됩니다. 본문 전체 OCR 및 모든 표 자동 인식은 지원하지 않습니다.\n`);
  report(onProgress, { phase: 'zipping', percent: 96, failures: failures.length });
  const zipBlob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 3 } });
  report(onProgress, { phase: 'done', percent: 100, failures: failures.length, completedFiles });
  return { zipBlob, mergedParts, files: completedFiles, pages: completedPages, editableTextBoxes, editableTables, failures };
}

export function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 60000);
}
