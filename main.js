import './style.css';
import { initTableEditor } from './tableEditor.js';
import { convertPdfFiles, downloadBlob, LIMITS } from './converter.js';

const app = document.querySelector('#app');
const state = { files: [], isBusy: false, results: null, tableEdits: {} };
const icons = {
  upload: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m17 8-5-5-5 5"/><path d="M12 3v12"/>',
  file: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M8 13h8M8 17h8"/>',
  trash: '<path d="M3 6h18M8 6V4h8v2M6 6l1 15h10l1-15M10 11v6M14 11v6"/>',
  check: '<path d="m20 6-11 11-5-5"/>',
  download: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5M12 15V3"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  lock: '<rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>',
  arrow: '<path d="M5 12h14m-7-7 7 7-7 7"/>',
  layers: '<rect x="8" y="8" width="13" height="13" rx="2"/><path d="M16 8V5a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h3"/>',
  spark: '<path d="m12 3 1.9 5.7L20 11l-6.1 2.3L12 19l-1.9-5.7L4 11l6.1-2.3L12 3Z"/>',
  info: '<circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/>',
  x: '<path d="M18 6 6 18M6 6l12 12"/>'
};
function icon(name, size=20) { return `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${icons[name]}</svg>`; }
function escapeHtml(str) { return str.replace(/[&<>"']/g, char => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[char])); }
function formatBytes(size) { return size < 1024 * 1024 ? `${(size/1024).toFixed(1)} KB` : `${(size/1024/1024).toFixed(1)} MB`; }

app.innerHTML = `
  <header class="header"><div class="container nav"><div class="brand"><div class="brand-mark">${icon('layers', 23)}</div><span>slide<span class="brand-emphasis">craft</span></span><span class="brand-badge">STUDIO</span></div><div class="nav-aside"><span class="status-dot"></span> 브라우저에서 안전하게 변환</div></div></header>
  <main class="container">
    <div class="intro"><div class="eyebrow">${icon('spark', 15)} PDF TO POWERPOINT, REIMAGINED</div><h1>PDF를 <span class="gradient-text">편집 가능한 PPT</span>로.<br/>여러 파일도 한 번에.</h1><p>PDF를 올리기만 하세요. 개별 PPT부터 전체 통합 PPT까지<br class="desktop"/> 한 번에 만들고 ZIP으로 다운로드할 수 있어요.</p><div class="feature-pills"><span>${icon('check', 15)} 텍스트·표 셀 편집 가능</span><span>${icon('check', 15)} 일괄 변환</span><span>${icon('check', 15)} 통합 PPT 자동 생성</span></div></div>
    <section class="workspace">
      <div class="panel-main">
        <div class="panel-heading"><div><div class="overline">STEP 01</div><h2>PDF 파일 업로드</h2><p>여러 개의 PDF를 동시에 선택할 수 있어요.</p></div><span class="counter" id="fileCount">0개 파일</span></div>
        <div class="dropzone" id="dropzone" tabindex="0" role="button" aria-label="PDF 파일 선택"><div class="drop-icon">${icon('upload', 30)}</div><div class="drop-title">여기로 PDF 파일을 끌어다 놓으세요</div><div class="drop-sub">또는 아래 버튼을 눌러 파일을 선택하세요</div><span class="select-button">${icon('plus', 16)} 파일 선택하기</span><span class="drop-caption">PDF · 최대 30개 · 파일당 최대 200페이지</span></div>
        <input id="fileInput" type="file" accept=".pdf,application/pdf" multiple hidden />
        <div class="file-list" id="fileList"></div>
        <div class="panel-divider"></div>
        <div class="panel-heading secondary-heading"><div><div class="overline">STEP 02</div><h2>변환 설정</h2></div></div>
        <div class="option selected"><div class="option-icon">${icon('file', 21)}</div><div><strong>편집 가능한 PowerPoint</strong><p>PDF 글씨는 텍스트 상자로, 지정한 표는 실제 PPT 표로 변환합니다.</p></div><span class="option-check">${icon('check', 17)}</span></div>
        <div class="option selected"><div class="option-icon">${icon('layers', 21)}</div><div><strong>개별 PPT + 통합 PPT</strong><p>각각의 PPT와 전체 통합본을 ZIP에 담습니다.</p></div><span class="option-check">${icon('check', 17)}</span></div>
        <section id="tableEditor" class="table-editor">
          <h3>이미지 표 → 편집 가능한 PPT 표</h3>
          <p>PDF에 표가 이미지로 들어있다면 변환 전에 표 위치를 선택하세요. 셀 내용은 OCR로 추출한 뒤 수정할 수 있습니다. 표가 없는 페이지는 이 단계를 건너뛰세요.</p>
          <div class="table-controls"><label>PDF 파일 <select id="tableFile"></select></label><label>페이지 <input id="tablePage" type="number" min="1" value="5" /></label><button type="button" id="tableLoad">페이지 표시</button></div>
          <div class="table-preview"><canvas id="tableCanvas"></canvas><div id="tableOverlay" aria-label="드래그하여 표 영역 선택"></div></div>
          <div class="table-controls"><label>행 <input id="tableRows" type="number" min="1" max="30" value="6" /></label><label>열 <input id="tableCols" type="number" min="1" max="12" value="3" /></label><button type="button" id="tableOCR">표 글씨 OCR 인식</button></div>
          <p class="table-hint">표 안을 드래그 → 행·열 지정 → OCR 실행 → 아래 셀 글씨 교정 → 저장 순서로 진행하세요.</p>
          <div id="tableGrid" class="table-grid"></div>
          <button type="button" id="tableSave" class="table-save">이 표를 편집형으로 저장</button><p id="tableStatus" class="table-status" role="status"></p>
        </section>
        <section class="limits-panel" aria-label="변환 처리 한도"><strong>현재 프로그램의 처리 한도</strong><div class="limits-grid"><span>한 번에 PDF <b>최대 30개</b></span><span>PDF당 <b>최대 200페이지</b></span><span>전체 <b>최대 1,000페이지</b></span><span>PDF 전체 <b>최대 500MB</b></span><span>파일 하나당 <b>최대 200MB</b></span><span>통합본 <b>500장씩 자동 분할</b></span></div><p>5개 단위로 순차 처리하고 작업 중 오류가 난 PDF는 건너뜁니다. 총 500MB / 1,000장 처리는 검증된 성능 보장이 아닌 프로그램의 제한값이며, 브라우저 메모리에 따라 실패할 수 있습니다.</p></section><div class="privacy-hint">${icon('lock', 17)} 모든 변환은 이 브라우저 안에서 처리됩니다. PDF는 서버에 업로드되지 않아요.</div>
      </div>
      <aside class="side-panel"><div class="side-header"><span class="side-small">OUTPUT PREVIEW</span><h3>변환 결과물</h3><p>PDF를 선택하면 아래 파일을 받을 수 있어요.</p></div><div class="output-list"><div class="output-item"><div class="output-symbol ind">P</div><div><strong>개별 PowerPoint</strong><p>파일별 .pptx</p></div><span class="mini-check">${icon('check', 16)}</span></div><div class="output-item"><div class="output-symbol merge">${icon('layers', 20)}</div><div><strong>전체 통합 PowerPoint</strong><p>전체_통합_01.pptx (500장 초과 시 분할)</p></div><span class="mini-check">${icon('check', 16)}</span></div><div class="output-item"><div class="output-symbol zip">${icon('download', 20)}</div><div><strong>한 번에 ZIP 다운로드</strong><p>모든 결과물을 하나로</p></div><span class="mini-check">${icon('check', 16)}</span></div></div>
        <div class="summary"><div><span>선택된 PDF</span><strong id="summaryCount">0개</strong></div><div><span>총 파일 크기</span><strong id="summarySize">0 MB</strong></div><div><span>변환 방식</span><strong>텍스트 편집형</strong></div></div>
        <button id="convertBtn" class="convert-button" disabled>변환 시작하기 ${icon('arrow', 18)}</button>
        <div class="progress-section" id="progressSection" hidden><div class="file-progress" id="fileProgress"></div><div class="progress-title"><span id="progressLabel">변환 준비 중...</span><b id="progressPct">0%</b></div><div class="progress-track"><div class="progress-fill" id="progressFill"></div></div><p id="progressDetails"></p></div>
        <div class="result" id="result" hidden><div class="result-check">${icon('check', 20)}</div><h3>변환 완료!</h3><p id="resultInfo"></p><button id="zipDownload" class="download-button">${icon('download', 18)} 모든 결과 ZIP 다운로드</button><div id="mergedDownloads"></div><div class="failures" id="failureList" hidden></div></div>
        <p class="warning" id="errorMsg" role="alert" hidden></p><p class="side-note">${icon('info', 14)} 스캔 PDF · 특수 폰트 · 복잡한 벡터 도형은 완전한 편집 복원이 어려울 수 있습니다.</p>
      </aside>
    </section>
    <section class="how-section"><div class="how-intro"><span class="overline">HOW IT WORKS</span><h2>3단계면 끝!</h2></div><div class="how-grid"><div class="how-card"><div class="number">01</div><h3>PDF 업로드</h3><p>여러 개를 한 번에 선택하고 목록을 확인하세요.</p></div><div class="how-card"><div class="number">02</div><h3>자동 변환</h3><p>원본 시각 요소를 보존하며 텍스트를 편집 가능하게 변환해요.</p></div><div class="how-card"><div class="number">03</div><h3>ZIP 다운로드</h3><p>개별 PPT와 통합 PPT를 한 번에 받아보세요.</p></div></div></section>
  </main>
  <footer><div class="container footer-inner"><span>© 2026 SlideCraft Studio</span><span>Private by design · Powered by PDF.js & PptxGenJS</span></div></footer>`;

const fileInput = document.querySelector('#fileInput');
const dropzone = document.querySelector('#dropzone');
const convertBtn = document.querySelector('#convertBtn');
const list = document.querySelector('#fileList');
const resultEl = document.querySelector('#result');
const errorEl = document.querySelector('#errorMsg');
const progressEl = document.querySelector('#progressSection');
const tableEditor = initTableEditor({
  getFiles: () => state.files,
  onSave: (fileIndex, page, region) => {
    (state.tableEdits[fileIndex] ||= {});
    (state.tableEdits[fileIndex][page] ||= []).push(region);
  }
});

function updateFiles() {
  document.querySelector('#fileCount').textContent = `${state.files.length}개 파일`;
  document.querySelector('#summaryCount').textContent = `${state.files.length}개`;
  document.querySelector('#summarySize').textContent = formatBytes(state.files.reduce((sum, file) => sum + file.size, 0));
  convertBtn.disabled = state.files.length === 0 || state.isBusy;
  tableEditor.refreshFiles();
  list.innerHTML = state.files.map((file, i) => `<div class="file-entry"><div class="file-icon">${icon('file', 20)}</div><div class="file-entry-text"><strong title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</strong><span>${formatBytes(file.size)}</span></div><button type="button" class="remove-btn" data-index="${i}" aria-label="${escapeHtml(file.name)} 삭제" ${state.isBusy?'disabled':''}>${icon('x', 17)}</button></div>`).join('') + (state.files.length ? `<button class="add-more" type="button" id="addMore" ${state.isBusy?'disabled':''}>${icon('plus', 15)} PDF 추가하기</button>` : '');
  list.querySelectorAll('.remove-btn').forEach(el => el.addEventListener('click', () => { state.files.splice(Number(el.dataset.index), 1); state.tableEdits = {}; tableEditor.reset(); resetResults(); updateFiles(); }));
  list.querySelector('#addMore')?.addEventListener('click', () => fileInput.click());
}
function resetResults() { state.results = null; resultEl.hidden = true; errorEl.hidden = true; progressEl.hidden = true; }
function addFiles(files) {
  if (state.isBusy) return;
  const candidates = Array.from(files).filter(f => /\.pdf$/i.test(f.name));
  if (candidates.length !== files.length) { showError('PDF 파일만 추가할 수 있습니다.'); return; }
  if (state.files.length + candidates.length > 30) { showError('한 번에 최대 30개까지 업로드할 수 있습니다.'); return; }
  const newSize = state.files.reduce((a,b)=>a+b.size,0) + candidates.reduce((a,b)=>a+b.size,0);
  if (candidates.some(f => f.size > LIMITS.maxFileBytes)) { showError('개별 PDF는 200MB 이하만 가능합니다.'); return; }
  if (newSize > LIMITS.maxTotalBytes) { showError('선택한 PDF 전체 크기는 500MB 이하여야 합니다.'); return; }
  state.files.push(...candidates);
  resetResults(); updateFiles();
}
function showError(text) { errorEl.textContent = text; errorEl.hidden = false; }
dropzone.addEventListener('click', () => { if (!state.isBusy) fileInput.click(); });
dropzone.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); if (!state.isBusy) fileInput.click(); } });
fileInput.addEventListener('change', () => { addFiles(fileInput.files); fileInput.value = ''; });
dropzone.addEventListener('dragover', event => { event.preventDefault(); dropzone.classList.add('dragover'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
dropzone.addEventListener('drop', event => { event.preventDefault(); dropzone.classList.remove('dragover'); addFiles(event.dataTransfer.files); });

convertBtn.addEventListener('click', async () => {
  if (!state.files.length || state.isBusy) return;
  state.isBusy = true; resetResults(); updateFiles();
  convertBtn.innerHTML = `<span class="spinner"></span> 변환 중...`;
  progressEl.hidden = false;
  try {
    const results = await convertPdfFiles([...state.files], updateProgress, state.tableEdits);
    state.results = results;
    resultEl.hidden = false;
    document.querySelector('#resultInfo').textContent = `${results.files}개 PDF · ${results.pages}장 슬라이드 · ${results.editableTextBoxes.toLocaleString()}개 편집 텍스트 · ${results.editableTables || 0}개 편집 표 · 오류 ${results.failures.length}건`;
    const partContainer = document.querySelector('#mergedDownloads');
    partContainer.replaceChildren();
    results.mergedParts.forEach((part) => {
      const button = document.createElement('button');
      button.type = 'button'; button.className = 'plain-button';
      button.textContent = `${part.name} (${part.pages}장) 다운로드`;
      button.addEventListener('click', () => downloadBlob(part.blob, part.name));
      partContainer.append(button);
    });
    const failureList = document.querySelector('#failureList');
    failureList.hidden = results.failures.length === 0;
    failureList.textContent = results.failures.length ? `건너뛴 항목: ${results.failures.map(f => `${f.name} (${f.reason})`).join(' / ')}` : '';
    resultEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    // Keep downloads user-initiated, especially for large ZIPs or popup-blocking browsers.
  } catch (error) {
    console.error(error);
    showError(error.message || '변환 도중 오류가 발생했습니다. 파일을 확인하고 다시 시도하세요.');
  } finally {
    state.isBusy = false; updateFiles();
    convertBtn.innerHTML = `다시 변환하기 ${icon('arrow', 18)}`;
  }
});
function updateProgress(detail) {
  const labelMap = { preflight:'처리 한도 확인 중', parsing:'PDF 분석 중', writing:'개별 PPT 제작 중', merging:'통합 PPT 제작 중', zipping:'ZIP 압축 중', done:'변환 완료' };
  document.querySelector('#progressLabel').textContent = labelMap[detail.phase] || '변환 중';
  document.querySelector('#progressPct').textContent = `${detail.percent}%`;
  document.querySelector('#progressFill').style.width = `${detail.percent}%`;
  document.querySelector('#progressDetails').textContent = detail.file ? `${detail.file}${detail.page ? ` · ${detail.page}/${detail.pageCount} 페이지` : ''}${detail.batch ? ` · 묶음 ${detail.batch}/${detail.batches}` : ''}` : '결과물을 준비하고 있습니다.';
  if (detail.completedFiles !== undefined) document.querySelector('#fileProgress').textContent = `파일 완료 ${detail.completedFiles}개 · 오류/건너뜀 ${detail.failures || 0}건`;
}
document.querySelector('#zipDownload').addEventListener('click', () => state.results && downloadBlob(state.results.zipBlob, 'SlideCraft_PPT_개별_통합.zip'));
updateFiles();
