/**
 * 敬拜歌譜無損轉調工具 — Google Docs 原生直連編輯版 (Google Apps Script)
 * ----------------------------------------------------------------------
 * 特色：
 * 1. 【100% 零破版】不經過 .docx 轉檔，直接透過 Google Docs 原生文字節點就地替換和弦，
 *    完整保留原文件的字體、字級、Tab 定位點、段落縮排、對齊空白、灰底色塊與表格！
 * 2. 【安全防呆機制】預設「先建立複本再轉調」，絕不修改原始檔案；亦支援貼上「已開好的複本可編輯連結」直接轉調。
 * 3. 【雙模式運行】
 *    - 模式 A（零設定試用）：在任何 Google Docs 內點「擴充功能 -> Apps Script」貼上即可於選單直接啟用。
 *    - 模式 B（團隊網頁版）：部署為「網頁應用程式」後，貼上連結即可一鍵生成轉調複本與乾淨 PDF 下載連結。
 */

const NOTES_SHARP = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B'];
const NOTES_FLAT  = ['C', 'Db', 'D', 'Eb', 'E', 'F', 'Gb', 'G', 'Ab', 'A', 'Bb', 'B'];

const NOTE_TO_SEMI = {
  'C': 0, 'B#': 0,
  'C#': 1, 'Db': 1, 'C♯': 1, 'D♭': 1,
  'D': 2,
  'D#': 3, 'Eb': 3, 'D♯': 3, 'E♭': 3,
  'E': 4, 'Fb': 4,
  'F': 5, 'E#': 5,
  'F#': 6, 'Gb': 6, 'F♯': 6, 'G♭': 6,
  'G': 7,
  'G#': 8, 'Ab': 8, 'G♯': 8, 'A♭': 8,
  'A': 9,
  'A#': 10, 'Bb': 10, 'A♯': 10, 'B♭': 10,
  'B': 11, 'Cb': 11
};

const FLAT_KEYS = new Set(['F', 'Bb', 'Eb', 'Ab', 'Db', 'Gb', 'Dm', 'Gm', 'Cm', 'Fm', 'Bbm', 'Ebm']);
const SCALE_DEGREES_SHARP = ['1', '#1', '2', '#2', '3', '4', '#4', '5', '#5', '6', 'b7', '7'];
const SCALE_DEGREES_FLAT  = ['1', 'b2', '2', 'b3', '3', '4', 'b5', '5', 'b6', '6', 'b7', '7'];

const CHORD_TOKEN_REGEX = /(?<![A-Za-z0-9])([A-G][#b♯♭]?)((?:maj|min|m|M|dim|aug|sus|add|no|[0-9]|[\+\^\(\)])*)(?:\/([A-G][#b♯♭]?))?(?![A-Za-z0-9])/g;
const SINGLE_CHORD_REGEX = /^([A-G][#b♯♭]?)((?:maj|min|m|M|dim|aug|sus|add|no|[0-9]|[\+\^\(\)])*)(?:\/([A-G][#b♯♭]?))?$/;
const INLINE_BRACKET_REGEX = /\[([A-G][#b♯♭]?)((?:maj|min|m|M|dim|aug|sus|add|no|[0-9]|[\+\^\(\)])*)(?:\/([A-G][#b♯♭]?))?\]/g;

const SECTION_HEADERS = new Set([
  'intro', 'verse', 'verse1', 'verse2', 'verse3', 'verse4',
  'pre-chorus', 'prechorus', 'chorus', 'chorus1', 'chorus2',
  'bridge', 'bridge1', 'bridge2', 'tag', 'outro', 'interlude',
  'ending', 'solo', 'instrumental', 'vamp', 'turnaround'
]);

function onOpen() {
  DocumentApp.getUi()
    .createMenu('🎵 敬拜歌譜轉調')
    .addItem('開啟無損轉調面板', 'showTransposerSidebar')
    .addToUi();
}

function showTransposerSidebar() {
  const html = HtmlService.createHtmlOutput(getSidebarHtml_(false))
    .setTitle('敬拜歌譜無損轉調')
    .setWidth(340);
  DocumentApp.getUi().showSidebar(html);
}

function doGet() {
  return HtmlService.createHtmlOutput(getSidebarHtml_(true))
    .setTitle('Google Docs 歌譜直連無損轉調')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function isSectionHeader_(line) {
  const s = line.trim().toLowerCase().replace(/:$/, '');
  if (!s) return false;
  if (SECTION_HEADERS.has(s)) return true;
  return /^(intro|verse|chorus|bridge|tag|outro|interlude|pre-?chorus|v)\s*\d*$/i.test(s);
}

function isChordLine_(line) {
  const stripped = line.trim();
  if (!stripped || isSectionHeader_(stripped)) return false;
  const chineseCount = (stripped.match(/[\u4e00-\u9fff]/g) || []).length;
  const tokens = stripped.split(/[\s｜|\-–—_\/、,，\(\)\[\]]+/).filter(Boolean);
  if (tokens.length === 0) return false;

  let chordMatches = 0;
  for (const tok of tokens) {
    if (SINGLE_CHORD_REGEX.test(tok)) chordMatches++;
  }
  const hasGrid = stripped.includes('｜') || stripped.includes('|') || stripped.includes('- -');
  if (hasGrid && chordMatches >= 1) return true;
  if (chineseCount > 0 && !hasGrid) return false;
  if (tokens.length <= 16 && (chordMatches / tokens.length) >= 0.6) {
    const englishWords = tokens.filter(t => t.length > 1 && /^[A-Za-z]+$/.test(t) && !/^[A-G](m|maj|min|dim|aug|sus|add)$/.test(t));
    if (englishWords.length >= 2) return false;
    return true;
  }
  return false;
}

function shouldUseFlats_(targetKey, accPref) {
  if (accPref === 'flat') return true;
  if (accPref === 'sharp') return false;
  return FLAT_KEYS.has(targetKey);
}

function transposeChordParts_(root, quality, bass, srcSemi, dstSemi, useFlats, mode) {
  const rootSemi = NOTE_TO_SEMI[root];
  if (rootSemi === undefined) return root + (quality || '') + (bass ? '/' + bass : '');

  const newRootSemi = (rootSemi - srcSemi + dstSemi + 120) % 12;
  const newRootNote = useFlats ? NOTES_FLAT[newRootSemi] : NOTES_SHARP[newRootSemi];
  const degDiff = (newRootSemi - dstSemi + 120) % 12;
  const newRootNum = useFlats ? SCALE_DEGREES_FLAT[degDiff] : SCALE_DEGREES_SHARP[degDiff];

  let newBassNote = '';
  let newBassNum = '';
  if (bass) {
    const bassSemi = NOTE_TO_SEMI[bass];
    if (bassSemi !== undefined) {
      const nbSemi = (bassSemi - srcSemi + dstSemi + 120) % 12;
      newBassNote = useFlats ? NOTES_FLAT[nbSemi] : NOTES_SHARP[nbSemi];
      const bDiff = (nbSemi - dstSemi + 120) % 12;
      newBassNum = useFlats ? SCALE_DEGREES_FLAT[bDiff] : SCALE_DEGREES_SHARP[bDiff];
    } else {
      newBassNote = bass;
      newBassNum = bass;
    }
  }

  const chordStr = newRootNote + (quality || '') + (newBassNote ? '/' + newBassNote : '');
  const numStr = newRootNum + (quality || '') + (newBassNum ? '/' + newBassNum : '');
  if (mode === 'number') return numStr;
  if (mode === 'dual') return `${chordStr}(${numStr})`;
  return chordStr;
}

function detectKeyFromParagraphs_(texts) {
  for (const line of texts.slice(0, 8)) {
    const m = line.trim().match(/(?:\/|Key\s*[:：]\s*|調性\s*[:：]\s*)([A-G][#b♯♭]?m?)\s*$/i);
    if (m) {
      const k = m[1];
      const root = k[0].toUpperCase() + k.slice(1).replace('m', '');
      if (NOTE_TO_SEMI[root] !== undefined) return root;
    }
  }
  const found = [];
  for (const line of texts) {
    if (isChordLine_(line)) {
      CHORD_TOKEN_REGEX.lastIndex = 0;
      let m;
      while ((m = CHORD_TOKEN_REGEX.exec(line)) !== null) {
        found.push(m[1]);
      }
    }
  }
  return found.length > 0 ? found[found.length - 1] : 'G';
}

/**
 * 計算單行最小化就地替換清單 (不破壞既有 Tab、前導空白與字元樣式)
 */
function computeMinimalEditsForLine_(line, srcKey, dstKey, accPref, mode) {
  const srcSemi = NOTE_TO_SEMI[srcKey] !== undefined ? NOTE_TO_SEMI[srcKey] : 7;
  const dstSemi = NOTE_TO_SEMI[dstKey] !== undefined ? NOTE_TO_SEMI[dstKey] : srcSemi;
  const useFlats = shouldUseFlats_(dstKey, accPref || 'auto');

  // 1. Metadata line (e.g. "JPCC/約書亞/G")
  const metaMatch = line.match(/(\/|Key\s*[:：]\s*|調性\s*[:：]\s*)([A-G][#b♯♭]?m?)(\s*)$/);
  if (metaMatch && !isChordLine_(line)) {
    const start = metaMatch.index + metaMatch[1].length;
    const end = start + metaMatch[2].length;
    const label = mode === 'number' ? `${dstKey} (級數)` : dstKey;
    if (label !== metaMatch[2]) {
      return [{ start, end, text: label }];
    }
    return [];
  }

  // 2. Inline bracket chords [G]
  INLINE_BRACKET_REGEX.lastIndex = 0;
  const bracketMatches = [...line.matchAll(INLINE_BRACKET_REGEX)];
  if (bracketMatches.length > 0) {
    const edits = [];
    for (const m of bracketMatches) {
      const newTok = '[' + transposeChordParts_(m[1], m[2] || '', m[3], srcSemi, dstSemi, useFlats, mode) + ']';
      if (newTok !== m[0]) {
        edits.push({ start: m.index, end: m.index + m[0].length, text: newTok });
      }
    }
    edits.sort((a, b) => b.start - a.start || b.end - a.end);
    return edits;
  }

  if (!isChordLine_(line)) return [];

  // 3. Chord line with smart column-alignment preservation
  CHORD_TOKEN_REGEX.lastIndex = 0;
  const matches = [...line.matchAll(CHORD_TOKEN_REGEX)];
  if (matches.length === 0) return [];

  const edits = [];
  let cursor = 0;
  let debt = 0;

  for (const m of matches) {
    const gapStart = cursor;
    const gapEnd = m.index;
    const gap = line.slice(gapStart, gapEnd);

    if (debt > 0) {
      const mSpaces = gap.match(/ {2,}/);
      if (mSpaces) {
        const consume = Math.min(debt, mSpaces[0].length - 1);
        if (consume > 0) {
          edits.push({
            start: gapStart + mSpaces.index,
            end: gapStart + mSpaces.index + consume,
            text: ''
          });
          debt -= consume;
        }
      }
      if (debt > 0) {
        const trail = gap.length - gap.trimEnd().length;
        const consume = Math.min(debt, Math.max(0, trail - 1));
        if (consume > 0) {
          edits.push({ start: gapEnd - consume, end: gapEnd, text: '' });
          debt -= consume;
        }
      }
    } else if (debt < 0) {
      const pad = -debt;
      const pipeIdx = gap.indexOf('｜') >= 0 ? gap.indexOf('｜') : gap.indexOf('|');
      const insertAt = pipeIdx >= 0 ? gapStart + pipeIdx : gapEnd;
      edits.push({ start: insertAt, end: insertAt, text: ' '.repeat(pad) });
      debt = 0;
    }

    const oldTok = m[0];
    const newTok = transposeChordParts_(m[1], m[2] || '', m[3], srcSemi, dstSemi, useFlats, mode);
    if (newTok !== oldTok) {
      edits.push({ start: m.index, end: m.index + oldTok.length, text: newTok });
    }
    debt += newTok.length - oldTok.length;
    cursor = m.index + oldTok.length;
  }

  edits.sort((a, b) => b.start - a.start || b.end - a.end);
  return edits;
}

function collectParagraphsFromContainer_(container, outList) {
  const numChildren = container.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const child = container.getChild(i);
    const type = child.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH || type === DocumentApp.ElementType.LIST_ITEM) {
      outList.push(child);
    } else if (type === DocumentApp.ElementType.TABLE) {
      const table = child.asTable();
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        for (let c = 0; c < row.getNumCells(); c++) {
          collectParagraphsFromContainer_(row.getCell(c), outList);
        }
      }
    }
  }
}

function getActiveDocInfo() {
  const doc = DocumentApp.getActiveDocument();
  const paragraphs = [];
  collectParagraphsFromContainer_(doc.getBody(), paragraphs);
  const texts = paragraphs.map(p => p.getText());
  return {
    title: doc.getName(),
    url: doc.getUrl(),
    detectedKey: detectKeyFromParagraphs_(texts)
  };
}

/**
 * 執行 Google Docs 原地轉調（可選擇先建立複本，或直接編輯指定文件）
 */
function runTransposeInGoogleDocs(params) {
  const targetKey = params.targetKey || 'F';
  const accPref = params.accPref || 'auto';
  const mode = params.mode || 'chord';
  const createCopy = params.createCopy !== false;

  let doc;
  if (params.docUrlOrId && params.docUrlOrId.trim()) {
    const m = params.docUrlOrId.match(/\/d\/([a-zA-Z0-9_-]{15,})/);
    const fileId = m ? m[1] : params.docUrlOrId.trim();
    if (createCopy) {
      const origFile = DriveApp.getFileById(fileId);
      const copyTitle = `${origFile.getName().replace(/\s+[A-G][#b♯♭]?m?$/, '')} (${targetKey}調)`;
      const copyFile = origFile.makeCopy(copyTitle);
      doc = DocumentApp.openById(copyFile.getId());
    } else {
      doc = DocumentApp.openById(fileId);
    }
  } else {
    const activeDoc = DocumentApp.getActiveDocument();
    if (createCopy) {
      const origFile = DriveApp.getFileById(activeDoc.getId());
      const copyTitle = `${origFile.getName().replace(/\s+[A-G][#b♯♭]?m?$/, '')} (${targetKey}調)`;
      const copyFile = origFile.makeCopy(copyTitle);
      doc = DocumentApp.openById(copyFile.getId());
    } else {
      doc = activeDoc;
    }
  }

  const paragraphs = [];
  collectParagraphsFromContainer_(doc.getBody(), paragraphs);
  const texts = paragraphs.map(p => p.getText());
  const srcKey = (params.sourceKey && params.sourceKey !== 'auto') ? params.sourceKey : detectKeyFromParagraphs_(texts);

  let modifiedParagraphs = 0;
  for (const para of paragraphs) {
    const fullText = para.getText();
    if (!fullText.trim()) continue;
    const edits = computeMinimalEditsForLine_(fullText, srcKey, targetKey, accPref, mode);
    if (edits.length === 0) continue;

    const textEl = para.editAsText();
    for (const ed of edits) {
      // 先插入新文字再刪除舊文字，可 100% 保留該位置的字體、顏色、字級與灰底樣式
      if (ed.text.length > 0) {
        textEl.insertText(ed.start, ed.text);
      }
      if (ed.end > ed.start) {
        const delStart = ed.start + ed.text.length;
        const delEndInclusive = ed.end + ed.text.length - 1;
        textEl.deleteText(delStart, delEndInclusive);
      }
    }
    modifiedParagraphs++;
  }

  doc.saveAndClose();
  const docId = doc.getId();
  return {
    ok: true,
    docId: docId,
    title: doc.getName(),
    docUrl: `https://docs.google.com/document/d/${docId}/edit`,
    pdfUrl: `https://docs.google.com/document/d/${docId}/export?format=pdf`,
    srcKey: srcKey,
    targetKey: targetKey,
    modifiedParagraphs: modifiedParagraphs
  };
}

function getSidebarHtml_(isWebApp) {
  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="utf-8">
  <style>
    body { font-family: "Noto Sans TC", Arial, sans-serif; margin: 0; padding: 14px; background: #faf9f6; color: #18181b; font-size: 13px; }
    .card { background: #fff; border: 1px solid #e4e0d8; border-radius: 10px; padding: 14px; margin-bottom: 12px; }
    h2 { font-size: 15px; margin: 0 0 8px 0; color: #b45309; }
    label { display: block; font-weight: 600; margin: 10px 0 4px; font-size: 12px; color: #52525b; }
    input[type="text"], select { width: 100%; box-sizing: border-box; padding: 7px 9px; border: 1px solid #d4d4d8; border-radius: 6px; font-size: 13px; }
    .key-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 5px; margin-top: 6px; }
    .key-btn { padding: 7px 4px; border: 1px solid #d4d4d8; background: #fff; border-radius: 6px; cursor: pointer; font-weight: 600; }
    .key-btn.active { background: #b45309; color: #fff; border-color: #b45309; }
    .btn-primary { width: 100%; padding: 10px; background: #b45309; color: #fff; border: none; border-radius: 8px; font-size: 13.5px; font-weight: 700; cursor: pointer; margin-top: 12px; }
    .btn-primary:hover { background: #92400e; }
    .btn-primary:disabled { opacity: 0.6; cursor: wait; }
    .check-row { display: flex; align-items: flex-start; gap: 8px; margin-top: 10px; font-size: 12px; color: #047857; background: #ecfdf5; border: 1px solid #a7f3d0; padding: 8px; border-radius: 6px; }
    .result-box { display: none; margin-top: 12px; padding: 10px; background: #fffbeb; border: 1px solid #f59e0b; border-radius: 8px; }
    .result-box a { display: inline-block; margin-top: 6px; margin-right: 8px; padding: 6px 10px; background: #b45309; color: #fff; text-decoration: none; border-radius: 6px; font-size: 12px; font-weight: 600; }
    .result-box a.sec { background: #18181b; }
  </style>
</head>
<body>
  <div class="card" style="${isWebApp ? 'max-width:460px;margin:20px auto;' : ''}">
    <h2>🎵 Google Docs 原生無損轉調</h2>
    ${isWebApp ? `
      <label>Google Docs 歌譜連結（或已開好的複本可編輯連結）</label>
      <input type="text" id="docUrl" placeholder="貼上 https://docs.google.com/document/d/.../edit" />
    ` : `
      <div id="activeInfo" style="font-size:12px;color:#52525b;">偵測中...</div>
    `}
    <label>原調（預設自動偵測）</label>
    <select id="srcKey">
      <option value="auto">自動偵測原調</option>
      ${['C','Db','D','Eb','E','F','F#','G','Ab','A','Bb','B'].map(k => `<option value="${k}">${k}</option>`).join('')}
    </select>

    <label>選擇目標調性</label>
    <div class="key-grid" id="keyGrid">
      ${['C','Db','D','Eb','E','F','F#','G','Ab','A','Bb','B'].map(k => `<button type="button" class="key-btn ${k==='F'?'active':''}" data-k="${k}">${k}</button>`).join('')}
    </div>

    <label>譜面模式</label>
    <select id="mode">
      <option value="chord">標準和弦譜</option>
      <option value="number">級數譜 (1-7)</option>
      <option value="dual">和弦 + 級數</option>
    </select>

    <label>升降記號</label>
    <select id="accPref">
      <option value="auto">依調性自動</option>
      <option value="sharp">強制升記號 (♯)</option>
      <option value="flat">強制降記號 (♭)</option>
    </select>

    <div class="check-row">
      <input type="checkbox" id="createCopy" checked style="margin-top:2px;">
      <label for="createCopy" style="margin:0;color:#047857;cursor:pointer;">
        先在我的雲端硬碟建立「複本」再轉調（推薦，絕對不改動原始歌譜）
      </label>
    </div>

    <button type="button" class="btn-primary" id="btnRun">開始原生轉調</button>

    <div class="result-box" id="resultBox"></div>
  </div>

  <script>
    let targetKey = 'F';
    document.querySelectorAll('.key-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.key-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        targetKey = btn.dataset.k;
      });
    });

    if (!${isWebApp}) {
      google.script.run.withSuccessHandler(info => {
        document.getElementById('activeInfo').textContent = '目前文件：' + info.title + '（偵測原調：' + info.detectedKey + '）';
      }).getActiveDocInfo();
    }

    document.getElementById('btnRun').addEventListener('click', () => {
      const btn = document.getElementById('btnRun');
      const box = document.getElementById('resultBox');
      btn.disabled = true;
      btn.textContent = '正在原地替換和弦...';
      box.style.display = 'none';

      const payload = {
        docUrlOrId: document.getElementById('docUrl') ? document.getElementById('docUrl').value : '',
        sourceKey: document.getElementById('srcKey').value,
        targetKey: targetKey,
        mode: document.getElementById('mode').value,
        accPref: document.getElementById('accPref').value,
        createCopy: document.getElementById('createCopy').checked
      };

      google.script.run
        .withSuccessHandler(res => {
          btn.disabled = false;
          btn.textContent = '開始原生轉調';
          box.style.display = 'block';
          box.innerHTML = '<div><strong>✅ 轉調完成！</strong>（' + res.srcKey + ' → ' + res.targetKey + '，共更新 ' + res.modifiedParagraphs + ' 行）</div>' +
            '<a href="' + res.docUrl + '" target="_blank">開啟轉調後 Google Docs</a>' +
            '<a href="' + res.pdfUrl + '" target="_blank" class="sec">直接下載乾淨 PDF</a>';
        })
        .withFailureHandler(err => {
          btn.disabled = false;
          btn.textContent = '開始原生轉調';
          alert('發生錯誤：' + err.message);
        })
        .runTransposeInGoogleDocs(payload);
    });
  </script>
</body>
</html>`;
}
