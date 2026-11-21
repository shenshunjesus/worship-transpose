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

/**
 * 將全形 ASCII（Ａ-Ｚ、ａ-ｚ、０-９、＃、／、｜、－ 等 U+FF01..U+FF5E）與全形空白/不換行空白
 * 1 對 1 正規化為半形 ASCII，確保回傳字串長度與原字串 100% 相同（字元索引偏移量完全不變）。
 * 解決中文輸入法誤打全形和弦（如 Ｇ、Ｃ、Ｅｍ、Ｇ/B、Ｄ／Ｆ＃）無法被偵測與轉調的問題。
 */
function normalizeFullwidthForMatch_(str) {
  let out = '';
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);
    if (code >= 0xFF01 && code <= 0xFF5E) {
      out += String.fromCharCode(code - 0xFEE0);
    } else if (code === 0x3000 || code === 0x00A0) {
      out += ' ';
    } else {
      out += str[i];
    }
  }
  return out;
}

function isSectionHeader_(line) {
  const s = normalizeFullwidthForMatch_(line).trim().toLowerCase().replace(/:$/, '');
  if (!s) return false;
  if (SECTION_HEADERS.has(s)) return true;
  return /^(?:->)?\s*(intro|verse|chorus|bridge|tag|outro|interlude|pre-?chorus|v|pc|c|b)\s*\d*$/i.test(s);
}

function isChordLine_(line) {
  const norm = normalizeFullwidthForMatch_(line).trim();
  if (!norm || isSectionHeader_(norm)) return false;
  const chineseCount = (norm.match(/[\u4e00-\u9fff]/g) || []).length;
  const tokens = norm.split(/[\s｜|\-–—―─_\/、,，\(\)\[\]\*]+/).filter(Boolean);
  if (tokens.length === 0) return false;

  let chordMatches = 0;
  for (const tok of tokens) {
    if (SINGLE_CHORD_REGEX.test(tok)) chordMatches++;
  }
  const hasGrid = /[｜|]/.test(norm) || /[-–—―─]\s*[-–—―─]/.test(norm);
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

/**
 * 偵測譜頭調性資訊行（回傳 { start, end, key, root }，座標與原字串 1:1 對齊）
 * 支援：
 * - "G key / 可萌 / 蛤 / 68拍"、"Gkey / ..."、"Ｇ key / ..."、"Eb調 / 約書亞"
 * - "Key: G / 68拍"、"Key：Eb"、"調性：D"、"原調：Bb"
 * - "JPCC/約書亞/G"、"Upper Room/活水/D"
 */
function matchMetadataKey_(line) {
  if (!line || !line.trim() || isChordLine_(line)) return null;
  const norm = normalizeFullwidthForMatch_(line);

  // 1. 前綴標記：Key: G / Key：G / 調性：G / 原調：G（可出現在行中任意位置）
  let m = norm.match(/(?:^|[\s\/|(\[（【,;，；])((?:原調|調性|[Kk][Ee][Yy])\s*[:：]?\s*)([A-G][#b♯♭]?m?)(?=$|[\s\/|)\]）】.,;，；])/);
  if (m) {
    const keyStr = m[2];
    const root = keyStr[0].toUpperCase() + keyStr.slice(1).replace(/m$/, '');
    if (NOTE_TO_SEMI[root] !== undefined) {
      const start = m.index + (m[0].length - keyStr.length);
      return { start, end: start + keyStr.length, key: keyStr, root };
    }
  }

  // 2. 後綴標記：G key / Gkey / G Key / G調 / Eb 調（可出現在行中任意位置）
  m = norm.match(/(?:^|[\s\/|(\[（【,;，；])([A-G][#b♯♭]?m?)(\s*(?:[Kk][Ee][Yy]|調))(?=$|[\s\/|)\]）】.,;，；])/);
  if (m) {
    const keyStr = m[1];
    const suffix = m[2];
    const root = keyStr[0].toUpperCase() + keyStr.slice(1).replace(/m$/, '');
    const isEnglishSentence = !/[\u4e00-\u9fff\/|：:]/.test(norm) && (norm.match(/[A-Za-z]{2,}/g) || []).length >= 3;
    if (NOTE_TO_SEMI[root] !== undefined && !isEnglishSentence) {
      const start = m.index + (m[0].length - keyStr.length - suffix.length);
      return { start, end: start + keyStr.length, key: keyStr, root };
    }
  }

  // 3. 行尾斜線標記：JPCC/約書亞/G、Upper Room/活水/D
  m = norm.match(/(\/\s*)([A-G][#b♯♭]?m?)(\s*)$/);
  if (m) {
    const keyStr = m[2];
    const root = keyStr[0].toUpperCase() + keyStr.slice(1).replace(/m$/, '');
    if (NOTE_TO_SEMI[root] !== undefined) {
      const start = m.index + m[1].length;
      return { start, end: start + keyStr.length, key: keyStr, root };
    }
  }

  return null;
}

function detectKeyFromTitle_(docTitle) {
  if (!docTitle) return null;
  const norm = normalizeFullwidthForMatch_(String(docTitle)).trim();
  const meta = matchMetadataKey_(norm);
  if (meta) return meta.root;
  const m = norm.match(/(?:^|[\s\-_\/(])([A-G][#b♯♭]?m?)(?:\s*(?:[Kk][Ee][Yy]|調))?[)]?\s*$/);
  if (m) {
    const k = m[1];
    const root = k[0].toUpperCase() + k.slice(1).replace(/m$/, '');
    if (NOTE_TO_SEMI[root] !== undefined) return root;
  }
  return null;
}

function buildCopyTitle_(origName, targetKey) {
  const cleaned = String(origName || '歌譜')
    .replace(/(?:\s*[\(（]\s*[A-GＡ-Ｇ][#b♯♭＃ｂ]?m?(?:\s*(?:key|調))?\s*[\)）]|\s+[A-GＡ-Ｇ][#b♯♭＃ｂ]?m?(?:\s*(?:key|調))?)\s*$/i, '')
    .trim();
  return `${cleaned || origName} (${targetKey}調)`;
}

function detectKeyFromParagraphs_(texts, docTitle) {
  const lines = [];
  for (const t of texts) {
    for (const sub of String(t || '').split(/[\r\n\u000b]+/)) {
      lines.push(sub);
    }
  }
  // 1. 優先從前 12 行譜頭偵測明確調性標記（如 "G key / 可萌 / 68拍"、"JPCC/約書亞/G"）
  for (const line of lines.slice(0, 12)) {
    const meta = matchMetadataKey_(line);
    if (meta) return meta.root;
  }
  // 2. 次優先從文件標題偵測（如 "受難路 G"、"因祢十架 (A調)"）
  const titleKey = detectKeyFromTitle_(docTitle);
  if (titleKey) return titleKey;

  // 3. 最後以全曲和弦（含全形正規化後）進行順階和弦計分推測大調主音
  const chords = [];
  for (const line of lines) {
    if (isChordLine_(line)) {
      const norm = normalizeFullwidthForMatch_(line);
      CHORD_TOKEN_REGEX.lastIndex = 0;
      let m;
      while ((m = CHORD_TOKEN_REGEX.exec(norm)) !== null) {
        const semi = NOTE_TO_SEMI[m[1]];
        if (semi !== undefined) {
          const q = m[2] || '';
          const isMinor = /^m(?!aj)/i.test(q) || /^min/i.test(q);
          chords.push({ root: m[1], semi, isMinor });
        }
      }
    }
  }
  if (chords.length === 0) return 'G';

  let bestKeySemi = chords[chords.length - 1].semi;
  let bestScore = -1e9;
  for (let k = 0; k < 12; k++) {
    let score = 0;
    for (const c of chords) {
      const diff = (c.semi - k + 12) % 12;
      if (!c.isMinor && diff === 0) score += 3;
      else if (!c.isMinor && (diff === 5 || diff === 7)) score += 2;
      else if (c.isMinor && diff === 9) score += 2;
      else if (c.isMinor && (diff === 2 || diff === 4)) score += 1;
      else score -= 2;
    }
    const first = chords[0];
    const last = chords[chords.length - 1];
    if (!last.isMinor && last.semi === k) score += 3;
    if (!first.isMinor && first.semi === k) score += 1.5;
    if (score > bestScore) {
      bestScore = score;
      bestKeySemi = k;
    }
  }
  const PREFERRED_NAMES = ['C', 'Db', 'D', 'Eb', 'E', 'F', 'F#', 'G', 'Ab', 'A', 'Bb', 'B'];
  return PREFERRED_NAMES[bestKeySemi];
}

/**
 * 計算單行最小化就地替換清單 (不破壞既有 Tab、前導空白與字元樣式，並完整支援全形和弦)
 */
function computeMinimalEditsForLine_(line, srcKey, dstKey, accPref, mode) {
  const srcSemi = NOTE_TO_SEMI[srcKey] !== undefined ? NOTE_TO_SEMI[srcKey] : 7;
  const dstSemi = NOTE_TO_SEMI[dstKey] !== undefined ? NOTE_TO_SEMI[dstKey] : srcSemi;
  const useFlats = shouldUseFlats_(dstKey, accPref || 'auto');
  const normLine = normalizeFullwidthForMatch_(line);

  // 1. Metadata line (e.g. "G key / 可萌 / 蛤 / 68拍", "JPCC/約書亞/G")
  const metaMatch = matchMetadataKey_(line);
  if (metaMatch) {
    const label = mode === 'number' ? `${dstKey} (級數)` : dstKey;
    if (label !== metaMatch.key) {
      return [{ start: metaMatch.start, end: metaMatch.end, text: label }];
    }
    return [];
  }

  // 2. Inline bracket chords [G]（亦支援全形 ［Ｇ］）
  INLINE_BRACKET_REGEX.lastIndex = 0;
  const bracketMatches = [...normLine.matchAll(INLINE_BRACKET_REGEX)];
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

  // 3. Chord line with smart column-alignment preservation (正規化後比對，座標與原字串 1:1 對應)
  CHORD_TOKEN_REGEX.lastIndex = 0;
  const matches = [...normLine.matchAll(CHORD_TOKEN_REGEX)];
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
      let mSpacesEnd = -1;
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
        mSpacesEnd = mSpaces.index + mSpaces[0].length;
      }
      if (debt > 0 && mSpacesEnd < gap.length) {
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

/**
 * 支援段內軟換行 (Shift+Enter -> \u000b / \r / \n)：逐子行獨立計算最小化替換，
 * 避免跨行誤判歌詞/段落標籤或讓對齊空白 debt 跨行污染。
 */
function computeMinimalEditsForParagraph_(fullText, srcKey, dstKey, accPref, mode) {
  if (!/[\r\n\u000b]/.test(fullText)) {
    return computeMinimalEditsForLine_(fullText, srcKey, dstKey, accPref, mode);
  }
  const allEdits = [];
  let offset = 0;
  const parts = fullText.split(/([\r\n\u000b])/);
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    if (i % 2 === 0 && part.trim()) {
      const subEdits = computeMinimalEditsForLine_(part, srcKey, dstKey, accPref, mode);
      for (const ed of subEdits) {
        allEdits.push({
          start: ed.start + offset,
          end: ed.end + offset,
          text: ed.text
        });
      }
    }
    offset += part.length;
  }
  allEdits.sort((a, b) => b.start - a.start || b.end - a.end);
  return allEdits;
}

function getSafeHeadingAttrs_(body, para, headingCache) {
  let headingKey = 'NORMAL';
  let headingEnum = DocumentApp.ParagraphHeading.NORMAL;
  try {
    if (typeof para.getHeading === 'function') {
      const h = para.getHeading();
      if (h) {
        headingEnum = h;
        headingKey = String(h);
      }
    }
  } catch (e) {}
  if (!headingCache[headingKey]) {
    try {
      headingCache[headingKey] = body.getHeadingAttributes(headingEnum) || {};
    } catch (e) {
      headingCache[headingKey] = {};
    }
  }
  return headingCache[headingKey];
}

/**
 * 沿著 Google Docs 樣式繼承鏈（字元 -> 段落 -> 標題樣式 -> Body）解析真實有效樣式，
 * 防止 getAttributes() 回傳 null 時無法還原字級或清除誤繼承灰底。
 */
function captureEffectiveStyle_(textEl, para, idx, headingCache, body, bodyAttrs) {
  const A = DocumentApp.Attribute;
  const paraAttrs = (typeof para.getAttributes === 'function' ? para.getAttributes() : null) || {};
  const headingAttrs = getSafeHeadingAttrs_(body, para, headingCache);

  const rawFontSize = textEl.getFontSize(idx);
  const rawFontFamily = textEl.getFontFamily(idx);
  const rawBg = textEl.getBackgroundColor(idx);
  const rawFg = textEl.getForegroundColor(idx);
  const rawBold = textEl.isBold(idx);
  const rawItalic = textEl.isItalic(idx);
  const rawUnderline = textEl.isUnderline(idx);
  const rawStrike = textEl.isStrikethrough(idx);

  function resolve(rawVal, attrKey) {
    if (rawVal !== undefined && rawVal !== null) return rawVal;
    if (paraAttrs[attrKey] !== undefined && paraAttrs[attrKey] !== null) return paraAttrs[attrKey];
    if (headingAttrs[attrKey] !== undefined && headingAttrs[attrKey] !== null) return headingAttrs[attrKey];
    if (bodyAttrs[attrKey] !== undefined && bodyAttrs[attrKey] !== null) return bodyAttrs[attrKey];
    return null;
  }

  const fontSize = resolve(rawFontSize, A.FONT_SIZE);
  let lineFallbackSize = null;
  if (fontSize === null) {
    const txt = textEl.getText();
    for (let k = idx + 1; k < txt.length; k++) {
      const ch = txt[k];
      if (ch === '\r' || ch === '\n' || ch === '\u000b') break;
      const sz = textEl.getFontSize(k);
      if (sz !== null && sz !== undefined) {
        lineFallbackSize = sz;
        break;
      }
    }
  }

  return {
    fontSize: fontSize,
    rawFontSize: rawFontSize !== undefined ? rawFontSize : null,
    lineFallbackSize: lineFallbackSize,
    fontFamily: resolve(rawFontFamily, A.FONT_FAMILY),
    rawFontFamily: rawFontFamily !== undefined ? rawFontFamily : null,
    backgroundColor: rawBg !== undefined ? rawBg : null,
    foregroundColor: resolve(rawFg, A.FOREGROUND_COLOR),
    rawForegroundColor: rawFg !== undefined ? rawFg : null,
    bold: resolve(rawBold, A.BOLD),
    italic: resolve(rawItalic, A.ITALIC),
    underline: resolve(rawUnderline, A.UNDERLINE),
    strikethrough: resolve(rawStrike, A.STRIKETHROUGH)
  };
}

/**
 * 判斷此替換位置是否位於段落起點、Tab/軟換行之後、或樣式邊界之後：
 * 這些位置若在同一筆 saveAndClose() 交易內直接刪除原首字，DocumentApp 合併 diff 後
 * 會讓新和弦在 Kix 端誤繼承前一段/前一個字元的大字級或灰底色塊！
 */
function needsBoundaryShield_(textEl, fullText, startIdx) {
  if (startIdx <= 0) return true;
  const prevCh = fullText[startIdx - 1];
  if (prevCh === '\t' || prevCh === '\r' || prevCh === '\n' || prevCh === '\u000b') return true;
  if (textEl.getFontSize(startIdx - 1) !== textEl.getFontSize(startIdx)) return true;
  if (textEl.getBackgroundColor(startIdx - 1) !== textEl.getBackgroundColor(startIdx)) return true;
  if (textEl.getFontFamily(startIdx - 1) !== textEl.getFontFamily(startIdx)) return true;
  if (textEl.getForegroundColor(startIdx - 1) !== textEl.getForegroundColor(startIdx)) return true;
  if (textEl.isBold(startIdx - 1) !== textEl.isBold(startIdx)) return true;
  return false;
}

function insertTextSafe_(textEl, offset, str) {
  if (offset >= textEl.getText().length) {
    textEl.appendText(str);
  } else {
    textEl.insertText(offset, str);
  }
}

function rangeNeedsUpdate_(textEl, s, e, getterName, expectedVal) {
  for (let i = s; i <= e; i++) {
    if (textEl[getterName](i) !== expectedVal) return true;
  }
  return false;
}

function enforceStyleOnRange_(textEl, s, e, style) {
  if (!style || s > e || s < 0 || e >= textEl.getText().length) return;
  const A = DocumentApp.Attribute;

  // 1. 背景色（防止誤繼承 Intro / Verse1 / ->V2 / Interlude 灰底 #cccccc）
  if (rangeNeedsUpdate_(textEl, s, e, 'getBackgroundColor', style.backgroundColor)) {
    textEl.setBackgroundColor(s, e, style.backgroundColor);
  }

  // 2. 字級（防止誤繼承上行歌詞 17pt 或段落標籤 11pt）
  if (typeof style.fontSize === 'number' && style.fontSize > 0) {
    if (rangeNeedsUpdate_(textEl, s, e, 'getFontSize', style.fontSize)) {
      try {
        textEl.setFontSize(s, e, style.fontSize);
      } catch (err) {
        textEl.setFontSize(s, e, Math.round(style.fontSize));
      }
    }
  } else {
    let hasExplicitSize = false;
    for (let i = s; i <= e; i++) {
      if (textEl.getFontSize(i) !== null) {
        hasExplicitSize = true;
        break;
      }
    }
    if (hasExplicitSize) {
      const clearSize = {};
      clearSize[A.FONT_SIZE] = null;
      textEl.setAttributes(s, e, clearSize);
      let stillWrong = false;
      for (let i = s; i <= e; i++) {
        if (textEl.getFontSize(i) !== null) {
          stillWrong = true;
          break;
        }
      }
      if (stillWrong) {
        const fb = (typeof style.lineFallbackSize === 'number' && style.lineFallbackSize > 0)
          ? Math.round(style.lineFallbackSize)
          : 11;
        textEl.setFontSize(s, e, fb);
      }
    }
  }

  // 3. 字型
  if (style.fontFamily !== null && style.fontFamily !== undefined) {
    let famDrifted = false;
    for (let i = s; i <= e; i++) {
      const actual = textEl.getFontFamily(i);
      if (style.rawFontFamily !== null) {
        if (actual !== style.rawFontFamily) { famDrifted = true; break; }
      } else if (actual !== null && actual !== style.fontFamily) {
        famDrifted = true;
        break;
      }
    }
    if (famDrifted) {
      try {
        textEl.setFontFamily(s, e, style.fontFamily);
      } catch (err) {
        // 忽略非 Google Fonts 內建名稱的字型設定例外，避免中斷轉調
      }
    }
  }

  // 4. 文字顏色
  if (style.rawForegroundColor !== null) {
    if (rangeNeedsUpdate_(textEl, s, e, 'getForegroundColor', style.rawForegroundColor)) {
      textEl.setForegroundColor(s, e, style.rawForegroundColor);
    }
  } else {
    let fgDrifted = false;
    for (let i = s; i <= e; i++) {
      const actual = textEl.getForegroundColor(i);
      if (actual !== null && actual !== style.foregroundColor) {
        fgDrifted = true;
        break;
      }
    }
    if (fgDrifted) {
      textEl.setForegroundColor(s, e, style.foregroundColor);
    }
  }

  // 5. 粗體 / 斜體 / 底線 / 刪除線
  const boolChecks = [
    ['isBold', 'setBold', style.bold],
    ['isItalic', 'setItalic', style.italic],
    ['isUnderline', 'setUnderline', style.underline],
    ['isStrikethrough', 'setStrikethrough', style.strikethrough]
  ];
  for (const [getter, setter, expected] of boolChecks) {
    const targetBool = expected === true;
    let drifted = false;
    for (let i = s; i <= e; i++) {
      const actual = textEl[getter](i);
      if (expected === true && actual !== true) {
        drifted = true;
        break;
      }
      if (expected !== true && actual === true) {
        drifted = true;
        break;
      }
    }
    if (drifted) {
      textEl[setter](s, e, targetBool);
    }
  }
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
  const docName = doc.getName();
  const paragraphs = [];
  collectParagraphsFromContainer_(doc.getBody(), paragraphs);
  const texts = paragraphs.map(p => p.getText());
  return {
    title: docName,
    url: doc.getUrl(),
    detectedKey: detectKeyFromParagraphs_(texts, docName)
  };
}

/**
 * 執行 Google Docs 原地轉調（可選擇先建立複本，或直接編輯指定文件）
 * 採用「兩階段防護字元提交 (Two-Phase Guard-Char Commit) + 樣式二次校驗」：
 * - 第一階段：對位於段首 / Tab / 軟換行 / 樣式邊界的和弦，先在原首字之後插入新和弦並保留原首字作為「樣式護盾」再 saveAndClose()，
 *   徹底阻斷 DocumentApp 合併 diff 時在段首直接替換而誤繼承上一段大字級或灰底色塊的問題。
 * - 第二階段：重新開啟文件，僅刪除 1 字元護盾，並逐字比對還原原始字級、字型、灰底與顏色。
 */
function runTransposeInGoogleDocs(params) {
  const targetKey = params.targetKey || 'F';
  const accPref = params.accPref || 'auto';
  const mode = params.mode || 'chord';
  const createCopy = params.createCopy !== false;

  let doc;
  let origDocName = '';
  if (params.docUrlOrId && params.docUrlOrId.trim()) {
    const m = params.docUrlOrId.match(/\/d\/([a-zA-Z0-9_-]{15,})/);
    const fileId = m ? m[1] : params.docUrlOrId.trim();
    if (createCopy) {
      const origFile = DriveApp.getFileById(fileId);
      origDocName = origFile.getName();
      const copyTitle = buildCopyTitle_(origDocName, targetKey);
      const copyFile = origFile.makeCopy(copyTitle);
      doc = DocumentApp.openById(copyFile.getId());
    } else {
      doc = DocumentApp.openById(fileId);
      origDocName = doc.getName();
    }
  } else {
    const activeDoc = DocumentApp.getActiveDocument();
    if (createCopy) {
      const origFile = DriveApp.getFileById(activeDoc.getId());
      origDocName = origFile.getName();
      const copyTitle = buildCopyTitle_(origDocName, targetKey);
      const copyFile = origFile.makeCopy(copyTitle);
      doc = DocumentApp.openById(copyFile.getId());
    } else {
      doc = activeDoc;
      origDocName = doc.getName();
    }
  }

  const docId = doc.getId();
  const docTitle = doc.getName();
  const body = doc.getBody();
  const bodyAttrs = body.getAttributes() || {};
  const headingCache = {};

  const paragraphs = [];
  collectParagraphsFromContainer_(body, paragraphs);
  const texts = paragraphs.map(p => p.getText());
  const srcKey = (params.sourceKey && params.sourceKey !== 'auto') ? params.sourceKey : detectKeyFromParagraphs_(texts, origDocName);

  let modifiedParagraphs = 0;
  const deferredGuards = [];

  for (let paraIdx = 0; paraIdx < paragraphs.length; paraIdx++) {
    const para = paragraphs[paraIdx];
    const textEl = para.editAsText();
    const fullText = textEl.getText();
    if (!fullText.trim()) continue;
    const edits = computeMinimalEditsForParagraph_(fullText, srcKey, targetKey, accPref, mode);
    if (edits.length === 0) continue;

    const paraGuards = [];

    for (const ed of edits) {
      let delta = 0;
      if (ed.text.length > 0 && ed.end > ed.start) {
        const origStyle = captureEffectiveStyle_(textEl, para, ed.start, headingCache, body, bodyAttrs);
        const useShield = needsBoundaryShield_(textEl, fullText, ed.start);

        insertTextSafe_(textEl, ed.start + 1, ed.text);
        if (ed.end - ed.start > 1) {
          textEl.deleteText(ed.start + 1 + ed.text.length, ed.end + ed.text.length - 1);
        }

        if (useShield) {
          // 第一階段先保留 ed.start 原首字作為樣式護盾，避免 Kix 合併成段首替換而繼承上一段樣式
          delta = 1 + ed.text.length - (ed.end - ed.start);
          for (const g of paraGuards) {
            g.guardIdx += delta;
          }
          paraGuards.push({
            guardIdx: ed.start,
            newLen: ed.text.length,
            guardChar: fullText[ed.start],
            origStyle: origStyle
          });
        } else {
          textEl.deleteText(ed.start, ed.start);
          enforceStyleOnRange_(textEl, ed.start, ed.start + ed.text.length - 1, origStyle);
          delta = ed.text.length - (ed.end - ed.start);
          for (const g of paraGuards) {
            g.guardIdx += delta;
          }
        }
      } else if (ed.text.length > 0) {
        // 純補空白 (debt < 0)
        const curLen = textEl.getText().length;
        const refIdx = ed.start < curLen ? ed.start : Math.max(0, ed.start - 1);
        const refStyle = curLen > 0 ? captureEffectiveStyle_(textEl, para, refIdx, headingCache, body, bodyAttrs) : null;
        insertTextSafe_(textEl, ed.start, ed.text);
        if (refStyle) {
          enforceStyleOnRange_(textEl, ed.start, ed.start + ed.text.length - 1, refStyle);
        }
        delta = ed.text.length;
        for (const g of paraGuards) {
          g.guardIdx += delta;
        }
      } else if (ed.end > ed.start) {
        // 純刪除多餘空白 (debt > 0)
        textEl.deleteText(ed.start, ed.end - 1);
        delta = -(ed.end - ed.start);
        for (const g of paraGuards) {
          g.guardIdx += delta;
        }
      }
    }

    if (paraGuards.length > 0) {
      deferredGuards.push({ paraIdx: paraIdx, guards: paraGuards });
    }
    modifiedParagraphs++;
  }

  doc.saveAndClose();

  // 第二階段：重新載入已寫入的文件，移除 1 字元樣式護盾並強制校驗字級與背景色
  if (deferredGuards.length > 0) {
    const doc2 = DocumentApp.openById(docId);
    const paragraphs2 = [];
    collectParagraphsFromContainer_(doc2.getBody(), paragraphs2);

    for (const item of deferredGuards) {
      if (item.paraIdx >= paragraphs2.length) continue;
      const textEl2 = paragraphs2[item.paraIdx].editAsText();
      // item.guards 已依 guardIdx 由大到小排列，刪除較後方護盾不會影響前方索引
      for (const g of item.guards) {
        const curText = textEl2.getText();
        if (g.guardIdx >= 0 && g.guardIdx < curText.length && curText.charAt(g.guardIdx) === g.guardChar) {
          textEl2.deleteText(g.guardIdx, g.guardIdx);
          enforceStyleOnRange_(textEl2, g.guardIdx, g.guardIdx + g.newLen - 1, g.origStyle);
        }
      }
    }
    doc2.saveAndClose();
  }

  return {
    ok: true,
    docId: docId,
    title: docTitle,
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
