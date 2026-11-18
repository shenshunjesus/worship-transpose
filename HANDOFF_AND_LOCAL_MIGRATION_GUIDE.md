# 敬拜歌譜無損轉調工具 · 完整專案進度整理與本地接力手冊

> 本手冊彙整了本次在雲端開發的所有核心演算法、檔案清單、測試結果與架構設計，讓您在本機電腦上的專案能 100% 無縫接力、立即運行與維護。

---

## 📌 一、專案目標與解決的核心問題

一般的和弦轉調工具（如 Ultimate Guitar 或各大吉他譜網站）只能處理純文字，一旦遇到教會敬拜團正式的 Word 歌譜（`.docx`），複製貼上後會遭遇：
1. **排版崩潰**：段落標籤（`Intro`, `Verse`, `Chorus`, `Bridge`）的灰底標記（`#cccccc`）消失。
2. **字型跑掉**：繁體微軟正黑體、Roboto 等指定字型被還原為預設字型。
3. **字級打亂**：標題（15pt）、和弦（13pt）、歌詞（17pt）全部混在一起。
4. **升降記號錯誤**：轉到 `Eb` 調卻出現 `A#`、`D#` 等非調內和弦；或轉到 `G` 調出現 `Gb`。

**本專案成功解決了上述所有問題**，實現了 **100% 格式無損的 `.docx` 與純文字轉調引擎**。

---

## 📂 二、完整檔案架構清單

雲端已完成的所有檔案均位於 `/home/antigravity/chord-transposer`，已打包為 `/tmp/worship-transpose.zip`：

```
worship-transpose/
├── index.html            # 100% 純前端網頁 App (單一靜態檔，含內建 ZIP/XML 引擎，零外部相依)
├── transpose_docx.py     # Python 核心無損 Word XML 轉調 CLI (標準庫實現，零需 pip 安裝)
├── gdocs_transpose.py    # Google Drive 整合自動化批次轉調腳本
├── README.md             # 完整專案說明文件 (包含使用範例與功能清單)
├── .gitignore            # Git 忽略檔案設定 (__pycache__ 等)
├── samples/              # 原檔範本庫
│   └── 坐在寶座上的那位 D.docx   # 原始 D 調 Word 歌譜原檔
└── output/               # 測試驗證產出的轉調成品 (.docx)
    ├── 坐在寶座上的那位 C.docx      (D -> C)
    ├── 坐在寶座上的那位 Eb.docx     (D -> Eb)
    ├── 坐在寶座上的那位 E.docx      (D -> E)
    ├── 坐在寶座上的那位 F.docx      (D -> F)
    ├── 坐在寶座上的那位 G.docx      (D -> G)
    └── 坐在寶座上的那位 級數譜.docx   (D -> C 級數譜 1-7)
```

---

## 🧠 三、核心演算法與技術細節

### 1. 無損 Word XML 操作機制 (`transpose_docx.py`)
- **不透過 docx 套件重排版**：一般的 `python-docx` 寫入時會重寫整個 XML 樣式表，造成微軟正黑體與灰底標記遺失。
- **直接操作 `word/document.xml` 原始位元流**：
  1. 使用 Python 內建 `zipfile` 讀取 `.docx`。
  2. 使用正則表達式精確鎖定 `<w:p>`（段落）與內部的 `<w:t>`（文字節點）。
  3. 保留所有的 `<w:pPr>`（段落樣式）、`<w:rPr>`（文字樣式，如 `<w:shd w:fill="cccccc"/>` 灰底、`<w:sz w:val="26"/>` 字級、`<w:rFonts>` 字型）。
  4. 只替換 `<w:t>` 內部的純文字，完全不更動 XML 結構與任何標籤屬性。
  5. 重新以 ZIP Deflate 打包，產出 100% 相容的 `.docx`。

### 2. 智慧和弦識別正規表達式
```python
CHORD_TOKEN_RE = re.compile(
    r'(?<![A-Za-z0-9])'
    r'([A-G][#b♯♭]?)'                                              # Root (根音)
    r'((?:maj|min|m|M|dim|aug|sus|add|no|[0-9]|[\+\^\(\)])*)'      # Quality (和弦性質)
    r'(?:/([A-G][#b♯♭]?))?'                                       # Slash Bass (轉位/低音)
    r'(?![A-Za-z0-9])'
)
```
- 支援斜線和弦（如 `D/F#` ➔ `Eb/G`）。
- 支援小節線拍子標記（如 `G - Bm - ｜D - - -`）與自由空白對齊排版。
- 支援歌詞內嵌和弦標籤（如 `[G]我想認識[Bm]祢`）。

### 3. 五度圈智慧升降記號判斷（Enharmonic Rule）
- **自然降號調性**：`F`, `Bb`, `Eb`, `Ab`, `Db`, `Gb` 及其實音小調 `Dm`, `Gm`, `Cm`, `Fm`, `Bbm`, `Ebm` 自動使用降號記號（`b`）。
- **自然升號調性**：`G`, `D`, `A`, `E`, `B`, `F#` 自動使用升號記號（`#`）。
- 同時支援手動強制切換（`sharp` / `flat`）。

### 4. Nashville 級數譜 (1-7) 換算
- 根據調性大音階（Major Scale）半音距離計算：
  - `0`: 1（主音）
  - `2`: 2（或 2m）
  - `4`: 3（或 3m）
  - `5`: 4
  - `7`: 5
  - `9`: 6（或 6m）
  - `11`: 7（或 7dim）
- 範例：《坐在寶座上的那位》Verse1:
  - D 調：`D - - - ｜- - - - | Bm - - - | - - - -`
  - 級數譜：`1 - - - ｜- - - - | 6m - - - | - - - -`
  - 副歌斜線和弦：`Em` ➔ `2m`、`D/F#` ➔ `1/3`

### 5. 木吉他 Capo 移調夾自動換算
- 自動比對吉他手最容易按弦的 6 個開放把位調性（`C`, `G`, `D`, `A`, `E`, `F`）。
- 計算公式：`Capo = (TargetSemitone - PlaySemitone) % 12`。
- 範例：若目標為 `Eb` 調（半音值 3）：
  - 夾 Capo 1 彈 D 調指法（`3 - 2 = 1`）
  - 夾 Capo 3 彈 C 調指法（`3 - 0 = 3`）

---

## 💻 四、本地端無縫接力執行方式

在您本機的電腦專案中，您完全不受任何雲端權限限制，可直接按以下步驟接力：

### 1. 本地 CLI 工具執行
零依賴，只要有 Python 3 即可直接執行：

```bash
# 轉調為任意目標調性 (例: Eb 調)
python3 transpose_docx.py "坐在寶座上的那位 D.docx" --to Eb

# 轉為 Nashville 級數譜 (1-7)
python3 transpose_docx.py "坐在寶座上的那位 D.docx" --to C --mode number

# 指定輸出檔名
python3 transpose_docx.py "坐在寶座上的那位 D.docx" --to G --output "坐在寶座上的那位_G調.docx"
```

### 2. 本地 Web App 執行
`index.html` 是一個完全自包含的網頁應用程式：
- 本機直接雙擊 `index.html` 在 Chrome / Safari / Edge 打開即可使用。
- 支援拖曳 `.docx` 或 `.txt`、即時轉調預覽、點選 Capo、一鍵下載原排版 `.docx`、列印 A4/存 PDF。

### 3. 將程式碼推上 GitHub（從本機執行 100% 順暢）
在您本機電腦的 Terminal 中：

```bash
# 進入您的本機專案目錄
cd <您的本地專案資料夾>

# 複製此專案的檔案 (index.html, transpose_docx.py, README.md 等)
# 提交並推送到 GitHub
git add .
git commit -m "feat: complete worship chord transposer app"
git push -u origin main
```

---

## 🚀 五、接力後的建議延伸方向

1. **多首歌曲曲庫打包工具**：
   - 擴充 `transpose_docx.py`，支援傳入整個資料夾（例如一次把 20 首敬拜讚美詩歌的 `.docx` 同步產出全套 `C`, `D`, `E`, `G`, `A` 調）。
2. **PDF 直接轉調預覽**：
   - 在本地端可加入 `docx2pdf` 或 LibreOffice headless 轉換，讓網頁端除了下載 `.docx` 之外也能直接輸出渲染好的 PDF 歌譜。
3. **即時簡譜／伴奏譜對照**：
   - 在網頁端加入數字簡譜模式或吉他指法和弦圖（Chord Diagrams）SVG 提示。
