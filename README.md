# 敬拜歌譜無損轉調工具 · Worship Chord & .docx Transposer

[![Live Demo](https://img.shields.io/badge/線上工具-點擊開啟-brightgreen?style=for-the-badge&logo=googlechrome&logoColor=white)](https://shenshunjesus.github.io/worship-transpose/)
[![Zero Upload](https://img.shields.io/badge/隱私安全-100%25%20瀏覽器本地處理-blue?style=for-the-badge&logo=lock&logoColor=white)](https://shenshunjesus.github.io/worship-transpose/)
[![No Install](https://img.shields.io/badge/免安裝-支援手機與電腦-orange?style=for-the-badge)](https://shenshunjesus.github.io/worship-transpose/)

> **專為敬拜團設計的歌譜與和弦轉調工具。**  
> 支援將 Google Docs / Word（`.docx`）歌譜無損轉調為任意目標調性或 Nashville 級數譜（1-7），**100% 保留原始字體大小、微軟正黑體、灰底段落標籤（Intro / Verse / Chorus）與排版格式**。

🔗 **線上直接使用**：[https://shenshunjesus.github.io/worship-transpose/](https://shenshunjesus.github.io/worship-transpose/)

---

## ✨ 核心特色

1. **🔒 100% 瀏覽器本地處理（零上傳、免授權）**
   - 內建純 JavaScript 的 ZIP 與 Word XML (`word/document.xml`) 解析引擎。
   - 所有 `.docx` 檔案均在使用者電腦或手機的記憶體中即時處理，**絕不傳送至任何後端伺服器**，連網路斷線都能離線使用。
   - 毋需要求使用者授權讀取 Google Drive，任何人都能安心拖入檔案。

2. **📄 原排版無損轉調（Lossless .docx Formatting）**
   - 精確修改 Word XML 文字節點，不更動樣式結構。
   - 100% 保留微軟正黑體、字級大小（標題 15pt、和弦 13pt、歌詞 17pt）與灰底色塊標籤（`#cccccc`）。
   - 自動將譜頭調性資訊（如 `Upper Room/活水/D`）同步更新為目標調（如 `Upper Room/活水/Eb`）。

3. **♯/♭ 智慧升降記號判斷**
   - 依目標調性自動遵循五度圈慣例（F, Bb, Eb, Ab 調自動用 `b`，G, D, A, E 調自動用 `#`）。
   - 亦支援手動強制全升（`#`）或全降（`b`）。

4. **🎼 Nashville 級數譜 (1-7) 與雙行對照**
   - 一鍵轉為級數譜（例如 `4 - 6m - ｜1 - - -`），適合樂手彈奏。
   - 支援「標準和弦」、「級數譜」、「和弦 + 級數」三種顯示模式。

5. **🎸 吉他手 Capo 移調夾換算**
   - 自動計算最適合木吉他手按弦的 Capo 格數（例如實音 `Eb` 調提示 `Capo 1 彈 D 調指法`、`Capo 3 彈 C 調指法`）。
   - 點選即可直接預覽並匯出吉他指法專用譜。

6. **📦 批次匯出常用調**
   - 一鍵打包匯出 C, D, Eb, E, F, G, A, Bb 等常用敬拜調性及級數譜，快速建立樂團曲庫。

---

## 🚀 使用方式

### 方法一：直接使用 GitHub Pages 網頁版
打開 [https://shenshunjesus.github.io/worship-transpose/](https://shenshunjesus.github.io/worship-transpose/)，直接拖曳 `.docx` 檔案進去即可即時轉調與下載。

### 方法二：終端機 CLI 工具 (`transpose_docx.py`)
若想在本地指令列批次處理：

```bash
# 轉調為 Eb 調
python3 transpose_docx.py "坐在寶座上的那位 D.docx" --to Eb

# 轉為 Nashville 級數譜
python3 transpose_docx.py "坐在寶座上的那位 D.docx" --to C --mode number

# 強制使用降記號 (b)
python3 transpose_docx.py "坐在寶座上的那位 D.docx" --to Db --accidental flat
```

### 方法三：Google Drive 直接轉調工具 (`gdocs_transpose.py`)
搭配 `gdrive` CLI，可直接指定雲端檔案 ID 進行轉調：

```bash
python3 gdocs_transpose.py "<GOOGLE_DRIVE_FILE_ID>" --to "C,Eb,F,G,number"
```

---

## 📂 專案結構

```
worship-transpose/
├── index.html            # 單一靜態網頁 App (100% 純前端，無外部相依)
├── transpose_docx.py     # Python 核心無損 Word XML 轉調引擎 (支援 CLI)
├── gdocs_transpose.py    # Google Drive 整合自動化腳本
├── samples/              # 範本歌譜資料夾
│   └── 坐在寶座上的那位 D.docx
└── README.md             # 說明文件
```

---

## 📜 授權與宣告
本專案為開源工具，歡迎教會敬拜團自由使用與分享！
