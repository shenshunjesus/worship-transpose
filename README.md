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

### 方法一：直接在 Google Docs 使用 Apps Script 原生轉調（推薦・免開網頁、100% 零破版）

不需要透過 `.docx` 轉檔，也不用打開任何外部網頁，直接把 [`google_docs_addon.gs`](./google_docs_addon.gs) 貼到 Google Docs 的 Apps Script，就能在原文件裡**就地替換和弦**（完整保留微軟正黑體、字級、Tab 定位點、縮排、灰底色塊與表格），且**預設會先自動建立複本，絕不更動原始歌譜**。

#### 📋 5 步驟設定教學

1. **複製腳本程式碼**
   - 點開本專案中的 [`google_docs_addon.gs`](./google_docs_addon.gs)。
   - 點擊程式碼右上角的 **複製按鈕（Copy raw file）**（或點 `Raw` 後全選 `Ctrl+A` → `Ctrl+C`）。

2. **在 Google Docs 開啟 Apps Script**
   - 打開你要轉調的 Google Docs 歌譜。
   - 點擊上方選單列的 **`擴充功能`** → **`Apps Script`**，會開啟新的腳本編輯器分頁。

3. **貼上程式碼並儲存**
   - 將編輯器內預設的 `function myFunction() { }` **全部刪除**。
   - 貼上剛才複製的 `google_docs_addon.gs` 完整內容。
   - 按下上方工具列的 **儲存圖示**（或按 `Ctrl + S` / Mac 為 `Cmd + S`）。

4. **首次執行並完成 Google 授權**
   - 確認工具列函式選單為 **`onOpen`**，點擊 **`執行`**（▶）。
   - 跳出「需要授權」視窗時，點 **`審查權限`** → 選擇你的 Google 帳號。
   - 若出現 **「Google 尚未驗證這個應用程式」** 提示（因腳本是你自行建立的個人專案），請點左下角 **`進階`** → **`前往「未命名專案」（不安全）`** → **`允許`**。
   - > 🔒 **安全說明**：此腳本完全在你自己的 Google 帳號下執行，不會將任何檔案或權限傳送到外部伺服器。

5. **回到歌譜開始一鍵轉調**
   - 切回原本的 Google Docs 歌譜分頁並 **重新整理頁面**（約等 3～5 秒）。
   - 上方選單列會多出 **`🎵 敬拜歌譜轉調`** → 點選 **`開啟無損轉調面板`**，右側即會出現轉調操作面板：
     - **原調**：預設自動偵測，也可手動指定。
     - **目標調性**：點選欲轉換的調（`C` ～ `B`）。
     - **譜面模式**：支援 **標準和弦譜**、**級數譜 (1-7)**、**和弦 + 級數**。
     - **升降記號**：支援依調性自動、強制升記號（`♯`）或強制降記號（`♭`）。
     - **安全複本**：預設勾選「先在我的雲端硬碟建立『複本』再轉調」（推薦保留勾選，避免改到原檔）。
   - 點擊 **`開始原生轉調`**，完成後即可直接點按鈕 **開啟轉調後的 Google Docs** 或 **直接下載乾淨 PDF**！

<details>
<summary>💡 進階用法：不想每份歌譜都貼一次腳本？部署成自己的「網址轉調工具」</summary>

`google_docs_addon.gs` 同時支援網頁應用程式模式（`doGet`），只需部署一次，之後貼上任何歌譜連結就能轉調：

1. 在剛才存好的 Apps Script 編輯器右上角，點 **`部署`** → **`新增部署作業`**。
2. 點齒輪圖示選 **`網頁應用程式`**。
3. **執行身分**選 `我`，**誰可以存取**選 `僅限我自己`（或依團隊需求調整），點擊 **`部署`**。
4. 複製產生的 **網頁應用程式網址** 並加入書籤。以後打開該網址，貼上任一份 Google Docs 歌譜連結即可直接建立轉調複本與下載 PDF！

</details>

---

### 方法二：直接使用 GitHub Pages 網頁版（`.docx` 本地處理）
打開 [https://shenshunjesus.github.io/worship-transpose/](https://shenshunjesus.github.io/worship-transpose/)，直接拖曳 `.docx` 檔案進去即可即時轉調與下載。

### 方法三：終端機 CLI 工具 (`transpose_docx.py`)
若想在本地指令列批次處理：

```bash
# 轉調為 Eb 調
python3 transpose_docx.py "坐在寶座上的那位 D.docx" --to Eb

# 轉為 Nashville 級數譜
python3 transpose_docx.py "坐在寶座上的那位 D.docx" --to C --mode number

# 強制使用降記號 (b)
python3 transpose_docx.py "坐在寶座上的那位 D.docx" --to Db --accidental flat
```

### 方法四：Google Drive 直接轉調工具 (`gdocs_transpose.py`)
搭配 `gdrive` CLI，可直接指定雲端檔案 ID 進行轉調：

```bash
python3 gdocs_transpose.py "<GOOGLE_DRIVE_FILE_ID>" --to "C,Eb,F,G,number"
```

---

## 📂 專案結構

```
worship-transpose/
├── google_docs_addon.gs  # Google Docs 原生 Apps Script 無損轉調腳本（免開網頁直接套用）
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
