# 日常帳本記錄器 — 部署指南

## 架構說明

```
GitHub Pages (index.html)  →  POST  →  Google Apps Script  →  Google Sheets
```

---

## Step 1：上傳到 GitHub Pages

1. 在 GitHub 建立新 repo（例如 `ledger`）
2. 把 `index.html` 推上去
3. Settings → Pages → Branch: main, / (root) → Save
4. 幾分鐘後可用 `https://你的名字.github.io/ledger/` 開啟

---

## Step 2：設定 Google Apps Script

1. 開啟你的 2026 帳本 Google Sheets
2. 上方選單：**擴充功能 → Apps Script**
3. 把 `Code.gs` 的內容**全部貼上**（取代原有內容）
4. 存檔（Ctrl+S）

### 部署

1. 點右上角 **部署 → 新增部署作業**
2. 類型選 **網頁應用程式**
3. 設定：
   - 描述：`帳本記錄器`
   - 以誰的身分執行：**我自己**
   - 誰可以存取：**所有人**（Anyone）
4. 按 **部署** → 複製「網頁應用程式網址」

---

## Step 3：填入網址

1. 在 `https://你的名字.github.io/ledger/` 開啟記錄器
2. 點下方「⚙️ Google Sheets 連線設定」
3. 貼上 GAS 網址
4. Sheet 名稱填對（例如 `2026`）
5. 按**儲存設定**

---

## 注意事項

- **每次修改 Code.gs 要重新部署**（版本號會跳）
- GAS 免費方案每天有 API 呼叫上限，個人日常使用完全足夠
- 資料是 append（新增列），不會覆蓋舊資料
- 錢包餘額欄位可以空著不填

---

## 欄位對應表

| 前端欄位 | Sheets 欄位 |
|---|---|
| 日期 | A（Excel 序號格式） |
| 備忘 | B |
| 早餐 / 明細 | C / D |
| 午餐 / 明細 | E / F |
| 晚餐 / 明細 | G / H |
| 下午茶 / 明細 | I / J |
| 飲料 / 明細 | K / L |
| 消夜 / 明細 | M / N |
| 汽車油錢 | O |
| 機車油錢 | Q |
| 交通 1 / 2 | S / U |
| 停車費 1 / 2 | W / Y |
| 美安消費 1~3 | AA / AC / AE |
| 雜項 1~3 | AG / AI / AK |
| 錢包餘額 | AM |

---

## 問題排除

| 問題 | 解法 |
|---|---|
| 上傳沒反應 | 確認 GAS 網址正確、重新部署 |
| 出現 CORS 錯誤 | 正常！`no-cors` 模式下看不到回應，但資料有進去 |
| Sheet 名稱找不到 | GAS 會自動建立新分頁 |
| 日期格式怪 | A 欄格式改成數字，數值是正常 Excel 序號 |
