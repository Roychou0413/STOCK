# 專案說明

- `index.html`：手機用的倉庫盤點掃描網頁（單一檔案）。
- `apps-script/Code.gs`：Google Apps Script 後端。每人密碼登入（「使用者」分頁），所有請求都帶 `code` 參數；修改紀錄寫在「紀錄」分頁。
- 前端需相容舊版後端（不帶 action 的存活測試回應沒有 `version` 時視為 v1）。
- 回覆一律使用繁體中文。

## 每次改動都要提供預覽頁

每次修改完成後：
1. 先在新分支 commit 並 push（不要直接推 `main`）。
2. 回覆中必須附上該分支的預覽網址，讓使用者用手機實測：
   `https://raw.githack.com/Roychou0413/STOCK/<分支名稱>/index.html`
   （HTTPS，可開相機；push 後約 1～5 分鐘生效，若是舊版請強制重新整理）
3. 附上本次要測試的項目清單。
