# 排班頁導航 400 Bad Request — 排查紀錄

## 症狀

排班頁 topbar 點「上個月 / 下個月 / 今天」→ Google Drive 錯誤頁
```
很抱歉，目前無法開啟這個檔案。
請檢查網址並再試一次。
```

DevTools Network 顯示：
```
GET https://script.google.com/macros/s/AKfycbzEj6KAWlidiv4FjjpZhp9zGRHghjk... 
Status: 400 (Bad Request)
```

LINE OAuth callback 直接渲染排班頁正常工作（含 staff mode 限制：載入預設 / 批次確認消失、員工下拉鎖定）。

## 部署資訊（2026-05-19 19:07）

- **部署版本**：v32 - prefer live URL
- **部署作業 ID**：`AKfycbzEj6KAWlidiv4FjjpZhp9zGRHghjkQMzpFnLuQ55SVWnsyiZl-lmH2HU_FzmHhtgJY`
- **網址**：`https://script.google.com/macros/s/AKfycbzEj6KAWlidiv4FjjpZhp9zGRHghjkQMzpFnLuQ55SVWnsyiZl-lmH2HU_FzmHhtgJY/exec`

## 已嘗試的修法

1. **window.top.location.href** — 被 iframe sandbox 擋（user activation 錯誤）
2. **navigateTop() 動態建 `<a target="_top">` 後 .click()** — 程式式 click 不算 user activation，仍失敗
3. **改成真實 `<a target="_top">` 連結**（server 預先建好 URL）— 仍 400
4. **`getWebAppUrl_()` 優先用 `ScriptApp.getService().getUrl()`** — 仍 400

## 觀察

- callback URL 落地時 schedule HTML 渲染**沒問題**（直接 server 端 render，沒導航）
- 任何 top-level 導航到同個 deployment URL 都返回 400
- Apps Script Web App 對某些 query params 組合可能有過濾？

## 下次可嘗試

### A. 確認 URL 完整內容
打開 DevTools → Network → 點失敗請求 → Headers → 截「Request URL」**完整 URL**（含所有 params）。
特別檢查：
- sid 內容是否異常
- year / month 數字是否正確
- 是否有空格或控制字元

### B. 直接貼 URL 測試
1. Apps Script 編輯器跑 `ShowWebAppUrl` 拿到 production /exec URL
2. 手動加 `?action=schedule&year=2026&month=6` 貼到網址列 Enter
3. 看是否 400
   - 200 → 我們動態建的 URL 有特殊字元
   - 400 → Apps Script 對 query params 過濾異常

### C. 對照 home.html 那條 LINE 登入連結
home 頁的「使用 LINE 登入」`<a target="_top" href="...">` 點下去能跳到 LINE，運作正常。但那個 URL 是 line.me，**不是 script.google.com**。

可能 Apps Script 內部跳轉到同 host 的 /macros/s URL 有特別限制。

### D. 改用 google.script.run 模式刷月份
不用 navigation，直接呼 `google.script.run.getScheduleViewData(year, month)` 後 in-place 重繪。和 refreshData 同模式。topbar 上 / 下個月 / 今天改成普通 button + onclick → call function。
這方式根本不用 top-level navigation。

**建議優先試 D**，繞過 navigation 問題從根本。

## 待做（依序）

- [x] #1 換班申請流程
- [/] #2 員工自排權限（含 auth）— **staff UI 限制 OK，但月份切換 400 待解**
- [ ] #3 請假申請流程
- [ ] #4 手機響應式

## 已知工作正常的功能

- LINE 登入後自動進排班頁
- 排班 CRUD（含樂觀更新）
- 換班 / 讓班完整流程（待對方同意 → 待店長核可 → 已生效）
- 載入預設 / 批次確認 / 狀態切換 / 編輯
- 拜土地公自動偵測（用 Intl 農曆）
- Staff mode UI 限制（按鈕隱藏、員工下拉鎖自己）
- 員工本月統計表（含休、班數等）
