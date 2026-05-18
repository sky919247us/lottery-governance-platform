# 彩券行公司治理平台

模組化內部系統。目前涵蓋：

- **Core** — 員工 / 門市 / 權限 / LINE 綁定
- **Attendance**（P1 規劃中）— 排班 + 打卡 + 出勤稽核

完整規劃請見 [docs/規劃.md](docs/規劃.md)。

## 模組結構

```
core/         Core 模組（Google Apps Script + Sheets）
attendance/   Attendance 模組（P1 開工）
docs/         規劃文件
```

## 開發環境

- Node.js LTS
- Google Apps Script + [clasp](https://github.com/google/clasp)
- Git

## 部署流程

```powershell
cd core
clasp push
```

從 Apps Script Editor 跑一次 `Bootstrap()` 函式 → 自動建立全部 8 張 Sheet + 預設資料。

## 後端 Sheets

- **Core Sheet**：`[彩券行] Core`（敏感資料以保護範圍限制只 admin 開啟）
- **Attendance Sheet**：`[彩券行] Attendance 2026`（每年新開，按年分頁）
