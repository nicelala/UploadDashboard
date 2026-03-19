# Changelog

## [2.0.0] - 2026-03-xx
### Added
- Error code 機制：以 `error_keywords.csv`（兩欄：error_code, keyword）外部化「訊息 → 代碼」對應，支援「規則 → 重新載入規則」熱載。
- UI「規則」選單：開啟 / 重新載入 / 規則說明。
- 未命中任何規則時，指派 `E000` 並跳出提醒對話框（提供一鍵開啟 `error_keywords.csv`）。

### Changed
- **事件 CSV 表頭改為**  
  `event_id,created_at,status,trigger,message,error_code`
- 訊息前綴統一（`manual:` / `scheduler:`），利於規則比對與查詢。

### Notes
- 舊日誌不回填 `error_code`；新事件才會套用規則。
- 建議以 One Directory + Window Based 打包；不隨發佈包附 `error_keywords.csv`（由使用端自行維護）。

---

## [1.0.0] - 2026-02-xx
### Initial
- 兩階段（download/upload）執行與日誌。
- 事件 CSV 表頭：  
  `created_at,status,trigger,message,event_id`