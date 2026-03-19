# 事件資料 Schema（v1 vs v2）

## v1.0.0
- 表頭：`created_at,status,trigger,message,event_id`
- `status`: success/fail
- `trigger`: manual/scheduler/unknown
- `message`: `{trigger}: {stage}: ...` 色彩不一（歷史訊息可能不完全一致）

## v2.0.0
- 表頭：`event_id,created_at,status,trigger,message,error_code`
- `error_code`:
  - `success` 留空
  - `fail`：先比對 `error_keywords.csv`（由上而下、大小寫不敏感、包含即命中）→ 未命中套用內建保底 → 仍無則 `E000`
- 內建保底碼（節錄）
  - E001 Timeout、E002 NotConfigured、E003 BatNotFound、E004 SubprocessException
  - E005 NonZeroReturnCode、E006 CommandNotFound、E007 SystemCannotFindFile、E008 AccessDenied
  - E101 DownloadFailedGeneric、E201 UploadFailedGeneric、E000 Unknown（保底）