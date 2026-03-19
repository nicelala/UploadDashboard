# Release Checklist

- [ ] 更新 /VERSION 與 /CHANGELOG.md（寫明破壞性變更、日期）
- [ ] 以 One Directory + Window Based 打包，於乾淨機器驗證（含「規則 → 重新載入規則」）
- [ ] 產出 zip：`UploadDashboard_vX.Y.Z_onedir_YYYYMMDD.zip`（內容為 dist\UploadDashboard\*）
- [ ] 建立並推送 Tag：
      git add VERSION CHANGELOG.md docs/schema.md docs/release-checklist.md
      git commit -m "chore(release): cut vX.Y.Z"
      git tag -a vX.Y.Z -m "vX.Y.Z - summary"
      git push origin main --tags
- [ ] 記錄 zip 的 SHA256（可附在 CHANGELOG 同一版塊）
- [ ] 不隨發佈包附 `error_keywords.csv`（由使用端自行維護）