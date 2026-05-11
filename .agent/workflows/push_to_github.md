---
description: 自動提交並推送程式碼至 GitHub (同步雲端部署)
---
這是一個快速將本地更動推送到 GitHub 的工作流程，主要用於觸發 Streamlit Cloud 等雲端服務的重新部署。

// turbo-all
1. 檢查目前的檔案變更狀態
   ```bash
   git status
   ```

2. 將所有變更加入暫存區
   ```bash
   git add .
   ```

3. 提交變更 (將自動套用概括性的修正訊息，如果需要客製化訊息可以手動調整)
   ```bash
   git commit -m "Auto-sync updates and fixes to Github"
   ```

4. 推送至遠端 GitHub
   ```bash
   git push
   ```
