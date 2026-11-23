---
name: safe-deploy-streamlit
description: >-
  Use this skill whenever verifying, committing, and pushing Hotel Master code or data snapshots
  to GitHub for Streamlit Cloud deployment. Replaces the deprecated legacy push_to_github workflow.
---

# 安全驗證與推送部署技能（取代舊版 push_to_github）

> **絕對禁令**：嚴禁執行 `git add .` 或 `git add -A`！專案根目錄或暫存區可能存在未追蹤的除錯檔或機密檔案，所有提交必須採**白名單精準加檔**。

## 標準上線 4 步驟

### Step 1：Python 語法編譯檢核
對本次修改的所有 `.py` 檔案執行 `py_compile`，確認零語法錯誤：
```powershell
C:\Users\88691\AppData\Local\Python\pythoncore-3.14-64\python.exe -m py_compile app.py <modified_module.py>
```

### Step 2：核心邏輯煙霧測試（Smoke Test）
針對新寫或修改的計算函式（如報價、對帳、匯出 Excel/ZIP）執行簡短斷言驗證，確保：
- 無寫死假資料、無欄位缺失、數值與原始表單對齊。
- 確認在無區網 SQL Server 的雲端情境下可正常讀取本地備份（`data/pms_backup/`）。

### Step 3：檢視變更與白名單精準加檔
1. 先檢查變更範圍：
   ```powershell
   git status --short; git diff --stat
   ```
2. **確認未包含任何機密檔**（`.streamlit/secrets.toml`、`*.json` 憑證檔）與 `_archive/` 舊腳本。
3. 僅將本次確定要上線的檔案逐一加入暫存區並提交（使用語意化 Conventional Commit）：
   ```powershell
   git add <file1> <file2>; git commit -m "feat(<scope>): <concise summary>"
   ```

### Step 4：推送至遠端觸發 Streamlit Cloud 部署
```powershell
git push origin main
```
推送完成後回報 commit hash（如 `cd5d078..24e5425`），並提醒使用者 Streamlit Cloud 約 1～2 分鐘完成自動重新部署。
