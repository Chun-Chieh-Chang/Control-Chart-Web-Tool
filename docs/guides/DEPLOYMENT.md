# GitHub Pages 部署指南 (Deployment Guide)

## 部署方式 (Deployment Method)

本專案使用 **GitHub Actions** 自動部署至 GitHub Pages。

### 自動部署流程 (Automated Deployment Flow)

```mermaid
graph LR
    A[Push to main] --> B[GitHub Actions Triggered]
    B --> C[Build & Deploy]
    C --> D[GitHub Pages Updated]
```

### 配置說明 (Configuration Details)

#### 1. GitHub Actions 工作流程
- **檔案位置**: `.github/workflows/deploy.yml`
- **觸發條件**: 
  - 推送到 `main` 分支時自動觸發
  - 支援手動觸發 (workflow_dispatch)
- **部署權限**: 
  - `contents: read` - 讀取儲存庫內容
  - `pages: write` - 寫入 GitHub Pages
  - `id-token: write` - 用於身份驗證

#### 2. 並發控制 (Concurrency Control)
- **群組**: `pages`
- **策略**: 不取消進行中的部署，確保生產部署完整執行
- **目的**: 避免部署衝突，維持頁面穩定性

#### 3. 部署步驟 (Deployment Steps)
1. **Checkout**: 檢出專案程式碼
2. **Setup Pages**: 配置 GitHub Pages 環境
3. **Upload artifact**: 上傳整個專案作為靜態網站
4. **Deploy**: 部署到 GitHub Pages

### GitHub Pages 設定 (GitHub Pages Settings)

在 GitHub Repository 設定中：
1. 進入 **Settings** → **Pages**
2. **Source** 設定為: `GitHub Actions`
3. 無需指定分支（由 Actions 自動處理）

### 部署狀態查看 (Deployment Status)

- **方式 1**: 查看 Repository 的 **Actions** 標籤頁
- **方式 2**: 查看 Commit 旁的狀態圖示
- **方式 3**: 在 **Deployments** 區域查看部署歷史

### 注意事項 (Important Notes)

1. **靜態網站**: 本專案為純前端靜態網站，無需建置步驟
2. **.nojekyll**: 已包含此檔案，確保所有檔案正確部署
3. **部署時間**: 通常在 1-2 分鐘內完成
4. **快取問題**: 若更新未顯示，請清除瀏覽器快取或使用無痕模式

### 本地測試 (Local Testing)

部署前建議本地測試：

```bash
# 使用 Python 內建伺服器
python -m http.server 8000

# 使用 Node.js http-server
npx http-server -p 8000
```

然後訪問 `http://localhost:8000`

### 疑難排解 (Troubleshooting)

#### 問題：Actions 執行失敗
**解決方案**: 
- 檢查 Actions 標籤頁的錯誤訊息
- 確認 Repository 有啟用 Actions 權限
- 驗證 Pages 設定正確

#### 問題：更新未反映在網站上
**解決方案**:
- 確認 Actions 執行成功（綠色勾勾）
- 等待 1-2 分鐘讓 CDN 更新
- 清除瀏覽器快取（Ctrl+Shift+R）

#### 問題：404 錯誤
**解決方案**:
- 確認 `index.html` 存在於根目錄
- 檢查 `.nojekyll` 檔案存在
- 確認 Pages 設定中的 Source 為 GitHub Actions

---

**最後更新**: 2026-01-31  
**維護者**: Chun-Chieh-Chang
