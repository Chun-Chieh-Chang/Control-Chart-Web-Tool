# GitHub Pages 部署指南 / Deployment Guide

## 中文說明

### 步驟 1: 初始化 Git 倉庫

```bash
cd c:\Users\3kids\Downloads\ControlChart
git init
git add .
git commit -m "Initial commit: SPC Analysis Tool"
```

### 步驟 2: 創建 GitHub 倉庫

1. 訪問 [GitHub](https://github.com/new)
2. 倉庫名稱: `Mouldex-Control-Chart`
3. 描述: `SPC Statistical Process Control Analysis Tool`
4. 設為 Public（必須，才能使用 GitHub Pages）
5. 不要初始化 README（我們已經有了）
6. 點擊 "Create repository"

### 步驟 3: 推送到 GitHub

```bash
git remote add origin https://github.com/YOUR_USERNAME/Mouldex-Control-Chart.git
git branch -M main
git push -u origin main
```

### 步驟 4: 啟用 GitHub Pages

#### 方法 A: 使用 `web` 資料夾

1. 進入倉庫設定: Settings → Pages
2. Source: Deploy from a branch
3. Branch: `main`
4. Folder: `/web`
5. 點擊 Save

#### 方法 B: 使用 `docs` 資料夾（推薦）

如果使用此方法，需先移動檔案：

```bash
# 將 web 資料夾重命名為 docs
mv web docs
git add .
git commit -m "Rename web to docs for GitHub Pages"
git push
```

然後在 GitHub 設定：
1. Settings → Pages
2. Branch: `main`
3. Folder: `/docs`
4. Save

### 步驟 5: 訪問網站

等待 1-2 分鐘，訪問:
```
https://YOUR_USERNAME.github.io/Mouldex-Control-Chart
```

---

## English Instructions

### Step 1: Initialize Git Repository

```bash
cd c:\Users\3kids\Downloads\ControlChart
git init
git add .
git commit -m "Initial commit: SPC Analysis Tool"
```

### Step 2: Create GitHub Repository

1. Visit [GitHub](https://github.com/new)
2. Repository name: `Mouldex-Control-Chart`
3. Description: `SPC Statistical Process Control Analysis Tool`
4. Set to Public (required for GitHub Pages)
5. Don't initialize README (we already have one)
6. Click "Create repository"

### Step 3: Push to GitHub

```bash
git remote add origin https://github.com/YOUR_USERNAME/Mouldex-Control-Chart.git
git branch -M main
git push -u origin main
```

### Step 4: Enable GitHub Pages

#### Option A: Using `web` folder

1. Go to repository Settings → Pages
2. Source: Deploy from a branch
3. Branch: `main`
4. Folder: `/web`
5. Click Save

#### Option B: Using `docs` folder (Recommended)

If using this method, move files first:

```bash
# Rename web folder to docs
mv web docs
git add .
git commit -m "Rename web to docs for GitHub Pages"
git push
```

Then in GitHub settings:
1. Settings → Pages
2. Branch: `main`
3. Folder: `/docs`
4. Save

### Step 5: Visit Website

Wait 1-2 minutes, then visit:
```
https://YOUR_USERNAME.github.io/Mouldex-Control-Chart
```

---

## 故障排除 / Troubleshooting

### 問題：頁面顯示 404
檢查:
- GitHub Pages 是否已啟用
- Branch 和 Folder 設定是否正確
- 等待幾分鐘讓 GitHub 部署

### 問題：JavaScript 錯誤
確保:
- 所有 CDN 連結可訪問
- 瀏覽器支援 ES6 Modules
- 瀏覽器控制台檢查錯誤訊息

### 問題：Excel 上傳失敗
確認:
- 檔案格式正確（.xlsx 或 .xls）
- 檔案大小合理（< 10MB）
- Excel 結構符合要求（參考 README）
