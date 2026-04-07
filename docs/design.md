# Design System Reference: Vercel

> 本文件記錄 Vercel 設計系統的核心原則，作為 SPC 工具 UI 優化的參考依據。

## 1. 設計哲學

Vercel 的設計理念是「工程師基礎設施的極簡主義」—— 每個元素都必須有其存在的理由。界面就像編譯器對待代碼一樣——去除所有多餘的內容，只保留結構。

### 核心原則
- **功能優先**：裝飾性元素必須有實際用途
- **負空間**：大量留白，內容主導
- **精確性**：每個像素都有其意義

## 2. 字體系統

### 主字體：Geist Sans
- **負間距 (Negative Letter-Spacing)**：
  - 56px 標題：-2.88px
  - 48px 標題：-2.4px
  - 正文：自然間距
- **效果**：標題感覺被壓縮、緊湊、緊急

### 等寬字體：Geist Mono
- 用於程式碼、數據標籤、終端輸出
- 啟用 OpenType `"liga"` 連字功能

### 應用建議
```css
/* 標題使用緊湊間距 */
h1, h2, h3 {
  letter-spacing: -0.02em;
}

/* 數據使用等寬字體 */
.mono-data {
  font-family: 'Geist Mono', 'Courier New', monospace;
}
```

## 3. 色彩系統

### 主色調
| 顏色 | 代碼 | 用途 |
|------|------|------|
| Vercel Black | `#171717` | 主要標題、深色表面 |
| Pure White | `#ffffff` | 頁面背景、卡片表面 |
| True Black | `#000000` | 代碼/控制台上下文 |

### 工作流程強調色
| 名稱 | 代碼 | 用途 |
|------|------|------|
| Ship Red | `#ff5b4f` | 生產部署流程 |
| Preview Pink | `#de1d8d` | 預覽部署流程 |
| Develop Blue | `#0a72ef` | 開發流程 |

### 灰色漸層
```
Gray 900 (#171717) → 主文字、標題
Gray 600 (#4d4d4d) → 次要文字
Gray 500 (#666666) → 淡化文字
Gray 400 (#808080) → 佔位符
Gray 100 (#ebebeb) → 邊框、分割線
Gray 50  (#fafafa) → 表面淡化
```

## 4. 邊框系統 (Shadow-as-Border)

### 核心技術
Vercel 不使用傳統 CSS 邊框，而是使用陰影模擬邊框：

```css
/* 替代傳統邊框 */
box-shadow: 0px 0px 0px 1px rgba(0, 0, 0, 0.08);
```

### 優勢
1. 無 box model 邊界問題
2. 圓角不會被裁剪
3. 過渡動畫更平滑
4. 視覺重量更輕

### 應用
```css
/* 卡片邊框 */
.card {
  border-radius: 8px;
  box-shadow: 0px 0px 0px 1px rgba(0, 0, 0, 0.08);
}

/* 多層陰影堆疊 */
.elevated {
  box-shadow: 
    0px 0px 0px 1px rgba(0, 0, 0, 0.08),  /* 邊框 */
    0px 8px 16px rgba(0, 0, 0, 0.08),     /* 提升 */
    0px 16px 24px rgba(0, 0, 0, 0.04);   /* 環境 */
}
```

## 5. 互動元素

### 焦點環 (Focus Ring)
```css
:focus-visible {
  outline: 2px solid hsla(212, 100%, 48%, 1);
  outline-offset: 2px;
}
```

### 圓角系統
- 按鈕：4px-8px（保守半徑，無全圓形）
- 輸入框：6px
- 卡片：8px
- 徽章：9999px（藥丸形）

### 狀態指示器
- 徽章使用藥丸形狀 + 淡化背景
- 成功：綠色
- 警告：黃色
- 錯誤：紅色

## 6. 應用於 SPC 工具

### 建議變更

#### 字體調整
- 標題使用負間距
- 數據/統計值使用等寬字體

#### 邊框替換
- 將傳統邊框改為 shadow-as-border
- 管制圖邊框採用更輕的樣式

#### 色彩優化
- 深色模式使用 #171717 而非純黑
- 根據數據狀態使用對應強調色（紅/黃/綠）

#### 間距系統
- 維持 8px 為基礎單位
- 數據展示區域保持緊湊

## 7. 圖表配色參考

### 統計圖表強調色
| 狀態 | 顏色 | 用途 |
|------|------|------|
| 正常 | `#10b981` (Green) | 管制界限內 |
| 警告 | `#f59e0b` (Amber) | 接近界限 |
| 異常 | `#ef4444` (Red) | 超出界限 |
| 資訊 | `#0a72ef` (Blue) | 中心線 |

## 8. 相關資源

- [Vercel Design System](https://vercel.com/design)
- [Geist Font](https://vercel.com/font)
- [Tailwind CSS](https://tailwindcss.com)（已在本專案使用）

---

> 本文件應與 `.cursorrules` 和 `.agent/workflows/developer_sop.md` 配合使用。