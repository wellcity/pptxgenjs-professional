# pptxgenjs-professional

使用 pptxgenjs 產生專業級 PowerPoint 簡報的技能。採用 **Midnight Executive** 設計系統（微軟風格深藍配色），包含統一的 Header/Footer、數據卡片、架構圖、流程圖和完整的 icon 系統。

## 快速開始

```bash
# 安裝依賴（只需一次）
cd /tmp && npm init -y && npm install pptxgenjs react react-dom sharp react-icons

# 複製範本
cp references/template.js my-presentation.js

# 修改內容後產生
node my-presentation.js
# 輸出：output.pptx
```

## 檔案說明

| 檔案 | 說明 |
|------|------|
| `SKILL.md` | 完整設計系統文件 |
| `references/template.js` | 最小範本，快速開始 |
| `references/microsoft-agent-framework.js` | 9 頁完整範例（Microsoft Agent Framework 主題） |

## 設計系統

### 配色：Midnight Executive

| 名稱 | Hex | 用途 |
|------|-----|------|
| navy | `1E2761` | 主色：深海軍藍 |
| teal | `0078D4` | 主強調色（微軟藍） |
| mid | `2E4A8F` | 中間藍 |
| ice | `CADCFC` | 淡冰藍 |
| coral | `E85D4C` | 珊瑚色 |
| gold | `F5A623` | 金色 |
| success | `10B981` | 綠色 |
| purple | `6B5CE7` | 紫色 |
| darkBg | `0D1B3E` | 深色背景（標題/結論頁） |

### 標準版面結構

每頁包含：
- **Header**（深藍底、白字，高度 1.1"）
- **Footer**（深藍底、淺藍小字，高度 0.4"）
- **內容區**（淡灰或白色卡片背景）

## 頁面類型

| 類型 | 說明 |
|------|------|
| 標題頁 | 深色背景 + 裝飾圓形 |
| 定義 + 三特色 | 滿寬定義區 + 3 張 icon 卡片 |
| 左文右圖 | 左側功能列表、右側架構圖 |
| 3×2 網格卡片 | 6 張帶 icon、顏色、數據的卡片 |
| 步驟流程 | 垂直時間線，編號圓圈連接 |
| 統計儀表板 | 2×2 或 2×3 統計數據方塊 |
| 總結頁 | 深色背景 + 重點列表 |

## QA 檢查清單

```bash
python3 -m markitdown output.pptx 2>/dev/null | grep -iE "xxxx|lorem|ipsum|placeholder" || echo "✅ 無佔位符"
```

- [ ] 所有文字已替換，無「XXXX」「Lorem」等佔位符
- [ ] 每頁有視覺元素（不只是純文字）
- [ ] 深淺背景交替（標題/結論 = 深色，內容 = 淺色）
- [ ] 不同頁面使用不同版面配置（避免重複）

## 依賴版本

```json
{
  "pptxgenjs": "^4.0.1",
  "react": "^19.2.0",
  "react-dom": "^19.2.0",
  "react-icons": "^5.6.0",
  "sharp": "^0.34.0"
}
```
