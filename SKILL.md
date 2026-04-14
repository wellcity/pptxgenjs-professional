---
name: pptxgenjs-professional
description: "使用 pptxgenjs 快速產生專業級 PowerPoint 簡報。包含完整設計系統： Midnight Executive 配色、統一 Header/Footer、數據卡片、架構圖、流程圖、icon 系統。只需替換內容即可產生多頁專業簡報。用於任何需要 .pptx 输出的场景。"
version: "1.0.0"
author: "Bruce Hung"
license: MIT

tags:
  - powerpoint
  - presentation
  - pptxgenjs
  - business

languages:
  - zh
  - en
---

# PptxGenJS 專業簡報產生技能

## 快速開始

```bash
# 安裝依賴（只需一次）
cd /tmp && npm init -y && npm install pptxgenjs react react-dom sharp react-icons

# 複製範本並修改內容
cp /path/to/template.js my-presentation.js
# 編輯 my-presentation.js 中的內容（標題、描述、數據）
node my-presentation.js
# 輸出：my-presentation.pptx
```

## 設計系統

### 配色：Midnight Executive

```javascript
const C = {
  navy:    "1E2761",   // 主色：深海軍藍
  mid:     "2E4A8F",   // 中間藍
  teal:    "0078D4",   // Microsoft 藍（主強調色）
  ice:     "CADCFC",   // 淡冰藍
  white:   "FFFFFF",
  offWhite:"F5F7FA",
  darkText:"1A1A2E",
  mutedText:"5A6B8A",  // 灰色說明文字
  accent:  "00B4D8",   // 青色點綴
  success: "10B981",   // 綠色
  coral:   "E85D4C",   // 珊瑚色
  gold:    "F5A623",   // 金色
  purple:  "6B5CE7",   // 紫色
  darkBg:  "0D1B3E",   // 深色背景（標題/結論頁）
};
```

### 字體
- 標題：`Arial Black`（數字用 `Arial Black` 突出）
- 內文：`Calibri`
- 標題大小：36-46pt
- 內文大小：11-14pt
- 說明文字：10-12pt

## 標準版面結構

每頁包含：
- **Header**（深藍底，白字，高度 1.1"）
- **Footer**（深藍底，淺藍小字，高度 0.4"，含標題和年份）
- **內容區**（offWhite 或白色卡片背景）

## 工具函式

### makeShadow() — 卡片陰影
```javascript
const makeShadow = () => ({
  type: "outer", blur: 10, offset: 4, angle: 135, color: "000000", opacity: 0.22
});
// ⚠️ 每次都要新的物件，pptxgenjs 會 mutate 這個物件
```

### iconPng(Icon, color, size) — Icon 轉 PNG
```javascript
async function iconPng(Icon, color = "#0078D4", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Icon, { color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
// 用法：
const iconData = await iconPng(FaRocket, "#FFFFFF", 256);
slide.addImage({ data: iconData, x: 1, y: 1, w: 0.6, h: 0.6 });
```

### addHeader(slide, title) — 頁面頂部導航欄
```javascript
function addHeader(slide, title) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 1.1, fill: { color: C.navy } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.05, w: 13.3, h: 0.06, fill: { color: C.teal } });
  slide.addText(title, { x: 0.5, y: 0.25, w: 12, h: 0.7, fontSize: 26, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
}
```

### addFooter(slide, text) — 頁面底部 Footer
```javascript
function addFooter(slide, text = "Title · Subtitle") {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 7.1, w: 13.3, h: 0.4, fill: { color: C.navy } });
  slide.addText(text, { x: 0.5, y: 7.15, w: 9, h: 0.3, fontSize: 10, fontFace: "Calibri", color: C.ice, margin: 0 });
  slide.addText("2026", { x: 11.5, y: 7.15, w: 1.5, h: 0.3, fontSize: 10, fontFace: "Calibri", color: C.ice, align: "right", margin: 0 });
}
```

### statBox(slide, x, y, value, label, color) — 統計數據卡片
```javascript
// 輸出：2.2" x 1.3" 的統計數字卡片
statBox(slide, 0.5, 5.5, "600+", "可用模型", C.teal);
```

### bulletList(slide, items, x, y, w, h, fontSize) — 項目符號列表
```javascript
const items = ["第一點", "第二點", "第三點"];
bulletList(slide, items, 0.5, 2.0, 5.0, 3.0, 13);
```

### makeCard(slide, x, y, w, h, accentColor) — 帶左側色條的卡片
```javascript
// 快速建立一個白色卡片，左側有彩色豎條
makeCard(slide, 0.5, 1.5, 4.0, 2.5, C.teal);
```

## 頁面類型範本

### 1. 標題頁（深色背景）
```javascript
async function slideTitle() {
  const slide = pres.addSlide();
  slide.background = { color: C.darkBg };

  // 裝飾圓（右上）
  slide.addShape(pres.shapes.OVAL, { x: 8, y: -2.5, w: 8, h: 8,
    fill: { color: C.navy, transparency: 50 }, line: { color: C.teal, width: 2 } });

  // 左側豎線
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 0.12, h: 4.2, fill: { color: C.teal } });

  slide.addText("主標題", { x: 0.85, y: 1.6, w: 7, h: 0.9, fontSize: 52, fontFace: "Arial Black", color: C.white });
  slide.addText("副標題", { x: 0.85, y: 2.4, w: 9, h: 1.0, fontSize: 46, fontFace: "Arial Black", color: C.teal });

  // Tag 標籤（底部）
  const tags = ["Tag1", "Tag2", "Tag3"];
  tags.forEach((tag, i) => {
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.85 + i * 2.1, y: 5.1, w: 1.9, h: 0.38, fill: { color: C.teal, transparency: 80 } });
    slide.addText(tag, { x: 0.85 + i * 2.1, y: 5.1, w: 1.9, h: 0.38, fontSize: 10, color: C.ice, align: "center", valign: "middle" });
  });
}
```

### 2. 定義 + 三欄特色（首頁內容頁）
```javascript
async function slideThreePillars() {
  const pillars = [
    { icon: FaBrain, title: "特色一", desc: "說明文字...", color: C.teal },
    { icon: FaPlug, title: "特色二", desc: "說明文字...", color: C.coral },
    { icon: FaCloud, title: "特色三", desc: "說明文字...", color: C.success },
  ];
  for (let i = 0; i < pillars.length; i++) {
    const x = 0.5 + i * 4.15;
    const iconData = await iconPng(pillars[i].icon, "#FFFFFF", 256);
    // 主卡片
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 2.7, w: 3.95, h: 3.8, fill: { color: C.white }, shadow: makeShadow() });
    // 頂部色塊
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 2.7, w: 3.95, h: 1.0, fill: { color: c } });
    // 數字
    slide.addText(`0${i + 1}`, { x: x + 0.15, y: 2.75, w: 1, h: 0.8, fontSize: 36, fontFace: "Arial Black", color: C.white });
    // Icon
    slide.addImage({ data: iconData, x: x + 2.8, y: 2.78, w: 0.85, h: 0.85 });
    // 標題 + 說明
    slide.addText(pillars[i].title, { x: x + 0.2, y: 3.8, w: 3.55, h: 0.5, fontSize: 15, color: C.navy, bold: true });
    slide.addText(pillars[i].desc, { x: x + 0.2, y: 4.35, w: 3.55, h: 2.0, fontSize: 12, color: C.mutedText });
  }
}
```

### 3. 左文右圖（功能說明 + 架構圖）
```javascript
async function slideTwoColumn() {
  // 左：功能列表
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 5.8, h: 5.6, fill: { color: C.white }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 5.8, h: 0.08, fill: { color: C.teal } });
  slide.addText("標題", { x: 0.7, y: 1.45, w: 5.4, h: 0.5, fontSize: 16, color: C.navy, bold: true });
  bulletList(slide, items, 0.7, 2.5, 5.4, 3.2, 13);

  // 右：架構分層
  slide.addShape(pres.shapes.RECTANGLE, { x: 6.6, y: 1.3, w: 6.2, h: 5.6, fill: { color: C.white }, shadow: makeShadow() });
  const layers = [
    { label: "第一層", sub: "子說明", color: C.ice, textColor: C.darkText },
    { label: "第二層", sub: "子說明", color: C.teal, textColor: C.white },
    { label: "第三層", sub: "子說明", color: C.mid, textColor: C.white },
  ];
  for (let i = 0; i < layers.length; i++) {
    const y = 2.05 + i * 0.92;
    slide.addShape(pres.shapes.RECTANGLE, { x: 6.8, y, w: 5.8, h: 0.78, fill: { color: layers[i].color } });
    slide.addText(layers[i].label, { x: 7.0, y: y + 0.08, w: 5.4, h: 0.38, fontSize: 13, color: layers[i].textColor, bold: true });
    slide.addText(layers[i].sub, { x: 7.0, y: y + 0.42, w: 5.4, h: 0.3, fontSize: 10, color: layers[i].textColor });
  }
}
```

### 4. 網格卡片（2x3 或 3x2）
```javascript
async function slideGrid() {
  const items = [
    { icon: FaCogs, title: "標題", desc: "說明", color: C.teal },
    { icon: FaShieldAlt, title: "標題", desc: "說明", color: C.coral },
    { icon: FaCode, title: "標題", desc: "說明", color: C.success },
    { icon: FaChartLine, title: "標題", desc: "說明", color: C.gold },
    { icon: FaRocket, title: "標題", desc: "說明", color: C.purple },
    { icon: FaUsers, title: "標題", desc: "說明", color: C.mid },
  ];
  for (let i = 0; i < items.length; i++) {
    const row = Math.floor(i / 3);
    const col = i % 3;
    const x = 0.5 + col * 4.15;
    const y = 1.3 + row * 2.8;
    const iconData = await iconPng(items[i].icon, "#FFFFFF", 256);
    // 主卡片
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.95, h: 2.55, fill: { color: C.white }, shadow: makeShadow() });
    // 頂部色條
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.95, h: 0.9, fill: { color: items[i].color } });
    // 數字
    slide.addText(String(i + 1).padStart(2, "0"), { x: x + 0.15, y: y + 0.1, w: 0.7, h: 0.7, fontSize: 28, fontFace: "Arial Black", color: C.white });
    // Icon
    slide.addImage({ data: iconData, x: x + 2.85, y: y + 0.12, w: 0.75, h: 0.75 });
    // 標題 + 說明
    slide.addText(items[i].title, { x: x + 0.2, y: y + 1.0, w: 3.55, h: 0.45, fontSize: 14, color: C.navy, bold: true });
    slide.addText(items[i].desc, { x: x + 0.2, y: y + 1.5, w: 3.55, h: 0.95, fontSize: 11, color: C.mutedText });
  }
}
```

### 5. 步驟流程（垂直時間線）
```javascript
async function slideSteps() {
  const steps = [
    { num: "01", title: "步驟一", desc: "說明", icon: FaServer, color: C.teal },
    { num: "02", title: "步驟二", desc: "說明", icon: FaRocket, color: C.gold },
    { num: "03", title: "步驟三", desc: "說明", icon: FaBrain, color: C.coral },
  ];
  for (let i = 0; i < steps.length; i++) {
    const y = 1.3 + i * 1.1;
    const iconData = await iconPng(steps[i].icon, "#FFFFFF", 256);
    // 連接線
    if (i < steps.length - 1)
      slide.addShape(pres.shapes.RECTANGLE, { x: 1.3, y: y + 0.85, w: 0.04, h: 0.25, fill: { color: steps[i].color, transparency: 50 } });
    // 號碼圓
    slide.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.05, w: 0.8, h: 0.8, fill: { color: steps[i].color } });
    slide.addText(steps[i].num, { x: 0.5, y: y + 0.05, w: 0.8, h: 0.8, fontSize: 20, fontFace: "Arial Black", color: C.white, align: "center", valign: "middle" });
    // 內容卡片
    slide.addShape(pres.shapes.RECTANGLE, { x: 1.55, y, w: 11.25, h: 0.9, fill: { color: C.white }, shadow: makeShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: 1.55, y, w: 0.08, h: 0.9, fill: { color: steps[i].color } });
    slide.addImage({ data: iconData, x: 1.75, y: y + 0.15, w: 0.55, h: 0.55 });
    slide.addText(steps[i].title, { x: 2.5, y: y + 0.08, w: 4, h: 0.4, fontSize: 14, color: C.navy, bold: true });
    slide.addText(steps[i].desc, { x: 2.5, y: y + 0.48, w: 10, h: 0.35, fontSize: 12, color: C.mutedText });
  }
}
```

### 6. 統計數據卡（2x2 排列）
```javascript
async function slideStats() {
  const stats = [
    { value: "600+", label: "可用模型", color: C.teal },
    { value: "1,000+", label: "企業連接器", color: C.gold },
    { value: "0", label: "需寫程式碼", color: C.success },
    { value: "Multi", label: "發布管道", color: C.coral },
  ];
  for (let i = 0; i < stats.length; i++) {
    const x = 9.0 + (i % 2) * 2.0;
    const y = 3.2 + Math.floor(i / 2) * 1.8;
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 1.85, h: 1.55, fill: { color: C.navy }, shadow: makeShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 1.85, h: 0.07, fill: { color: stats[i].color } });
    slide.addText(stats[i].value, { x, y: y + 0.2, w: 1.85, h: 0.7, fontSize: 22, fontFace: "Arial Black", color: stats[i].color, align: "center" });
    slide.addText(stats[i].label, { x, y: y + 0.95, w: 1.85, h: 0.45, fontSize: 10, color: C.ice, align: "center" });
  }
}
```

## 完整範本檔案

參考 `/home/wellcity/microsoft-agent-framework.js` — 包含 9 頁完整範例：
1. 標題頁（深色背景 + 裝飾圓）
2. 定義 + 三特色欄
3. 左文右架構圖
4. 3x2 網格卡片
5. 模型比較（帶數據突出）
6. Copilot Studio（深色背景 + 統計卡）
7. 企業應用（3x2 網格）
8. 步驟流程
9. 總結（深色背景 + 重點列表）

## QA 檢查清單

完成後務必執行：
```bash
python3 -m markitdown output.pptx 2>/dev/null | grep -iE "xxxx|lorem|ipsum|placeholder" || echo "✅ 無佔位符"
```

檢查：
- [ ] 所有文字已替換，無「XXXX」「Lorem」等佔位符
- [ ] 每頁有視覺元素（不只是純文字）
- [ ] 深淺背景交替（標題/結論深色，內容淺色）
- [ ] 不同頁面使用不同版面配置（不要每頁長得一樣）

## 常見問題

**Q: Icon 顯示不出來？**
檢查 icon 名稱是否在 react-icons 的 `fa` 套件中存在。某些 icon 名稱在不同版本有差異，先用 `node -e "const i = require('react-icons/fa'); console.log('FaRocket' in i)"` 測試。

**Q: 卡片陰影看起來怪怪的？**
每次 `makeShadow()` 都要 new 一個物件，不要 reuse 否則 pptxgenjs 會把物件 mutation 造成問題。

**Q: 文字被截斷？**
檢查 `w` 寬度是否足夠，特別是中文在同樣 fontSize 下比英文寬，建議留 10% 餘裕。

**Q: 深色背景頁面Footer看不見？**
深色背景頁要自己加 Footer 或把 footer 那行拿掉，預設 `addFooter` 是給淺色背景頁用的。

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
