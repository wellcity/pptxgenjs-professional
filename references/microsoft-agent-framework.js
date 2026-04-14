const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaBrain, FaPlug, FaCloud, FaCogs, FaShieldAlt, FaCode, FaChartLine, FaRocket, FaUsers, FaLightbulb, FaCheckCircle, FaDatabase, FaNetworkWired, FaLock, FaBolt, FaServer, FaLayerGroup, FaSyncAlt } = require("react-icons/fa");

// ─── 配色 ───
const C = {
  navy: "1E2761", mid: "2E4A8F", teal: "0078D4",
  ice: "CADCFC", white: "FFFFFF", offWhite: "F5F7FA",
  darkText: "1A1A2E", mutedText: "5A6B8A",
  accent: "00B4D8", success: "10B981",
  darkBg: "0D1B3E", coral: "E85D4C",
  gold: "F5A623", purple: "6B5CE7",
};

const makeShadow = () => ({ type: "outer", blur: 10, offset: 4, angle: 135, color: "000000", opacity: 0.22 });
const makeCard = (slide, x, y, w, h, accentColor = C.teal) => {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.white }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.1, h, fill: { color: accentColor } });
};

async function iconPng(Icon, color = "#0078D4", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Icon, { color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

let pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Bruce Hung";
pres.title = "Microsoft Agent Framework";

// ─── 工具函式 ───
function addHeader(slide, title) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 1.1, fill: { color: C.navy } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.05, w: 13.3, h: 0.06, fill: { color: C.teal } });
  slide.addText(title, { x: 0.5, y: 0.25, w: 12, h: 0.7, fontSize: 26, fontFace: "Arial", color: C.white, bold: true, margin: 0 });
}

function addFooter(slide, text = "Microsoft Agent Framework · Azure AI 代理服務") {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 7.1, w: 13.3, h: 0.4, fill: { color: C.navy } });
  slide.addText(text, { x: 0.5, y: 7.15, w: 9, h: 0.3, fontSize: 10, fontFace: "Calibri", color: C.ice, margin: 0 });
  slide.addText("2026", { x: 11.5, y: 7.15, w: 1.5, h: 0.3, fontSize: 10, fontFace: "Calibri", color: C.ice, align: "right", margin: 0 });
}

function statBox(slide, x, y, value, label, color = C.teal) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.2, h: 1.3, fill: { color: C.white }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.2, h: 0.08, fill: { color: color } });
  slide.addText(value, { x, y: y + 0.15, w: 2.2, h: 0.7, fontSize: 32, fontFace: "Arial Black", color, align: "center", margin: 0 });
  slide.addText(label, { x, y: y + 0.85, w: 2.2, h: 0.35, fontSize: 11, fontFace: "Calibri", color: C.mutedText, align: "center", margin: 0 });
}

function bulletList(slide, items, x, y, w, h, fontSize = 13) {
  slide.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 10 } })), {
    x, y, w, h, fontSize, fontFace: "Calibri", color: C.darkText, margin: 0
  });
}

// ─── 第 1 頁：標題 ───
async function slide1() {
  const slide = pres.addSlide();
  slide.background = { color: C.darkBg };

  // 大裝飾圓（右上）
  slide.addShape(pres.shapes.OVAL, { x: 8, y: -2.5, w: 8, h: 8, fill: { color: C.navy, transparency: 50 }, line: { color: C.teal, width: 2 } });
  slide.addShape(pres.shapes.OVAL, { x: 9.5, y: 3.5, w: 5, h: 5, fill: { color: C.teal, transparency: 75 }, line: { color: C.accent, width: 1 } });
  slide.addShape(pres.shapes.OVAL, { x: 0.5, y: 5, w: 3, h: 3, fill: { color: C.mid, transparency: 70 }, line: { color: C.ice, width: 1 } });

  // 左側豎線
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 0.12, h: 4.2, fill: { color: C.teal } });

  // 細線裝飾
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 4, h: 0.04, fill: { color: C.teal } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.66, w: 6, h: 0.04, fill: { color: C.teal, transparency: 50 } });

  slide.addText("Microsoft", { x: 0.85, y: 1.6, w: 7, h: 0.9, fontSize: 52, fontFace: "Arial Black", color: C.white, margin: 0 });
  slide.addText("Agent Framework", { x: 0.85, y: 2.4, w: 9, h: 1.0, fontSize: 46, fontFace: "Arial Black", color: C.teal, margin: 0 });

  slide.addText("Azure AI 代理服務  ·  Copilot Studio  ·  AI Studio", {
    x: 0.85, y: 3.55, w: 8, h: 0.5, fontSize: 16, fontFace: "Calibri", color: C.accent, margin: 0
  });

  // 分隔線
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.85, y: 4.2, w: 3, h: 0.03, fill: { color: C.ice, transparency: 50 } });

  slide.addText("企業級 AI 代理：建構、部署與編排智能系統", {
    x: 0.85, y: 4.4, w: 7, h: 0.5, fontSize: 14, fontFace: "Calibri", color: C.ice, margin: 0
  });

  // 右下角標籤
  const tags = ["GPT-4", "Llama 3", "Phi-4", "Azure Native", "Enterprise AI"];
  tags.forEach((tag, i) => {
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.85 + i * 2.1, y: 5.1, w: 1.9, h: 0.38, fill: { color: C.teal, transparency: 80 } });
    slide.addText(tag, { x: 0.85 + i * 2.1, y: 5.1, w: 1.9, h: 0.38, fontSize: 10, fontFace: "Calibri", color: C.ice, align: "center", valign: "middle", margin: 0 });
  });

  slide.addText("2026  |  Bruce Hung", { x: 0.85, y: 6.6, w: 4, h: 0.4, fontSize: 12, fontFace: "Calibri", color: C.mutedText, margin: 0 });
}

// ─── 第 2 頁：什麼是 ───
async function slide2() {
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  addHeader(slide, "什麼是 Microsoft Agent Framework？");
  addFooter(slide);

  // 定義區 - 滿寬卡片
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 12.3, h: 1.2, fill: { color: C.navy } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.45, w: 12.3, h: 0.05, fill: { color: C.teal } });
  slide.addText("一個完整的企業級平台，用於建構、部署和管理能夠推理、規劃並跨企業資料與服務採取行動的 AI 代理系統。", {
    x: 0.7, y: 1.4, w: 11.9, h: 1.0, fontSize: 17, fontFace: "Calibri", color: C.white, valign: "middle", margin: 0
  });

  // 三大特色
  const pillars = [
    { icon: FaBrain, title: "推理與規劃", desc: "LLM 驅動的鏈式思維推理引擎，讓代理能夠分解複雜任務並逐步執行", color: C.teal },
    { icon: FaPlug, title: "工具整合", desc: "無縫連接 API、資料庫、檔案系統與企業 SaaS，代理可主動呼叫外部能力", color: C.coral },
    { icon: FaCloud, title: "Azure 原生", desc: "企业级安全性、Entra ID 單一登入、RBAC 權限控制與 SOC 2 合規認證", color: C.success },
  ];

  for (let i = 0; i < pillars.length; i++) {
    const x = 0.5 + i * 4.15;
    const iconData = await iconPng(pillars[i].icon, "#FFFFFF", 256);
    const c = pillars[i].color;

    // 主卡片
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 2.7, w: 3.95, h: 3.8, fill: { color: C.white }, shadow: makeShadow() });
    // 頂部色塊
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 2.7, w: 3.95, h: 1.0, fill: { color: c } });
    // 數字標記
    slide.addText(`0${i + 1}`, { x: x + 0.15, y: 2.75, w: 1, h: 0.8, fontSize: 36, fontFace: "Arial Black", color: C.white, margin: 0 });
    // icon
    slide.addImage({ data: iconData, x: x + 2.8, y: 2.78, w: 0.85, h: 0.85 });

    slide.addText(pillars[i].title, { x: x + 0.2, y: 3.8, w: 3.55, h: 0.5, fontSize: 15, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
    slide.addText(pillars[i].desc, { x: x + 0.2, y: 4.35, w: 3.55, h: 2.0, fontSize: 12, fontFace: "Calibri", color: C.mutedText, margin: 0 });
  }

  // 底部裝飾
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 6.85, w: 13.3, h: 0.25, fill: { color: C.navy, transparency: 90 } });
}

// ─── 第 3 頁：Azure AI Agent Service ───
async function slide3() {
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  addHeader(slide, "Azure AI Agent Service");
  addFooter(slide);

  // 左：功能列表 + 數據
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 5.8, h: 5.6, fill: { color: C.white }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 5.8, h: 0.08, fill: { color: C.teal } });

  slide.addText("核心平台能力", { x: 0.7, y: 1.45, w: 5.4, h: 0.5, fontSize: 16, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
  slide.addText("全托管服務，用於建構和部署企業級 AI 代理", { x: 0.7, y: 1.95, w: 5.4, h: 0.45, fontSize: 12, fontFace: "Calibri", color: C.mutedText, margin: 0 });

  const features = [
    "托管式代理運行時與編排引擎",
    "支援多模型（OpenAI / Azure OpenAI / Llama）",
    "內建工具呼叫與函式執行",
    "記憶體與對話狀態管理",
    "企業 SSO 與 RBAC 權限控制",
    "Model-as-a-Service 彈性計費",
  ];
  slide.addText(features.map((f, i) => ({
    text: f, options: { bullet: true, breakLine: i < features.length - 1, paraSpaceAfter: 9 }
  })), { x: 0.7, y: 2.5, w: 5.4, h: 3.2, fontSize: 13, fontFace: "Calibri", color: C.darkText, margin: 0 });

  // 右：架構分層圖（視覺化）
  slide.addShape(pres.shapes.RECTANGLE, { x: 6.6, y: 1.3, w: 6.2, h: 5.6, fill: { color: C.white }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 6.6, y: 1.3, w: 6.2, h: 0.08, fill: { color: C.coral } });

  slide.addText("系統架構", { x: 6.8, y: 1.45, w: 5.8, h: 0.5, fontSize: 16, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });

  const layers = [
    { label: "使用者介面", sub: "Teams · Web · WhatsApp · 自定義", color: C.ice, textColor: C.darkText },
    { label: "代理編排層", sub: "Orchestration & Memory", color: C.teal, textColor: C.white },
    { label: "模型推理層", sub: "GPT-4 · Llama · Phi · Mistral", color: C.mid, textColor: C.white },
    { label: "工具與技能層", sub: "API · RAG · 程式執行 · 資料庫", color: C.navy, textColor: C.white },
    { label: "Azure 基礎設施", sub: "資安 · 監控 · 擴充性", color: C.darkBg, textColor: C.white },
  ];

  for (let i = 0; i < layers.length; i++) {
    const y = 2.05 + i * 0.92;
    const lc = layers[i];
    slide.addShape(pres.shapes.RECTANGLE, { x: 6.8, y, w: 5.8, h: 0.78, fill: { color: lc.color } });
    slide.addText(lc.label, { x: 7.0, y: y + 0.08, w: 5.4, h: 0.38, fontSize: 13, fontFace: "Arial", color: lc.textColor, bold: true, margin: 0 });
    slide.addText(lc.sub, { x: 7.0, y: y + 0.42, w: 5.4, h: 0.3, fontSize: 10, fontFace: "Calibri", color: lc.textColor, margin: 0 });
  }

  // 右下角數據
  slide.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 6.55, w: 5.8, h: 0.25, fill: { color: C.teal, transparency: 85 } });
}

// ─── 第 4 頁：核心能力 ───
async function slide4() {
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  addHeader(slide, "核心能力");
  addFooter(slide);

  const caps = [
    { icon: FaCogs, title: "多代理編排", desc: "協調多個專業代理共同處理複雜任務，支援階層式與網狀架構", color: C.teal },
    { icon: FaShieldAlt, title: "企業級安全", desc: "Entra ID、RBAC、資料主權保留、SOC 2 / ISO 27001 認證", color: C.coral },
    { icon: FaCode, title: "程式執行", desc: "沙盒環境安全執行代理生成的程式碼，確保推理過程可控", color: C.success },
    { icon: FaChartLine, title: "完整可觀測性", desc: "Azure Monitor + Application Insights 內建日誌、追蹤與指標", color: C.gold },
    { icon: FaRocket, title: "低程式碼開發", desc: "Copilot Studio 視覺化流程設計器，基本代理無需寫程式碼", color: C.purple },
    { icon: FaUsers, title: "人機協作", desc: "升級流程、審批閘道與導引式介入，關鍵決策有人類把關", color: C.mid },
  ];

  for (let i = 0; i < caps.length; i++) {
    const row = Math.floor(i / 3);
    const col = i % 3;
    const x = 0.5 + col * 4.15;
    const y = 1.3 + row * 2.8;
    const c = caps[i].color;
    const iconData = await iconPng(caps[i].icon, "#FFFFFF", 256);

    // 主卡片
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.95, h: 2.55, fill: { color: C.white }, shadow: makeShadow() });
    // 頂部色條
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.95, h: 0.9, fill: { color: c } });
    // 數字
    slide.addText(String(i + 1).padStart(2, "0"), { x: x + 0.15, y: y + 0.1, w: 0.7, h: 0.7, fontSize: 28, fontFace: "Arial Black", color: C.white, margin: 0 });
    // Icon
    slide.addImage({ data: iconData, x: x + 2.85, y: y + 0.12, w: 0.75, h: 0.75 });
    // 標題
    slide.addText(caps[i].title, { x: x + 0.2, y: y + 1.0, w: 3.55, h: 0.45, fontSize: 14, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
    // 說明
    slide.addText(caps[i].desc, { x: x + 0.2, y: y + 1.5, w: 3.55, h: 0.95, fontSize: 11, fontFace: "Calibri", color: C.mutedText, margin: 0 });
  }
}

// ─── 第 5 頁：支援的模型 ───
async function slide5() {
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  addHeader(slide, "支援的模型");
  addFooter(slide);

  const models = [
    { name: "Azure OpenAI GPT-4 / GPT-4o", badge: "旗艦", desc: "複雜推理與工具呼叫最強大的模型，支援多模態輸入", color: C.teal, stat: "200K", statLabel: "上下文 Tokens" },
    { name: "Meta Llama 3 / 3.1 / 3.2", badge: "開源", desc: "開放權重模型，支援自定義微調與本地部署", color: C.success, stat: "405B", statLabel: "最大參數量" },
    { name: "Microsoft Phi-3 / Phi-4", badge: "高效", desc: "小型語言模型，針對延遲與成本優化，適合即時應用", color: C.gold, stat: "14B", statLabel: "最大參數量" },
    { name: "Mistral / Mixtral", badge: "MoE", desc: "混合專家模型，多樣化任務處理，高效率推理", color: C.coral, stat: "8x22B", statLabel: "專家模型規模" },
    { name: "Azure AI Foundry 模型", badge: "目錄", desc: "600+ 模型透過 Model-as-a-Service 彈性取用", color: C.navy, stat: "600+", statLabel: "可用模型數" },
    { name: "自定義微調模型（BYOM）", badge: "自帶", desc: "自攜模型，完整掌控訓練資料與部署流程", color: C.purple, stat: "自訂", statLabel: "訓練方式" },
  ];

  for (let i = 0; i < models.length; i++) {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const x = 0.5 + col * 6.35;
    const y = 1.3 + row * 1.85;
    const m = models[i];

    // 主卡片
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 6.1, h: 1.65, fill: { color: C.white }, shadow: makeShadow() });
    // 左側色條
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.1, h: 1.65, fill: { color: m.color } });
    // Badge
    slide.addShape(pres.shapes.RECTANGLE, { x: x + 4.65, y: y + 0.12, w: 1.25, h: 0.35, fill: { color: m.color } });
    slide.addText(m.badge, { x: x + 4.65, y: y + 0.12, w: 1.25, h: 0.35, fontSize: 10, fontFace: "Calibri", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

    // 數據突出
    slide.addText(m.stat, { x: x + 0.2, y: y + 0.08, w: 1.8, h: 0.7, fontSize: 26, fontFace: "Arial Black", color: m.color, margin: 0 });
    slide.addText(m.statLabel, { x: x + 0.2, y: y + 0.72, w: 1.8, h: 0.3, fontSize: 9, fontFace: "Calibri", color: C.mutedText, margin: 0 });

    // 名稱與說明
    slide.addText(m.name, { x: x + 2.1, y: y + 0.1, w: 2.4, h: 0.5, fontSize: 13, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
    slide.addText(m.desc, { x: x + 2.1, y: y + 0.62, w: 3.8, h: 0.9, fontSize: 11, fontFace: "Calibri", color: C.mutedText, margin: 0 });
  }
}

// ─── 第 6 頁：Copilot Studio ───
async function slide6() {
  const slide = pres.addSlide();
  slide.background = { color: C.darkBg };

  // 裝飾
  slide.addShape(pres.shapes.OVAL, { x: 9.5, y: -2, w: 7, h: 7, fill: { color: C.navy, transparency: 45 }, line: { color: C.teal, width: 2 } });
  slide.addShape(pres.shapes.OVAL, { x: -1.5, y: 4.5, w: 4, h: 4, fill: { color: C.teal, transparency: 75 } });

  slide.addText("Microsoft", { x: 0.5, y: 0.4, w: 6, h: 0.8, fontSize: 32, fontFace: "Arial", color: C.ice, margin: 0 });
  slide.addText("Copilot Studio", { x: 0.5, y: 1.1, w: 9, h: 0.9, fontSize: 40, fontFace: "Arial Black", color: C.white, margin: 0 });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.05, w: 2.5, h: 0.06, fill: { color: C.teal } });
  slide.addText("低程式碼代理開發平台", { x: 0.5, y: 2.2, w: 6, h: 0.5, fontSize: 16, fontFace: "Calibri", color: C.accent, margin: 0 });

  const features = [
    "視覺化拖放式對話流程設計器",
    "預建代理範本，因應企業常見情境快速落地",
    "主題偵測與智慧路由，自動轉接合適代理",
    "透過 Power Automate 串接 1,000+ 企業連接器",
    "自訂生成式 AI 外掛與提示詞管理",
    "多元發布管道：Teams、Web、WhatsApp 等",
    "對話分析儀表板，掌握互動品質與趨勢",
  ];
  slide.addText(features.map((f, i) => ({
    text: f, options: { bullet: true, breakLine: i < features.length - 1, paraSpaceAfter: 12 }
  })), { x: 0.5, y: 2.85, w: 7.5, h: 3.8, fontSize: 14, fontFace: "Calibri", color: C.ice, margin: 0 });

  // 右側數據卡
  const stats = [
    { value: "1,000+", label: "企業連接器", color: C.teal },
    { value: "0", label: "需寫程式碼", color: C.success },
    { value: "Multi", label: "發布管道", color: C.gold },
    { value: "1,000+", label: "已發布代理", color: C.coral },
  ];
  for (let i = 0; i < stats.length; i++) {
    const x = 9.0 + (i % 2) * 2.0;
    const y = 3.2 + Math.floor(i / 2) * 1.8;
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 1.85, h: 1.55, fill: { color: C.navy }, shadow: makeShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 1.85, h: 0.07, fill: { color: stats[i].color } });
    slide.addText(stats[i].value, { x, y: y + 0.2, w: 1.85, h: 0.7, fontSize: 22, fontFace: "Arial Black", color: stats[i].color, align: "center", margin: 0 });
    slide.addText(stats[i].label, { x, y: y + 0.95, w: 1.85, h: 0.45, fontSize: 10, fontFace: "Calibri", color: C.ice, align: "center", margin: 0 });
  }

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 7.1, w: 13.3, h: 0.4, fill: { color: C.navy } });
  slide.addText("Microsoft Agent Framework · Copilot Studio", { x: 0.5, y: 7.15, w: 9, h: 0.3, fontSize: 10, fontFace: "Calibri", color: C.ice, margin: 0 });
}

// ─── 第 7 頁：企業應用場景 ───
async function slide7() {
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  addHeader(slide, "企業應用場景");
  addFooter(slide);

  const cases = [
    { icon: FaLightbulb, title: "客戶服務代理", desc: "24/7 智慧支援，具備脈絡感知回應、無縫轉接真人與知識庫整合能力", color: C.teal, tag: "服務" },
    { icon: FaUsers, title: "人力資源", desc: "政策問答、請假申請、入職引導、IT 服務台自動化，大幅減少人事行政負擔", color: C.coral, tag: "HR" },
    { icon: FaChartLine, title: "業務與行銷", desc: "潛在客戶評比、CRM 更新、市場研究綜整、提案生成，加速銷售週期", color: C.gold, tag: "業務" },
    { icon: FaCode, title: "開發者生產力", desc: "程式碼審查、文件生成、CI/CD 故障排除、架構諮詢，工程團隊效率提升", color: C.success, tag: "工程" },
    { icon: FaShieldAlt, title: "合規與法務", desc: "合約分析、政策執行、稽核軌跡生成、法規報告撰寫，降低人為錯誤風險", color: C.purple, tag: "法遵" },
    { icon: FaRocket, title: "資料與分析", desc: "報表生成、KPI 監控，自然語言查詢資料倉儲，讓非技術人員也能自助分析", color: C.mid, tag: "資料" },
  ];

  for (let i = 0; i < cases.length; i++) {
    const row = Math.floor(i / 3);
    const col = i % 3;
    const x = 0.5 + col * 4.15;
    const y = 1.3 + row * 2.75;
    const c = cases[i];
    const iconData = await iconPng(c.icon, "#FFFFFF", 256);

    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.95, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
    // 頂部色塊
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.95, h: 0.85, fill: { color: c.color } });
    // Tag
    slide.addText(c.tag, { x: x + 3.15, y: y + 0.12, w: 0.65, h: 0.28, fontSize: 9, fontFace: "Calibri", color: C.white, align: "center", valign: "middle", fill: { color: "000000", transparency: 70 }, margin: 0 });
    // Icon
    slide.addImage({ data: iconData, x: x + 0.15, y: y + 0.12, w: 0.6, h: 0.6 });
    // 標題
    slide.addText(c.title, { x: x + 0.85, y: y + 0.18, w: 2.2, h: 0.55, fontSize: 13, fontFace: "Arial", color: C.white, bold: true, valign: "middle", margin: 0 });
    // 說明
    slide.addText(c.desc, { x: x + 0.2, y: y + 1.0, w: 3.55, h: 1.35, fontSize: 11, fontFace: "Calibri", color: C.mutedText, margin: 0 });
  }
}

// ─── 第 8 頁：快速入門 ───
async function slide8() {
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  addHeader(slide, "快速入門");
  addFooter(slide);

  const steps = [
    { num: "01", title: "擁有 Azure 訂閱", desc: "建立或使用現有 Azure 帳戶，並啟用 AI 服務訂用", icon: FaServer, color: C.teal },
    { num: "02", title: "選擇開發工具", desc: "專業開發用 Azure AI Studio，低程式碼用 Copilot Studio", icon: FaRocket, color: C.gold },
    { num: "03", title: "定義代理行為", desc: "選擇模型、定義工具、撰寫指示與對話流程", icon: FaBrain, color: C.coral },
    { num: "04", title: "測試與迭代", desc: "使用內建模擬器與劇本測試代理行為，調整提示詞直到滿意", icon: FaCheckCircle, color: C.success },
    { num: "05", title: "部署與監控", desc: "發布至目標管道並啟用完整可觀測性堆疊，確保持續優化", icon: FaSyncAlt, color: C.mid },
  ];

  for (let i = 0; i < steps.length; i++) {
    const y = 1.3 + i * 1.1;
    const s = steps[i];
    const iconData = await iconPng(s.icon, "#FFFFFF", 256);

    // 連接線
    if (i < steps.length - 1) {
      slide.addShape(pres.shapes.RECTANGLE, { x: 1.3, y: y + 0.85, w: 0.04, h: 0.25, fill: { color: s.color, transparency: 50 } });
    }

    // 號碼圓
    slide.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.05, w: 0.8, h: 0.8, fill: { color: s.color } });
    slide.addText(s.num, { x: 0.5, y: y + 0.05, w: 0.8, h: 0.8, fontSize: 20, fontFace: "Arial Black", color: C.white, align: "center", valign: "middle", margin: 0 });

    // 內容卡片
    slide.addShape(pres.shapes.RECTANGLE, { x: 1.55, y: y, w: 11.25, h: 0.9, fill: { color: C.white }, shadow: makeShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: 1.55, y: y, w: 0.08, h: 0.9, fill: { color: s.color } });
    slide.addImage({ data: iconData, x: 1.75, y: y + 0.15, w: 0.55, h: 0.55 });
    slide.addText(s.title, { x: 2.5, y: y + 0.08, w: 4, h: 0.4, fontSize: 14, fontFace: "Arial", color: C.navy, bold: true, margin: 0 });
    slide.addText(s.desc, { x: 2.5, y: y + 0.48, w: 10, h: 0.35, fontSize: 12, fontFace: "Calibri", color: C.mutedText, margin: 0 });
  }
}

// ─── 第 9 頁：總結 ───
async function slide9() {
  const slide = pres.addSlide();
  slide.background = { color: C.darkBg };

  // 裝飾圓
  slide.addShape(pres.shapes.OVAL, { x: -2, y: 3, w: 6, h: 6, fill: { color: C.navy, transparency: 45 }, line: { color: C.teal, width: 1.5 } });
  slide.addShape(pres.shapes.OVAL, { x: 10, y: -1, w: 5, h: 5, fill: { color: C.teal, transparency: 70 }, line: { color: C.accent, width: 1 } });
  slide.addShape(pres.shapes.OVAL, { x: 7, y: 5, w: 3, h: 3, fill: { color: C.mid, transparency: 65 } });

  // 標題
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.5, w: 0.1, h: 1.0, fill: { color: C.teal } });
  slide.addText("總結", { x: 0.8, y: 0.5, w: 6, h: 0.9, fontSize: 40, fontFace: "Arial Black", color: C.white, margin: 0 });

  // 重點卡片
  const points = [
    { icon: FaLayerGroup, text: "Microsoft Agent Framework = Azure AI Agent Service + Copilot Studio + AI Studio", color: C.teal },
    { icon: FaNetworkWired, text: "一次建構，發布至 Teams、Web、WhatsApp 等 1,000+ 管道", color: C.success },
    { icon: FaDatabase, text: "支援 600+ 模型：GPT-4、Llama、Phi 系列，以及自帶模型（BYOM）", color: C.gold },
    { icon: FaLock, text: "企業級安全：Entra ID、RBAC、SOC 2、ISO 27001 合規認證", color: C.coral },
    { icon: FaBolt, text: "從無程式碼（Copilot Studio）到專業程式碼（Azure AI Studio）的完整光譜", color: C.purple },
  ];

  for (let i = 0; i < points.length; i++) {
    const y = 1.7 + i * 0.88;
    const p = points[i];
    const iconData = await iconPng(p.icon, "#FFFFFF", 256);

    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 9.5, h: 0.72, fill: { color: C.navy }, shadow: makeShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 0.08, h: 0.72, fill: { color: p.color } });
    slide.addImage({ data: iconData, x: 0.7, y: y + 0.1, w: 0.5, h: 0.5 });
    slide.addText(p.text, { x: 1.4, y, w: 8.4, h: 0.72, fontSize: 13, fontFace: "Calibri", color: C.ice, valign: "middle", margin: 0 });
  }

  // 右側：學習資源
  slide.addShape(pres.shapes.RECTANGLE, { x: 10.3, y: 1.7, w: 2.7, h: 3.6, fill: { color: C.navy }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 10.3, y: 1.7, w: 2.7, h: 0.08, fill: { color: C.teal } });
  slide.addText("學習資源", { x: 10.45, y: 1.85, w: 2.4, h: 0.4, fontSize: 14, fontFace: "Arial", color: C.white, bold: true, margin: 0 });

  const links = [
    "Azure AI Agent Service",
    "Copilot Studio",
    "Azure AI Studio",
    "AI Foundry",
  ];
  slide.addText(links.map((l, i) => ({
    text: l, options: { bullet: true, breakLine: i < links.length - 1, paraSpaceAfter: 8 }
  })), { x: 10.45, y: 2.35, w: 2.4, h: 2.8, fontSize: 11, fontFace: "Calibri", color: C.ice, margin: 0 });

  slide.addText("有問題？歡迎討論 →", { x: 0.5, y: 6.6, w: 6, h: 0.4, fontSize: 14, fontFace: "Calibri", color: C.teal, margin: 0 });
}

// ─── 執行 ───
(async () => {
  console.log("產生中...");
  await slide1(); console.log("✓ 第1頁");
  await slide2(); console.log("✓ 第2頁");
  await slide3(); console.log("✓ 第3頁");
  await slide4(); console.log("✓ 第4頁");
  await slide5(); console.log("✓ 第5頁");
  await slide6(); console.log("✓ 第6頁");
  await slide7(); console.log("✓ 第7頁");
  await slide8(); console.log("✓ 第8頁");
  await slide9(); console.log("✓ 第9頁");
  await pres.writeFile({ fileName: "/home/wellcity/Microsoft_Agent_Framework.pptx" });
  console.log("✅ 完成");
})();
