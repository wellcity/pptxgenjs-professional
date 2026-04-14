// pptxgenjs 專業簡報範本
// 用法：修改內容 → node template.js → 產生 .pptx
// 依賴：pptxgenjs, react, react-dom, sharp, react-icons

const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaRocket, FaCheckCircle, FaServer } = require("react-icons/fa");

// ─── 配色 ───
const C = {
  navy:"1E2761",mid:"2E4A8F",teal:"0078D4",
  ice:"CADCFC",white:"FFFFFF",offWhite:"F5F7FA",
  darkText:"1A1A2E",mutedText:"5A6B8A",
  accent:"00B4D8",success:"10B981",
  coral:"E85D4C",gold:"F5A623",purple:"6B5CE7",
  darkBg:"0D1B3E",
};

// ─── 工具 ───
const makeShadow = () => ({ type:"outer",blur:10,offset:4,angle:135,color:"000000",opacity:0.22 });

async function iconPng(Icon, color="#0078D4", size=256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Icon, { color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

function addHeader(slide, title) {
  slide.addShape(pres.shapes.RECTANGLE, { x:0,y:0,w:13.3,h:1.1, fill:{color:C.navy} });
  slide.addShape(pres.shapes.RECTANGLE, { x:0,y:1.05,w:13.3,h:0.06, fill:{color:C.teal} });
  slide.addText(title, { x:0.5,y:0.25,w:12,h:0.7, fontSize:26, fontFace:"Arial", color:C.white, bold:true, margin:0 });
}

function addFooter(slide, text="Title · Subtitle") {
  slide.addShape(pres.shapes.RECTANGLE, { x:0,y:7.1,w:13.3,h:0.4, fill:{color:C.navy} });
  slide.addText(text, { x:0.5,y:7.15,w:9,h:0.3, fontSize:10, fontFace:"Calibri", color:C.ice, margin:0 });
  slide.addText("2026", { x:11.5,y:7.15,w:1.5,h:0.3, fontSize:10, fontFace:"Calibri", color:C.ice, align:"right", margin:0 });
}

function bulletList(slide, items, x, y, w, h, fontSize=13) {
  slide.addText(items.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<items.length-1,paraSpaceAfter:10}})),
    {x,y,w,h,fontSize,fontFace:"Calibri",color:C.darkText,margin:0});
}

// ─── 簡報設定 ───
let pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3" x 7.5"
pres.author = "Your Name";
pres.title = "Presentation Title";

// ─── 請在這裡修改內容 ───

// 第 1 頁：標題
async function slideTitle() {
  const slide = pres.addSlide();
  slide.background = { color: C.darkBg };
  // 裝飾元素、標題、副標題...（見 SKILL.md 完整範例）
  // TODO: 替換為你的內容
}

// 第 2 頁：內容頁
async function slideContent() {
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  addHeader(slide, "頁面標題");
  addFooter(slide);
  // TODO: 替換為你的內容
}

// ─── 執行 ───
(async () => {
  console.log("產生中...");
  await slideTitle();
  console.log("✓ 第1頁");
  await slideContent();
  console.log("✓ 第2頁");
  await pres.writeFile({ fileName: "/home/wellcity/output.pptx" });
  console.log("✅ 完成：/home/wellcity/output.pptx");
})();
