const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// ── Colours ────────────────────────────────
const C = {
  navy:"0D1B2A", blue:"1A73E8", red:"EA4335",
  green:"34A853", yellow:"FBBC05", white:"FFFFFF",
  light:"F0F4FF", muted:"64748B", dark:"1E293B",
  card:"FFFFFF", slate:"334155",
};

// ✅ Safe image loader
function imgB64(imgPath) {
  const fullPath = path.join(__dirname, imgPath);
  if (!fs.existsSync(fullPath)) {
    console.error("❌ Missing file:", fullPath);
    process.exit(1);
  }
  return "image/png;base64," + fs.readFileSync(fullPath).toString("base64");
}

// Helpers
function contentBg(s){ s.background={color:"F4F7FF"}; }
function sectionBar(s,t){
  s.addShape("rect",{x:0,y:0,w:10,h:0.52,fill:{color:C.navy}});
  s.addText(t,{x:0.35,y:0.06,w:9.3,h:0.42,fontSize:14,bold:true,color:C.white});
}

// ──────────────────────────────────────────
let pres = new pptxgen();
pres.layout="LAYOUT_16x9";

// ── SLIDE 4 FIX (example) ──────────────────
{
  const s = pres.addSlide();
  contentBg(s);
  sectionBar(s,"User Profiles & Demographics");

  s.addImage({
    data: imgB64("charts/chart5_demographics.png"),
    x:0.25,y:0.6,w:9.5,h:4.8
  });
}

// ── FIX ALL OTHER IMAGES ───────────────────
const images = [
  ["charts/chart1_daily_sessions.png"],
  ["charts/chart2_sessions_dow.png"],
  ["charts/chart3_heatmap.png"],
  ["charts/chart4_top_channels.png"],
  ["charts/chart6_avg_duration.png"],
  ["charts/chart7_consumption_demo.png"],
  ["charts/chart9_province_channel.png"],
];

// Example usage (you already have slides — just fix paths like below)
// ❌ OLD: imgB64(imgB64("./chart3_heatmap.png"))
// ✅ NEW:
imgB64("charts/chart3_heatmap.png");

// ──────────────────────────────────────────
pres.writeFile({ fileName:"BrightTV_Viewership_Analytics.pptx" })
.then(()=>console.log("✅ Presentation created"))
.catch(e=>console.error(e));