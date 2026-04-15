const pptxgen = require("pptxgenjs");
const fs = require("fs");

// ── Colours ────────────────────────────────
const C = {
  navy:    "0D1B2A",
  blue:    "1A73E8",
  red:     "EA4335",
  green:   "34A853",
  yellow:  "FBBC05",
  white:   "FFFFFF",
  light:   "F0F4FF",
  muted:   "64748B",
  dark:    "1E293B",
  card:    "FFFFFF",
  slate:   "334155",
};

function imgB64(path) {
  return "image/png;base64," + fs.readFileSync(path).toString("base64");
}

// ── Helper: content slide background ──────
function contentBg(slide) {
  slide.background = { color: "F4F7FF" };
}

// ── Helper: section header stripe ─────────
function sectionBar(slide, label) {
  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.52, fill: { color: C.navy } });
  slide.addText(label, {
    x: 0.35, y: 0.06, w: 9.3, h: 0.42,
    fontSize: 14, bold: true, color: C.white, valign: "middle", margin: 0,
  });
}

// ── Helper: stat card ─────────────────────
function statCard(slide, x, y, w, h, value, label, accent) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: C.white },
    shadow: { type: "outer", blur: 8, offset: 2, angle: 135, color: "000000", opacity: 0.08 }
  });
  slide.addShape("rect", { x, y, w: 0.07, h, fill: { color: accent } });
  slide.addText(value, { x: x+0.15, y: y+0.08, w: w-0.2, h: h*0.55, fontSize: 26, bold: true, color: accent, valign: "middle", margin: 0 });
  slide.addText(label, { x: x+0.15, y: y+h*0.55, w: w-0.2, h: h*0.42, fontSize: 10, color: C.muted, valign: "top", margin: 0 });
}

// ── Helper: bullet list ───────────────────
function bulletList(slide, items, x, y, w, h, color) {
  const runs = items.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < items.length-1, fontSize: 13, color: C.dark }
  }));
  slide.addText(runs, { x, y, w, h, valign: "top", margin: 4 });
}

// ──────────────────────────────────────────
// BUILD PRESENTATION
// ──────────────────────────────────────────
let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title  = "BrightTV Viewership Analytics";

// ══════════════════════════════════════════
// SLIDE 1 – TITLE
// ══════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.navy };
  // Accent bar left
  s.addShape("rect", { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.blue } });
  // Title
  s.addText("BrightTV", { x: 0.55, y: 1.1, w: 9, h: 0.9, fontSize: 52, bold: true, color: C.white });
  s.addText("Viewership Analytics", { x: 0.55, y: 1.95, w: 9, h: 0.65, fontSize: 28, color: "93C5FD", bold: false });
  s.addText("Customer Value Management — Growth Strategy Presentation", {
    x: 0.55, y: 2.75, w: 9, h: 0.4, fontSize: 14, color: "CBD5E1", italic: true
  });
  // Divider
  s.addShape("rect", { x: 0.55, y: 3.3, w: 3, h: 0.04, fill: { color: C.blue } });
  // Date & notes
  s.addText("Q1 2016  |  Jan – Mar  |  10,000 Sessions Analysed", {
    x: 0.55, y: 3.55, w: 9, h: 0.35, fontSize: 12, color: "94A3B8"
  });
  // Bottom tag
  s.addShape("rect", { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.blue } });
  s.addText("CONFIDENTIAL  •  BrightTV CVM Team", {
    x: 0, y: 5.25, w: 10, h: 0.375, fontSize: 10, color: C.white, align: "center", valign: "middle", margin: 0
  });
}

// ══════════════════════════════════════════
// SLIDE 2 – AGENDA
// ══════════════════════════════════════════
{
  const s = pres.addSlide();
  contentBg(s);
  sectionBar(s, "Agenda");

  const items = [
    ["01", "Executive Summary & KPIs",         C.blue],
    ["02", "User Profiles & Demographics",      C.green],
    ["03", "Viewership Trends & Usage Patterns",C.yellow],
    ["04", "Factors Influencing Consumption",   C.red],
    ["05", "Content Recommendations",           "9C27B0"],
    ["06", "Growth Initiatives",                "00BCD4"],
  ];

  const cols = [[items[0],items[1],items[2]], [items[3],items[4],items[5]]];
  cols.forEach((col, ci) => {
    col.forEach((item, ri) => {
      const x = 0.45 + ci*4.9;
      const y = 0.75 + ri*1.5;
      s.addShape("rect", {
        x, y, w: 4.5, h: 1.25,
        fill: { color: C.white },
        shadow: { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.08 }
      });
      s.addShape("rect", { x, y, w: 0.07, h: 1.25, fill: { color: item[2] } });
      s.addText(item[0], { x: x+0.15, y: y+0.08, w: 0.55, h: 0.45, fontSize: 22, bold: true, color: item[2], margin: 0 });
      s.addText(item[1], { x: x+0.15, y: y+0.5, w: 4.15, h: 0.65, fontSize: 13, color: C.dark, valign: "top", margin: 0 });
    });
  });
}

// ══════════════════════════════════════════
// SLIDE 3 – EXECUTIVE SUMMARY KPIs
// ══════════════════════════════════════════
{
  const s = pres.addSlide();
  contentBg(s);
  sectionBar(s, "Executive Summary – Key Performance Indicators");

  const stats = [
    ["9,738",   "Total Sessions (Q1 2016)",        C.blue],
    ["4,211",   "Unique Active Subscribers",        C.green],
    ["5,375",   "Total Registered Users",           "9C27B0"],
    ["1,487 h", "Total Watch Time",                 C.red],
    ["9.2 min", "Avg Session Duration",             C.yellow],
    ["21",      "Channels Available",               "00BCD4"],
  ];

  stats.forEach((st, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    statCard(s, 0.35 + col*3.15, 0.68 + row*2.3, 2.95, 2.05, st[0], st[1], st[2]);
  });
}

// ══════════════════════════════════════════
// SLIDE 4 – USER DEMOGRAPHICS
// ══════════════════════════════════════════
{
  const s = pres.addSlide();
  contentBg(s);
  sectionBar(s, "User Profiles & Demographics");
  s.addImage({ data: imgB64("/home/claude/chart5_demographics.png"), x: 0.25, y: 0.6, w: 9.5, h: 4.8 });
}

// ══════════════════════════════════════════
// SLIDE 5 – DAILY SESSION TREND
// ══════════════════════════════════════════
{
  const s = pres.addSlide();
  contentBg(s);
  sectionBar(s, "Viewership Trends – Daily Sessions (SAST)");
  s.addImage({ data: imgB64("/home/claude/chart1_daily_sessions.png"), x: 0.3, y: 0.62, w: 9.4, h: 4.6 });
  s.addText("Shaded bars indicate weekends. ICC Cricket World Cup 2011 matches drove spikes in March 2016.",
    { x: 0.3, y: 5.2, w: 9.4, h: 0.35, fontSize: 10, color: C.muted, italic: true });
}

// ══════════════════════════════════════════
// SLIDE 6 – SESSIONS BY DAY OF WEEK
// ══════════════════════════════════════════
{
  const s = pres.addSlide();
  contentBg(s);
  sectionBar(s, "Viewership Trends – Sessions by Day of Week");

  s.addImage({ data: imgB64("/home/claude/chart2_sessions_dow.png"), x: 0.3, y: 0.62, w: 5.8, h: 4.65 });

  // Insight cards right side
  const insights = [
    [C.blue,  "Friday Peak",    "Friday generates the highest sessions (1,634), suggesting subscribers unwind after work/school."],
    [C.red,   "Monday Slump",   "Monday is the lowest day (928 sessions) — nearly 43% fewer than Friday."],
    [C.green, "Weekend Steady", "Weekends show moderate but consistent activity — a key engagement window."],
  ];
  insights.forEach((ins, i) => {
    const y = 0.72 + i*1.57;
    s.addShape("rect", {
      x: 6.3, y, w: 3.4, h: 1.4,
      fill: { color: C.white },
      shadow: { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.08 }
    });
    s.addShape("rect", { x: 6.3, y, w: 0.07, h: 1.4, fill: { color: ins[0] } });
    s.addText(ins[1], { x: 6.45, y: y+0.1, w: 3.1, h: 0.35, fontSize: 12, bold: true, color: C.dark, margin: 0 });
    s.addText(ins[2], { x: 6.45, y: y+0.45, w: 3.1, h: 0.85, fontSize: 10, color: C.muted, valign: "top", margin: 0 });
  });
}

// ══════════════════════════════════════════
// SLIDE 7 – HOURLY HEATMAP
// ══════════════════════════════════════════
{
  const s = pres.addSlide();
  contentBg(s);
  sectionBar(s, "Viewing Intensity – Day × Hour Heatmap (SAST)");
  s.addImage({ data: imgB64("/home/claude/chart3_heatmap.png"), x: 0.25, y: 0.62, w: 9.5, h: 4.65 });
  s.addText("Peak viewing: 17:00–21:00 SAST across all days. Weekend mornings (09:00–12:00) also show elevated activity.",
    { x: 0.3, y: 5.2, w: 9.4, h: 0.35, fontSize: 10, color: C.muted, italic: true });
}

// ══════════════════════════════════════════
// SLIDE 8 – TOP CHANNELS
// ══════════════════════════════════════════
{
  const s = pres.addSlide();
  contentBg(s);
  sectionBar(s, "Top Channels by Sessions & Watch Time");
  s.addImage({ data: imgB64("/home/claude/chart4_top_channels.png"), x: 0.25, y: 0.62, w: 9.5, h: 4.7 });
}

// ══════════════════════════════════════════
// SLIDE 9 – AVG DURATION BY CHANNEL
// ══════════════════════════════════════════
{
  const s = pres.addSlide();
  contentBg(s);
  sectionBar(s, "Average Session Duration by Channel");
  s.addImage({ data: imgB64("/home/claude/chart6_avg_duration.png"), x: 0.25, y: 0.62, w: 9.5, h: 4.7 });
  s.addText("Channels with longer avg sessions indicate deeper engagement — useful for premium packaging decisions.",
    { x: 0.3, y: 5.2, w: 9.4, h: 0.35, fontSize: 10, color: C.muted, italic: true });
}

// ══════════════════════════════════════════
// SLIDE 10 – CONSUMPTION BY DEMOGRAPHICS
// ══════════════════════════════════════════
{
  const s = pres.addSlide();
  contentBg(s);
  sectionBar(s, "Factors Influencing Consumption – Age Group & Gender");
  s.addImage({ data: imgB64("/home/claude/chart7_consumption_demo.png"), x: 0.25, y: 0.62, w: 9.5, h: 4.7 });
}

// ══════════════════════════════════════════
// SLIDE 11 – PROVINCE CHANNEL HEATMAP
// ══════════════════════════════════════════
{
  const s = pres.addSlide();
  contentBg(s);
  sectionBar(s, "Factors Influencing Consumption – Geography & Channel Preference");
  s.addImage({ data: imgB64("/home/claude/chart9_province_channel.png"), x: 0.25, y: 0.62, w: 9.5, h: 4.7 });
}

// ══════════════════════════════════════════
// SLIDE 12 – FACTORS SUMMARY
// ══════════════════════════════════════════
{
  const s = pres.addSlide();
  contentBg(s);
  sectionBar(s, "Key Factors Influencing Consumption");

  const factors = [
    [C.blue,   "Time of Day",    "Evening prime time (17:00–21:00 SAST) drives the majority of sessions. Lunch hour (12:00–13:00) shows a secondary spike."],
    [C.green,  "Content Type",   "Sport (SuperSport Live Events, ICC Cricket) and music/entertainment (Channel O, Trace TV, Africa Magic) dominate watch time."],
    [C.yellow, "Demographics",   "Males aged 25–44 account for the highest consumption. Female engagement is significantly under-represented."],
    [C.red,    "Day of Week",    "Friday is the clear peak day; Monday is the trough. Targeted promotions on low days could lift volumes by ~40%."],
    ["9C27B0", "Geography",      "Gauteng subscribers generate the highest activity. Rural provinces (Northern Cape, Free State) show untapped potential."],
    ["00BCD4", "Event Seasons",  "ICC Cricket World Cup viewership spikes confirm that live events are powerful consumption drivers."],
  ];

  factors.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.3 + col * 4.85;
    const y = 0.68 + row * 1.6;
    s.addShape("rect", {
      x, y, w: 4.55, h: 1.45,
      fill: { color: C.white },
      shadow: { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.08 }
    });
    s.addShape("rect", { x, y, w: 0.07, h: 1.45, fill: { color: f[0] } });
    s.addText(f[1], { x: x+0.15, y: y+0.08, w: 4.25, h: 0.35, fontSize: 12, bold: true, color: C.dark, margin: 0 });
    s.addText(f[2], { x: x+0.15, y: y+0.45, w: 4.25, h: 0.9, fontSize: 10, color: C.muted, valign: "top", margin: 0 });
  });
}

// ══════════════════════════════════════════
// SLIDE 13 – CONTENT RECOMMENDATIONS
// ══════════════════════════════════════════
{
  const s = pres.addSlide();
  contentBg(s);
  sectionBar(s, "Content Recommendations for Low-Consumption Days (Mon–Wed)");

  const recs = [
    [C.blue,  "🏋 Fitness & Wellness",  "Monday Motivation",  "Workout series, yoga, health documentaries — align with the 'new week' mindset. Target 06:00–08:00 SAST morning slot."],
    [C.green, "🎬 Binge-Ready Series",  "Series Drop Strategy","Release new episodes of drama/sitcom series on Monday evenings to incentivise midweek logins."],
    [C.yellow,"🎮 Gaming & eSports",    "Youth Midweek Hook", "eSports tournaments and gaming content attract the under-25 male segment during slow afternoon hours."],
    [C.red,   "📚 Education & Docu",   "Inform & Engage",     "Documentary marathons and current affairs draw the 35–54 age group on Tuesday/Wednesday evenings."],
    ["9C27B0","🎵 Music Countdowns",   "Chart Shows",         "Weekly Top-40 and Afrobeats countdowns on Channel O / Trace TV — low-production, high-appeal for under-30s."],
    ["00BCD4","🌍 Local Content",      "SA Stories",          "Local South African content (reality shows, local drama) drives cultural connection and loyalty mid-week."],
  ];

  recs.forEach((r, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.3 + col * 4.85;
    const y = 0.68 + row * 1.6;
    s.addShape("rect", {
      x, y, w: 4.55, h: 1.45,
      fill: { color: C.white },
      shadow: { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.08 }
    });
    s.addShape("rect", { x, y, w: 0.07, h: 1.45, fill: { color: r[0] } });
    s.addText(r[1] + "  —  " + r[2], { x: x+0.15, y: y+0.08, w: 4.25, h: 0.38, fontSize: 11, bold: true, color: C.dark, margin: 0 });
    s.addText(r[3], { x: x+0.15, y: y+0.48, w: 4.25, h: 0.88, fontSize: 10, color: C.muted, valign: "top", margin: 0 });
  });
}

// ══════════════════════════════════════════
// SLIDE 14 – GROWTH INITIATIVES
// ══════════════════════════════════════════
{
  const s = pres.addSlide();
  contentBg(s);
  sectionBar(s, "Growth Initiatives – Expanding the Subscriber Base");

  const initiatives = [
    {
      n: "01", color: C.blue,
      title: "Female Subscriber Drive",
      detail: "Only ~10% of subscribers are female. Launch targeted campaigns via social media (Instagram/TikTok) featuring lifestyle, reality, and local drama content. Partner with female influencers."
    },
    {
      n: "02", color: C.green,
      title: "Provincial Expansion (Rural SA)",
      detail: "Northern Cape, Free State and North West are under-indexed. Deploy affordable mobile-data bundles in partnership with MTN/Vodacom. Offer reduced-rate subscriptions for townships."
    },
    {
      n: "03", color: C.red,
      title: "Youth Acquisition (Under-18)",
      detail: "The under-18 segment shows low engagement. Launch a 'BrightTV Kids & Youth' tier with Cartoon Network, Boomerang and eSports. Offer student pricing and school holiday campaigns."
    },
    {
      n: "04", color: C.yellow,
      title: "Loyalty & Retention Programme",
      detail: "Reward long-session viewers with points redeemable for subscription discounts. A 'Streak Reward' (7-day consecutive login) can increase MAU and reduce churn."
    },
    {
      n: "05", color: "9C27B0",
      title: "Live Events & Sports Packaging",
      detail: "SuperSport & Cricket dominate consumption. Create sports season passes and live-event bundles to convert casual viewers into paid subscribers during peak sports seasons."
    },
    {
      n: "06", color: "00BCD4",
      title: "Social Sharing & Referral Engine",
      detail: "Leverage existing social media handles in the database. Build share-to-earn referral mechanics — reward subscribers for bringing in friends via their linked handles."
    },
  ];

  initiatives.forEach((ini, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.3 + col * 4.85;
    const y = 0.68 + row * 1.6;
    s.addShape("rect", {
      x, y, w: 4.55, h: 1.45,
      fill: { color: C.white },
      shadow: { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.08 }
    });
    s.addShape("rect", { x, y, w: 0.07, h: 1.45, fill: { color: ini.color } });
    s.addText(ini.n + "  " + ini.title, { x: x+0.15, y: y+0.08, w: 4.25, h: 0.38, fontSize: 11, bold: true, color: C.dark, margin: 0 });
    s.addText(ini.detail, { x: x+0.15, y: y+0.48, w: 4.25, h: 0.88, fontSize: 10, color: C.muted, valign: "top", margin: 0 });
  });
}

// ══════════════════════════════════════════
// SLIDE 15 – CLOSING
// ══════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape("rect", { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.blue } });
  s.addText("Key Takeaways", { x: 0.55, y: 0.7, w: 9, h: 0.65, fontSize: 34, bold: true, color: C.white });
  s.addShape("rect", { x: 0.55, y: 1.45, w: 2.5, h: 0.04, fill: { color: C.blue } });

  const takeaways = [
    "Friday evenings are peak — protect and enhance this window with premium content",
    "Sport & Music are the two content pillars driving the most engagement",
    "Female and rural subscribers are the largest untapped growth segments",
    "Monday–Wednesday need dedicated content strategies to lift low-day sessions",
    "Live events (cricket, football) are powerful subscriber acquisition tools",
    "A referral + loyalty programme can accelerate organic subscriber growth",
  ];

  takeaways.forEach((t, i) => {
    const y = 1.65 + i * 0.58;
    s.addShape("oval", { x: 0.55, y: y+0.09, w: 0.28, h: 0.28, fill: { color: C.blue } });
    s.addText(String(i+1), { x: 0.55, y: y+0.09, w: 0.28, h: 0.28, fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
    s.addText(t, { x: 0.95, y: y+0.04, w: 8.7, h: 0.45, fontSize: 13, color: "CBD5E1", valign: "middle", margin: 0 });
  });

  s.addShape("rect", { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.blue } });
  s.addText("Thank you  •  BrightTV CVM Analytics  •  2016", {
    x: 0, y: 5.25, w: 10, h: 0.375, fontSize: 11, color: C.white, align: "center", valign: "middle", margin: 0
  });
}

// ──────────────────────────────────────────
// WRITE FILE
// ──────────────────────────────────────────
pres.writeFile({ fileName: "BrightTV_Viewership_Analytics.pptx" })
  .then(() => console.log("✅ Presentation saved: BrightTV_Viewership_Analytics.pptx"))
  .catch(e => { console.error("❌", e); process.exit(1); });
