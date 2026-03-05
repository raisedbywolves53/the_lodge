const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// ============ BRAND CONSTANTS ============
const DEEP_TEAL = "366F78";
const MIDNIGHT = "1B262C";
const BLACK = "000000";
const WHITE = "FFFFFF";
const WARM_WHITE = "F7F5F2";     // warm off-white for light slides
const LIGHT_TEAL = "D6EBED";     // soft teal tint
const ACCENT_GOLD = "C9A96E";    // warm gold accent
const CARD_BG = "233840";        // slightly lighter than midnight for cards on dark
const BODY_LIGHT = "E8ECF0";     // readable light text on dark backgrounds

const HEADER_FONT = "Trebuchet MS";
const BODY_FONT = "Calibri";

// Logo paths
const logoCleanPath = "/sessions/admiring-ecstatic-hypatia/logo_transparent.png";
const logoDarkPath = "/sessions/admiring-ecstatic-hypatia/logo_dark_v2.png";

const logoClean = "image/png;base64," + fs.readFileSync(logoCleanPath).toString("base64");
const logoDark = "image/png;base64," + fs.readFileSync(logoDarkPath).toString("base64");

// Icon generation
const { FaUsers, FaHeart, FaBrain, FaRobot, FaEnvelope, FaDollarSign, FaCalendarAlt,
  FaCheckCircle, FaHome, FaBook, FaHandsHelping, FaCompass, FaRocket, FaShieldAlt,
  FaClock, FaComments, FaDatabase, FaServer, FaCog, FaGlobe, FaPaintBrush,
  FaLightbulb, FaStar, FaArrowRight, FaQuoteLeft, FaSearch, FaSitemap, FaPen,
  FaChartLine, FaLaptop, FaPhoneAlt, FaPlay, FaUserFriends, FaRegHandshake,
  FaMapMarkerAlt, FaMedal } = require("react-icons/fa");

function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

// Helper: fresh shadow factory (never reuse objects)
const makeShadow = () => ({ type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.12 });

async function buildPresentation() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Still Mind Creative";
  pres.title = "Lodge360 — A New Digital Home";

  // Pre-render icons - WHITE for dark backgrounds
  const iconUsersW = await iconToBase64Png(FaUsers, "#FFFFFF", 256);
  const iconHeartW = await iconToBase64Png(FaHeart, "#FFFFFF", 256);
  const iconBrainW = await iconToBase64Png(FaBrain, "#FFFFFF", 256);
  const iconRobotW = await iconToBase64Png(FaRobot, "#FFFFFF", 256);
  const iconEnvelopeW = await iconToBase64Png(FaEnvelope, "#FFFFFF", 256);
  const iconCheckW = await iconToBase64Png(FaCheckCircle, "#FFFFFF", 256);
  const iconHomeW = await iconToBase64Png(FaHome, "#FFFFFF", 256);
  const iconBookW = await iconToBase64Png(FaBook, "#FFFFFF", 256);
  const iconHandsW = await iconToBase64Png(FaHandsHelping, "#FFFFFF", 256);
  const iconCompassW = await iconToBase64Png(FaCompass, "#FFFFFF", 256);
  const iconRocketW = await iconToBase64Png(FaRocket, "#FFFFFF", 256);
  const iconShieldW = await iconToBase64Png(FaShieldAlt, "#FFFFFF", 256);
  const iconClockW = await iconToBase64Png(FaClock, "#FFFFFF", 256);
  const iconCommentsW = await iconToBase64Png(FaComments, "#FFFFFF", 256);
  const iconSearchW = await iconToBase64Png(FaSearch, "#FFFFFF", 256);
  const iconPhoneW = await iconToBase64Png(FaPhoneAlt, "#FFFFFF", 256);
  const iconGlobeW = await iconToBase64Png(FaGlobe, "#FFFFFF", 256);
  const iconStarW = await iconToBase64Png(FaStar, "#FFFFFF", 256);
  const iconLightbulbW = await iconToBase64Png(FaLightbulb, "#FFFFFF", 256);
  const iconDBW = await iconToBase64Png(FaDatabase, "#FFFFFF", 256);
  const iconCogW = await iconToBase64Png(FaCog, "#FFFFFF", 256);
  const iconLaptopW = await iconToBase64Png(FaLaptop, "#FFFFFF", 256);
  const iconChartW = await iconToBase64Png(FaChartLine, "#FFFFFF", 256);
  const iconPlayW = await iconToBase64Png(FaPlay, "#FFFFFF", 256);
  const iconPenW = await iconToBase64Png(FaPen, "#FFFFFF", 256);
  const iconSitemapW = await iconToBase64Png(FaSitemap, "#FFFFFF", 256);
  const iconMedalW = await iconToBase64Png(FaMedal, "#FFFFFF", 256);

  // TEAL icons for light backgrounds
  const iconUsersT = await iconToBase64Png(FaUsers, "#366F78", 256);
  const iconHeartT = await iconToBase64Png(FaHeart, "#366F78", 256);
  const iconCheckT = await iconToBase64Png(FaCheckCircle, "#366F78", 256);
  const iconSearchT = await iconToBase64Png(FaSearch, "#366F78", 256);
  const iconCompassT = await iconToBase64Png(FaCompass, "#366F78", 256);
  const iconRocketT = await iconToBase64Png(FaRocket, "#366F78", 256);
  const iconClockT = await iconToBase64Png(FaClock, "#366F78", 256);
  const iconShieldT = await iconToBase64Png(FaShieldAlt, "#366F78", 256);
  const iconStarT = await iconToBase64Png(FaStar, "#366F78", 256);
  const iconHandsT = await iconToBase64Png(FaHandsHelping, "#366F78", 256);
  const iconBrainT = await iconToBase64Png(FaBrain, "#366F78", 256);
  const iconPhoneT = await iconToBase64Png(FaPhoneAlt, "#366F78", 256);
  const iconGlobeT = await iconToBase64Png(FaGlobe, "#366F78", 256);
  const iconRobotT = await iconToBase64Png(FaRobot, "#366F78", 256);
  const iconEnvelopeT = await iconToBase64Png(FaEnvelope, "#366F78", 256);
  const iconBookT = await iconToBase64Png(FaBook, "#366F78", 256);

  // GOLD icons for accent
  const iconQuoteGold = await iconToBase64Png(FaQuoteLeft, "#C9A96E", 256);

  // ============================================================
  // SLIDE 1: TITLE
  // ============================================================
  let s1 = pres.addSlide();
  s1.background = { color: MIDNIGHT };

  // Subtle top accent line
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: DEEP_TEAL } });

  // Logo centered
  s1.addImage({ data: logoDark, x: 2.8, y: 0.4, w: 4.4, h: 2.5, sizing: { type: "contain", w: 4.4, h: 2.5 } });

  s1.addText("Lodge360", {
    x: 0.5, y: 2.9, w: 9, h: 0.9,
    fontSize: 48, fontFace: HEADER_FONT, color: WHITE, bold: true, align: "center", margin: 0
  });

  s1.addText("A New Digital Home", {
    x: 0.5, y: 3.7, w: 9, h: 0.5,
    fontSize: 22, fontFace: BODY_FONT, color: ACCENT_GOLD, align: "center", margin: 0
  });

  s1.addText("Website  \u00B7  AI Assistant  \u00B7  Email Strategy", {
    x: 0.5, y: 4.2, w: 9, h: 0.35,
    fontSize: 13, fontFace: BODY_FONT, color: BODY_LIGHT, align: "center", margin: 0
  });

  s1.addText("Prepared for Jim & Adrienne Tichy  \u00B7  March 2026", {
    x: 0.5, y: 5.0, w: 9, h: 0.3,
    fontSize: 11, fontFace: BODY_FONT, color: BODY_LIGHT, align: "center", margin: 0
  });
  s1.addText("Still Mind Creative", {
    x: 0.5, y: 5.25, w: 9, h: 0.25,
    fontSize: 10, fontFace: BODY_FONT, color: DEEP_TEAL, align: "center", margin: 0
  });

  // ============================================================
  // SLIDE 2: WHO THIS SITE IS FOR (shorter text, bullets)
  // ============================================================
  let s2 = pres.addSlide();
  s2.background = { color: WARM_WHITE };

  s2.addText("Who This Site Is For", {
    x: 0.6, y: 0.4, w: 8.8, h: 0.7,
    fontSize: 36, fontFace: HEADER_FONT, color: MIDNIGHT, bold: true, margin: 0
  });
  s2.addText("Before we talk about what to build, we need to talk about who we\u2019re building it for.", {
    x: 0.6, y: 1.1, w: 8.8, h: 0.35,
    fontSize: 13, fontFace: BODY_FONT, color: DEEP_TEAL, italic: true, margin: 0
  });

  // Card 1: The Parent
  s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.7, w: 4.25, h: 3.5, fill: { color: WHITE }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.7, w: 0.07, h: 3.5, fill: { color: DEEP_TEAL } });
  s2.addImage({ data: iconUsersT, x: 0.85, y: 1.9, w: 0.4, h: 0.4 });
  s2.addText("The Parent", {
    x: 1.35, y: 1.9, w: 3.2, h: 0.4,
    fontSize: 20, fontFace: HEADER_FONT, color: MIDNIGHT, bold: true, margin: 0
  });

  s2.addText([
    { text: "Financially comfortable, emotionally exhausted", options: { bullet: true, breakLine: true } },
    { text: "Paid for multiple treatment stays with no lasting results", options: { bullet: true, breakLine: true } },
    { text: "Searching in a state of crisis \u2014 scared, angry, or both", options: { bullet: true, breakLine: true } },
    { text: "Skeptical of promises after past failures", options: { bullet: true, breakLine: true } },
    { text: "Found Lodge360 through someone they trust", options: { bullet: true } },
  ], {
    x: 0.85, y: 2.5, w: 3.65, h: 2.5,
    fontSize: 11.5, fontFace: BODY_FONT, color: "444444", margin: 0, valign: "top", paraSpaceAfter: 6
  });

  // Card 2: The Client
  s2.addShape(pres.shapes.RECTANGLE, { x: 5.25, y: 1.7, w: 4.25, h: 3.5, fill: { color: WHITE }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: 5.25, y: 1.7, w: 0.07, h: 3.5, fill: { color: ACCENT_GOLD } });
  s2.addImage({ data: iconHeartT, x: 5.6, y: 1.9, w: 0.4, h: 0.4 });
  s2.addText("The Client", {
    x: 6.1, y: 1.9, w: 3.2, h: 0.4,
    fontSize: 20, fontFace: HEADER_FONT, color: MIDNIGHT, bold: true, margin: 0
  });

  s2.addText([
    { text: "Struggling with addiction, alcoholism, or mental health", options: { bullet: true, breakLine: true } },
    { text: "Been through the treatment cycle multiple times", options: { bullet: true, breakLine: true } },
    { text: "Chronic relapse or failure to launch", options: { bullet: true, breakLine: true } },
    { text: "Treatment gave them tools \u2014 Lodge360 teaches them to use those tools in real life", options: { bullet: true } },
  ], {
    x: 5.6, y: 2.5, w: 3.65, h: 2.5,
    fontSize: 11.5, fontFace: BODY_FONT, color: "444444", margin: 0, valign: "top", paraSpaceAfter: 6
  });

  // ============================================================
  // SLIDE 3: WHAT THESE FAMILIES ARE FEELING
  // ============================================================
  let s3 = pres.addSlide();
  s3.background = { color: MIDNIGHT };

  s3.addText("What These Families Are Feeling", {
    x: 0.6, y: 0.25, w: 8.8, h: 0.5,
    fontSize: 30, fontFace: HEADER_FONT, color: WHITE, bold: true, margin: 0
  });
  s3.addText("Your site must meet them here \u2014 not where we wish they were.", {
    x: 0.6, y: 0.8, w: 8.8, h: 0.35,
    fontSize: 13, fontFace: BODY_FONT, color: ACCENT_GOLD, italic: true, margin: 0
  });

  const feelings = [
    { quote: "\u201CWe\u2019ve spent so much money and nothing has worked.\u201D", label: "Frustration & Skepticism" },
    { quote: "\u201CI don\u2019t even know what to look for anymore.\u201D", label: "Overwhelm & Decision Fatigue" },
    { quote: "\u201CI\u2019m terrified something will happen to them.\u201D", label: "Fear & Desperation" },
    { quote: "\u201CSometimes I want to save them and give up at the same time.\u201D", label: "Anger & Love" }
  ];

  feelings.forEach((f, i) => {
    const yBase = 1.55 + i * 0.92;
    s3.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: yBase, w: 8.8, h: 0.78, fill: { color: CARD_BG } });
    s3.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: yBase, w: 0.06, h: 0.78, fill: { color: ACCENT_GOLD } });
    s3.addText(f.quote, {
      x: 0.95, y: yBase + 0.08, w: 6, h: 0.32,
      fontSize: 14, fontFace: BODY_FONT, color: WHITE, italic: true, margin: 0
    });
    s3.addText(f.label, {
      x: 0.95, y: yBase + 0.42, w: 6, h: 0.28,
      fontSize: 11, fontFace: BODY_FONT, color: DEEP_TEAL, bold: true, margin: 0
    });
  });

  s3.addText("The site doesn\u2019t need to sell. It needs to say: \u201CWe know. We\u2019ve been here. Here\u2019s why this time is different.\u201D", {
    x: 1.2, y: 5.1, w: 7.6, h: 0.4,
    fontSize: 11.5, fontFace: BODY_FONT, color: BODY_LIGHT, italic: true, margin: 0, align: "center"
  });

  // ============================================================
  // SLIDE 4: WHY FAMILIES WILL TRUST YOU (pyramid, "you" language)
  // ============================================================
  let s4 = pres.addSlide();
  s4.background = { color: WARM_WHITE };

  s4.addText("Why Families Will Trust You", {
    x: 0.6, y: 0.35, w: 8.8, h: 0.7,
    fontSize: 36, fontFace: HEADER_FONT, color: MIDNIGHT, bold: true, margin: 0
  });
  s4.addText("Your site needs to communicate these truths \u2014 without sounding like marketing.", {
    x: 0.6, y: 1.05, w: 8.8, h: 0.35,
    fontSize: 13, fontFace: BODY_FONT, color: DEEP_TEAL, italic: true, margin: 0
  });

  // PYRAMID: 3 top, 2 bottom (centered)
  const trustItems = [
    { title: "You\u2019ve both lived it", desc: "Jim and Adrienne\u2019s personal recovery journeys are why this place exists. They built Lodge360 from lived experience, not theory.", icon: iconHeartT },
    { title: "18 years. 9 homes. A waiting list.", desc: "Since 2008, you\u2019ve never needed marketing because word of mouth has been enough. The reputation is already built.", icon: iconStarT },
    { title: "Accountability over revenue", desc: "You\u2019ll remove someone to protect the culture. In an industry that keeps people enrolled for cash flow, you do the opposite.", icon: iconShieldT },
    { title: "This is everything to you", desc: "No side businesses. No distractions. You pour everything into this community and it shows.", icon: iconHandsT },
    { title: "12-step, lived daily", desc: "The Lodge was built on the principles and traditions of AA \u2014 not as a marketing angle, but as an operating philosophy.", icon: iconCompassT },
  ];

  // Row 1: 3 cards centered, Row 2: 2 cards centered (tighter gap)
  const cardW = 2.85;
  const cardH = 1.55;
  const gap = 0.2;
  const pyramidGap = 0.15; // tighter vertical gap between rows
  const row1Total = cardW * 3 + gap * 2;
  const row1Start = (10 - row1Total) / 2;
  const row2Total = cardW * 2 + gap;
  const row2Start = (10 - row2Total) / 2;

  trustItems.forEach((item, i) => {
    let xBase, yBase;
    if (i < 3) {
      xBase = row1Start + i * (cardW + gap);
      yBase = 1.6;
    } else {
      xBase = row2Start + (i - 3) * (cardW + gap);
      yBase = 1.6 + cardH + pyramidGap;
    }

    s4.addShape(pres.shapes.RECTANGLE, { x: xBase, y: yBase, w: cardW, h: cardH, fill: { color: WHITE }, shadow: makeShadow() });
    s4.addImage({ data: item.icon, x: xBase + 0.15, y: yBase + 0.15, w: 0.32, h: 0.32 });
    s4.addText(item.title, {
      x: xBase + 0.55, y: yBase + 0.12, w: cardW - 0.75, h: 0.35,
      fontSize: 12, fontFace: HEADER_FONT, color: MIDNIGHT, bold: true, margin: 0
    });
    s4.addText(item.desc, {
      x: xBase + 0.15, y: yBase + 0.55, w: cardW - 0.3, h: 0.9,
      fontSize: 10.5, fontFace: BODY_FONT, color: "555555", margin: 0, valign: "top"
    });
  });

  // ============================================================
  // SLIDE 5: WHAT THE SITE NEEDS TO DO (complete redesign)
  // ============================================================
  let s5 = pres.addSlide();
  s5.background = { color: WARM_WHITE };

  // Full-width dark header section
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.5, fill: { color: MIDNIGHT } });
  s5.addText("What Your Site Needs to Do", {
    x: 0.6, y: 0.25, w: 8.8, h: 0.6,
    fontSize: 36, fontFace: HEADER_FONT, color: WHITE, bold: true, margin: 0
  });
  s5.addText("Most visitors arrive already knowing about you. The site confirms what they\u2019ve heard.", {
    x: 0.6, y: 0.85, w: 8.8, h: 0.35,
    fontSize: 13, fontFace: BODY_FONT, color: ACCENT_GOLD, italic: true, margin: 0
  });

  const siteNeeds = [
    { num: "01", title: "Validate Trust", desc: "A parent heard about you from a referral. They\u2019re on your site to confirm: \u201CAre these people the real deal?\u201D The answer must be immediate.", icon: iconCheckT },
    { num: "02", title: "Meet Them Emotionally", desc: "Within 5 seconds: empathy, then credibility (18 years, NARR, Task Force), then a clear next step.", icon: iconHeartT },
    { num: "03", title: "Educate, Don\u2019t Sell", desc: "Answer the questions a skeptical parent has: What is this? How is it different? Why trust you?", icon: iconBookT },
    { num: "04", title: "Be Available 24/7", desc: "An AI assistant provides a warm, knowledgeable presence whenever someone needs answers \u2014 day or night.", icon: iconClockT },
  ];

  siteNeeds.forEach((item, i) => {
    const xBase = 0.5 + i * 2.35;
    const yBase = 1.8;
    const cw = 2.15;

    // Card
    s5.addShape(pres.shapes.RECTANGLE, { x: xBase, y: yBase, w: cw, h: 3.4, fill: { color: WHITE }, shadow: makeShadow() });

    // Number badge
    s5.addShape(pres.shapes.OVAL, { x: xBase + (cw - 0.55) / 2, y: yBase + 0.2, w: 0.55, h: 0.55, fill: { color: DEEP_TEAL } });
    s5.addText(item.num, {
      x: xBase + (cw - 0.55) / 2, y: yBase + 0.2, w: 0.55, h: 0.55,
      fontSize: 16, fontFace: HEADER_FONT, color: WHITE, bold: true, align: "center", valign: "middle", margin: 0
    });

    s5.addText(item.title, {
      x: xBase + 0.1, y: yBase + 0.9, w: cw - 0.2, h: 0.35,
      fontSize: 14, fontFace: HEADER_FONT, color: MIDNIGHT, bold: true, align: "center", margin: 0
    });

    s5.addText(item.desc, {
      x: xBase + 0.15, y: yBase + 1.35, w: cw - 0.3, h: 1.8,
      fontSize: 10.5, fontFace: BODY_FONT, color: "555555", align: "left", margin: 0, valign: "top"
    });
  });

  // ============================================================
  // SLIDE 6: HOW THE HOMEPAGE WORKS
  // ============================================================
  let s6 = pres.addSlide();
  s6.background = { color: MIDNIGHT };

  s6.addText("How the Homepage Works", {
    x: 0.6, y: 0.25, w: 8.8, h: 0.6,
    fontSize: 36, fontFace: HEADER_FONT, color: WHITE, bold: true, margin: 0
  });
  s6.addText("Someone was told \u201Ccheck out Lodge360.\u201D This is what they see and feel as they scroll.", {
    x: 0.6, y: 0.82, w: 8.8, h: 0.35,
    fontSize: 13, fontFace: BODY_FONT, color: ACCENT_GOLD, italic: true, margin: 0
  });

  const hpSections = [
    { num: "01", title: "You have this problem. We understand.", desc: "Empathy and credibility in the first 5 seconds. Not a sales pitch \u2014 recognition. Jim & Adrienne\u2019s presence is felt immediately." },
    { num: "02", title: "Here\u2019s who we are and why we do this.", desc: "Since 2008. 9 homes. A waiting list. Founded by two people who\u2019ve been through it themselves. NARR certified. Task Force members." },
    { num: "03", title: "Here\u2019s what we offer and how it works.", desc: "Programs explained clearly: Safety Net levels, Chart Your Own Course. Outcomes-focused, not feature-focused." },
    { num: "04", title: "Hear from people who\u2019ve been where you are.", desc: "Real stories, not testimonials. Families and residents sharing their experience in their own words." },
    { num: "05", title: "When you\u2019re ready, here\u2019s how to reach us.", desc: "Phone, form, AI chat. Multiple paths, zero pressure. \u201CWe\u2019re here when you\u2019re ready.\u201D" },
  ];

  hpSections.forEach((sec, i) => {
    const yBase = 1.35 + i * 0.82;
    s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yBase, w: 9, h: 0.7, fill: { color: CARD_BG } });

    // Number badge
    s6.addShape(pres.shapes.OVAL, { x: 0.7, y: yBase + 0.12, w: 0.45, h: 0.45, fill: { color: DEEP_TEAL } });
    s6.addText(sec.num, {
      x: 0.7, y: yBase + 0.12, w: 0.45, h: 0.45,
      fontSize: 14, fontFace: HEADER_FONT, color: WHITE, bold: true, align: "center", valign: "middle", margin: 0
    });

    s6.addText(sec.title, {
      x: 1.35, y: yBase + 0.05, w: 4, h: 0.3,
      fontSize: 13, fontFace: HEADER_FONT, color: WHITE, bold: true, margin: 0
    });
    s6.addText(sec.desc, {
      x: 1.35, y: yBase + 0.35, w: 7.9, h: 0.3,
      fontSize: 10, fontFace: BODY_FONT, color: BODY_LIGHT, margin: 0
    });
  });

  // ============================================================
  // SLIDE 7: SITE ARCHITECTURE (solution-forward)
  // ============================================================
  let s7 = pres.addSlide();
  s7.background = { color: WARM_WHITE };

  s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.2, fill: { color: MIDNIGHT } });
  s7.addText("Site Architecture", {
    x: 0.6, y: 0.15, w: 8.8, h: 0.6,
    fontSize: 36, fontFace: HEADER_FONT, color: WHITE, bold: true, margin: 0
  });
  s7.addText("Solution first. Depth as proof. Every page earns the next click.", {
    x: 0.6, y: 0.72, w: 8.8, h: 0.3,
    fontSize: 13, fontFace: BODY_FONT, color: ACCENT_GOLD, italic: true, margin: 0
  });

  // PRIMARY NAV label
  s7.addText("PRIMARY NAVIGATION", {
    x: 0.5, y: 1.35, w: 9, h: 0.25,
    fontSize: 9, fontFace: BODY_FONT, color: DEEP_TEAL, bold: true, charSpacing: 3, margin: 0
  });

  // Primary pages: 4 cards (solution-forward order)
  const primaryPages = [
    { title: "Programs", q: "\u201CWhat do you offer?\u201D", desc: "Safety Net levels (1\u20133), Chart Your Own Course 12-week program. Outcomes-focused \u2014 what\u2019s included, who it\u2019s for, what to expect.", icon: iconCompassW },
    { title: "Our Story", q: "\u201CWho are these people?\u201D", desc: "Jim and Adrienne\u2019s personal recovery journeys. They\u2019re relatable \u2014 they\u2019ve been where you are. Since 2008. NARR certified.", icon: iconHeartW },
    { title: "For Families", q: "\u201CDo you understand me?\u201D", desc: "Speaks directly to parents. Addresses fear, frustration, financial exhaustion. Real questions, real answers.", icon: iconUsersW },
    { title: "Get Started", q: "\u201CWhat do I do next?\u201D", desc: "Phone, form, AI chat. Multiple paths, no pressure. \u201CWhen you\u2019re ready, we\u2019re here.\u201D", icon: iconPhoneW },
  ];

  primaryPages.forEach((p, i) => {
    const xBase = 0.5 + i * 2.35;
    const cw = 2.15;
    const ch = 2.1;
    const yBase = 1.7;

    s7.addShape(pres.shapes.RECTANGLE, { x: xBase, y: yBase, w: cw, h: ch, fill: { color: MIDNIGHT }, shadow: makeShadow() });
    s7.addShape(pres.shapes.OVAL, { x: xBase + 0.12, y: yBase + 0.12, w: 0.35, h: 0.35, fill: { color: DEEP_TEAL } });
    s7.addImage({ data: p.icon, x: xBase + 0.17, y: yBase + 0.17, w: 0.25, h: 0.25 });

    s7.addText(p.title, {
      x: xBase + 0.55, y: yBase + 0.1, w: cw - 0.7, h: 0.35,
      fontSize: 15, fontFace: HEADER_FONT, color: WHITE, bold: true, margin: 0
    });
    s7.addText(p.q, {
      x: xBase + 0.12, y: yBase + 0.52, w: cw - 0.24, h: 0.25,
      fontSize: 10, fontFace: BODY_FONT, color: ACCENT_GOLD, italic: true, margin: 0
    });
    s7.addText(p.desc, {
      x: xBase + 0.12, y: yBase + 0.82, w: cw - 0.24, h: 1.15,
      fontSize: 9.5, fontFace: BODY_FONT, color: BODY_LIGHT, margin: 0, valign: "top"
    });
  });

  // SUPPORTING DEPTH label
  s7.addText("SUPPORTING DEPTH  (secondary nav / footer)", {
    x: 0.5, y: 3.95, w: 9, h: 0.25,
    fontSize: 9, fontFace: BODY_FONT, color: DEEP_TEAL, bold: true, charSpacing: 3, margin: 0
  });

  // Secondary pages: 3 smaller cards
  const secondaryPages = [
    { title: "Resources & Blog", desc: "Educational articles, guides, FAQ. Proves 18 years of expertise. Feeds the AI assistant\u2019s knowledge base.", icon: iconBookW },
    { title: "Community", desc: "Resident stories, alumni perspectives, gallery. Real proof from real people \u2014 not testimonials, evidence.", icon: iconHandsW },
    { title: "FAQ", desc: "Answers to the hardest questions: cost, process, what happens if it doesn\u2019t work. Organized by stage.", icon: iconCommentsW },
  ];

  secondaryPages.forEach((p, i) => {
    const xBase = 0.5 + i * 3.15;
    const cw = 2.9;
    const ch = 1.2;
    const yBase = 4.3;

    s7.addShape(pres.shapes.RECTANGLE, { x: xBase, y: yBase, w: cw, h: ch, fill: { color: CARD_BG }, shadow: makeShadow() });
    s7.addShape(pres.shapes.OVAL, { x: xBase + 0.12, y: yBase + 0.12, w: 0.3, h: 0.3, fill: { color: DEEP_TEAL } });
    s7.addImage({ data: p.icon, x: xBase + 0.17, y: yBase + 0.17, w: 0.2, h: 0.2 });
    s7.addText(p.title, {
      x: xBase + 0.5, y: yBase + 0.1, w: cw - 0.65, h: 0.3,
      fontSize: 13, fontFace: HEADER_FONT, color: WHITE, bold: true, margin: 0
    });
    s7.addText(p.desc, {
      x: xBase + 0.12, y: yBase + 0.5, w: cw - 0.24, h: 0.6,
      fontSize: 9.5, fontFace: BODY_FONT, color: BODY_LIGHT, margin: 0, valign: "top"
    });
  });

  // ============================================================
  // SLIDE 8: CONTENT STRATEGY
  // ============================================================
  // SLIDE 8: INSPIRATION — SITES THAT DO THIS WELL
  // ============================================================
  let sInspire = pres.addSlide();
  sInspire.background = { color: MIDNIGHT };

  sInspire.addText("Sites That Do This Well", {
    x: 0.6, y: 0.25, w: 8.8, h: 0.6,
    fontSize: 30, fontFace: HEADER_FONT, color: WHITE, bold: true, margin: 0
  });
  sInspire.addText("Real examples of the solution-forward, trust-first structure we\u2019re building for you.", {
    x: 0.6, y: 0.82, w: 8.8, h: 0.3,
    fontSize: 12, fontFace: BODY_FONT, color: ACCENT_GOLD, italic: true, margin: 0
  });

  const inspSites = [
    {
      name: "Design for Recovery",
      url: "designforrecovery.com",
      why: "Sober living peer. Leads with empathy, shows real homes, founder story woven throughout. NARR-visible. Trust signals above the fold.",
    },
    {
      name: "Aware Recovery Care",
      url: "awarerecoverycare.com",
      why: "Segments by audience (individuals, families, providers). Specific outcome metrics (86% satisfaction, 89% sobriety) build credibility over promises.",
    },
    {
      name: "Recovery Centers of Am.",
      url: "recoverycentersofamerica.com",
      why: "Mission-forward headline. Trust badges (Newsweek, CARF) integrated naturally. Clear primary vs. secondary nav. Family section prominent.",
    },
    {
      name: "Hazelden Betty Ford",
      url: "hazeldenbettyford.org",
      why: "Problem-recognition positioning. Three clear visitor pathways. Minimal, high-trust design. Self-assessment tool bridges interest to action.",
    },
  ];

  inspSites.forEach((site, i) => {
    const yBase = 1.35 + i * 1.05;
    sInspire.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yBase, w: 9, h: 0.9, fill: { color: CARD_BG } });
    sInspire.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yBase, w: 0.06, h: 0.9, fill: { color: ACCENT_GOLD } });

    sInspire.addText(site.name, {
      x: 0.85, y: yBase + 0.08, w: 3, h: 0.3,
      fontSize: 14, fontFace: HEADER_FONT, color: WHITE, bold: true, margin: 0
    });
    sInspire.addText(site.url, {
      x: 4, y: yBase + 0.08, w: 3, h: 0.3,
      fontSize: 10, fontFace: BODY_FONT, color: ACCENT_GOLD, margin: 0
    });
    sInspire.addText(site.why, {
      x: 0.85, y: yBase + 0.42, w: 8.4, h: 0.4,
      fontSize: 10, fontFace: BODY_FONT, color: BODY_LIGHT, margin: 0, valign: "top"
    });
  });

  sInspire.addText("These sites validate our structural approach. Lodge360 will combine the best of each with your authentic story.", {
    x: 0.6, y: 5.1, w: 8.8, h: 0.35,
    fontSize: 11, fontFace: BODY_FONT, color: BODY_LIGHT, italic: true, margin: 0, align: "center"
  });

  // ============================================================
  // SLIDE 9: CONTENT STRATEGY
  // ============================================================
  let s8 = pres.addSlide();
  s8.background = { color: WARM_WHITE };

  s8.addText("Content Strategy", {
    x: 0.6, y: 0.4, w: 8.8, h: 0.7,
    fontSize: 36, fontFace: HEADER_FONT, color: MIDNIGHT, bold: true, margin: 0
  });
  s8.addText("Every piece of content earns trust by answering what the parent actually needs to know.", {
    x: 0.6, y: 1.1, w: 8.8, h: 0.35,
    fontSize: 13, fontFace: BODY_FONT, color: DEEP_TEAL, italic: true, margin: 0
  });

  // Voice & Tone card
  s8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.7, w: 4.25, h: 3.6, fill: { color: WHITE }, shadow: makeShadow() });
  s8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.7, w: 0.07, h: 3.6, fill: { color: DEEP_TEAL } });
  s8.addText("Voice & Tone", {
    x: 0.85, y: 1.85, w: 3.7, h: 0.4,
    fontSize: 18, fontFace: HEADER_FONT, color: MIDNIGHT, bold: true, margin: 0
  });

  const voiceTone = [
    { label: "Direct", desc: "Name the reality these families live with" },
    { label: "Experienced", desc: "\u201CWe\u2019ve seen this hundreds of times\u201D" },
    { label: "Non-promotional", desc: "Educate and inform, let the reputation speak" },
    { label: "Founder-present", desc: "Jim and Adrienne\u2019s voice woven through naturally" },
  ];

  voiceTone.forEach((v, i) => {
    const y = 2.4 + i * 0.65;
    s8.addText(v.label, {
      x: 0.85, y, w: 3.65, h: 0.25,
      fontSize: 12, fontFace: BODY_FONT, color: DEEP_TEAL, bold: true, margin: 0
    });
    s8.addText(v.desc, {
      x: 0.85, y: y + 0.25, w: 3.65, h: 0.3,
      fontSize: 10.5, fontFace: BODY_FONT, color: "555555", margin: 0
    });
  });

  // Content Pillars card
  s8.addShape(pres.shapes.RECTANGLE, { x: 5.25, y: 1.7, w: 4.25, h: 3.6, fill: { color: WHITE }, shadow: makeShadow() });
  s8.addShape(pres.shapes.RECTANGLE, { x: 5.25, y: 1.7, w: 0.07, h: 3.6, fill: { color: ACCENT_GOLD } });
  s8.addText("Content Pillars", {
    x: 5.6, y: 1.85, w: 3.7, h: 0.4,
    fontSize: 18, fontFace: HEADER_FONT, color: MIDNIGHT, bold: true, margin: 0
  });

  const pillars = [
    "\u201CWhat to look for in a recovery residence\u201D",
    "\u201CQuestions parents should ask before choosing\u201D",
    "\u201CWhat the first 30 days look like\u201D",
    "FAQ: \u201CQuestions parents ask us\u201D",
    "Media features & community advocacy",
  ];

  pillars.forEach((p, i) => {
    s8.addText([{ text: p, options: { bullet: true } }], {
      x: 5.6, y: 2.4 + i * 0.5, w: 3.65, h: 0.4,
      fontSize: 11, fontFace: BODY_FONT, color: "555555", margin: 0
    });
  });

  // ============================================================
  // SLIDE 9: AI-POWERED ASSISTANT (robust description, better design)
  // ============================================================
  let s9 = pres.addSlide();
  s9.background = { color: WARM_WHITE };

  // Full-width dark top section
  s9.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.8, fill: { color: MIDNIGHT } });
  s9.addImage({ data: iconRobotW, x: 0.6, y: 0.3, w: 0.5, h: 0.5 });
  s9.addText("AI-Powered Assistant", {
    x: 1.2, y: 0.25, w: 8.2, h: 0.6,
    fontSize: 36, fontFace: HEADER_FONT, color: WHITE, bold: true, margin: 0
  });
  s9.addText("When someone needs answers and no one\u2019s available, the AI assistant steps in \u2014 with your voice, your values, and your expertise.", {
    x: 0.6, y: 0.95, w: 8.8, h: 0.65,
    fontSize: 14, fontFace: BODY_FONT, color: ACCENT_GOLD, italic: true, margin: 0
  });

  // Left column: What It Does
  s9.addText("What It Does", {
    x: 0.6, y: 2.05, w: 4.2, h: 0.35,
    fontSize: 18, fontFace: HEADER_FONT, color: MIDNIGHT, bold: true, margin: 0
  });

  const whatItDoes = [
    "Intelligently helps visitors get answers to serious, deeply personal questions about recovery, your programs, and next steps",
    "Provides warm, accurate information trained on your philosophy, your values, and your real approach to recovery",
    "Collects contact details and routes warm leads for personal follow-up by Jim or Adrienne",
    "Available 24/7 \u2014 meets families in their most vulnerable moments with genuine understanding",
    "Guides visitors through the process without pressure \u2014 when they\u2019re ready, it connects them to you",
  ];
  whatItDoes.forEach((item, i) => {
    s9.addText([{ text: item, options: { bullet: true } }], {
      x: 0.6, y: 2.5 + i * 0.58, w: 4.2, h: 0.52,
      fontSize: 10.5, fontFace: BODY_FONT, color: "444444", margin: 0, paraSpaceAfter: 4
    });
  });

  // Right column: How It\u2019s Built
  s9.addText("How It\u2019s Built", {
    x: 5.2, y: 2.05, w: 4.3, h: 0.35,
    fontSize: 18, fontFace: HEADER_FONT, color: MIDNIGHT, bold: true, margin: 0
  });

  const howBuilt = [
    "Custom-trained on your content, programs, values, and the way Jim and Adrienne actually speak to families",
    "Powered by Claude AI \u2014 sophisticated, nuanced, and conversational (not a generic chatbot)",
    "Draws from a curated knowledge base of resources, FAQs, and program details",
    "Escalates seamlessly to a real person when the conversation requires it",
    "HIPAA-conscious design \u2014 guides and connects, never diagnoses or makes clinical claims",
  ];
  howBuilt.forEach((item, i) => {
    s9.addText([{ text: item, options: { bullet: true } }], {
      x: 5.2, y: 2.5 + i * 0.58, w: 4.3, h: 0.52,
      fontSize: 10.5, fontFace: BODY_FONT, color: "444444", margin: 0, paraSpaceAfter: 4
    });
  });

  // ============================================================
  // SLIDE 10: EMAIL OUTREACH STRATEGY
  // ============================================================
  let s10 = pres.addSlide();
  s10.background = { color: MIDNIGHT };

  s10.addText("Email Outreach Strategy", {
    x: 0.6, y: 0.35, w: 8.8, h: 0.6,
    fontSize: 36, fontFace: HEADER_FONT, color: WHITE, bold: true, margin: 0
  });
  s10.addText("Your database is one of your most underutilized assets. Let\u2019s change that.", {
    x: 0.6, y: 0.95, w: 8.8, h: 0.35,
    fontSize: 13, fontFace: BODY_FONT, color: ACCENT_GOLD, italic: true, margin: 0
  });

  // Step 1 callout
  s10.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 9, h: 0.5, fill: { color: CARD_BG } });
  s10.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 0.06, h: 0.5, fill: { color: ACCENT_GOLD } });
  s10.addText([
    { text: "Step 1: Data Audit  \u2014  ", options: { bold: true, color: WHITE } },
    { text: "We start by understanding what you have \u2014 where contacts live, how many, and what we know about each.", options: { color: BODY_LIGHT } }
  ], {
    x: 0.85, y: 1.5, w: 8.45, h: 0.5,
    fontSize: 11, fontFace: BODY_FONT, margin: 0, valign: "middle"
  });

  // 4 audience segments
  const segments = [
    { title: "Parents & Families", desc: "Educational series: what to look for, questions to ask, the first 30 days.", freq: "Bi-weekly \u2192 weekly", icon: iconUsersW },
    { title: "Alumni", desc: "Personal notes from Jim or Adrienne, community milestones, encouragement.", freq: "Monthly", icon: iconHeartW },
    { title: "Referral Partners", desc: "Industry insights, program updates, thought leadership from Jim & Adrienne.", freq: "Monthly", icon: iconHandsW },
    { title: "Past Inquiries", desc: "Gentle re-engagement. New content. \u201CWe\u2019re still here when you\u2019re ready.\u201D", freq: "Bi-weekly \u2192 monthly", icon: iconClockW },
  ];

  segments.forEach((seg, i) => {
    const xBase = 0.5 + i * 2.3;
    const yBase = 2.25;
    const cw = 2.15;

    s10.addShape(pres.shapes.RECTANGLE, { x: xBase, y: yBase, w: cw, h: 2.95, fill: { color: CARD_BG } });

    // Icon in teal circle
    s10.addShape(pres.shapes.OVAL, { x: xBase + (cw - 0.45) / 2, y: yBase + 0.15, w: 0.45, h: 0.45, fill: { color: DEEP_TEAL } });
    s10.addImage({ data: seg.icon, x: xBase + (cw - 0.3) / 2, y: yBase + 0.225, w: 0.3, h: 0.3 });

    s10.addText(seg.title, {
      x: xBase + 0.1, y: yBase + 0.7, w: cw - 0.2, h: 0.35,
      fontSize: 12, fontFace: HEADER_FONT, color: WHITE, bold: true, align: "center", margin: 0
    });
    s10.addText(seg.desc, {
      x: xBase + 0.1, y: yBase + 1.1, w: cw - 0.2, h: 1.1,
      fontSize: 10, fontFace: BODY_FONT, color: BODY_LIGHT, align: "center", margin: 0, valign: "top"
    });

    // Frequency badge
    s10.addShape(pres.shapes.RECTANGLE, { x: xBase + 0.2, y: yBase + 2.4, w: cw - 0.4, h: 0.35, fill: { color: DEEP_TEAL } });
    s10.addText(seg.freq, {
      x: xBase + 0.2, y: yBase + 2.4, w: cw - 0.4, h: 0.35,
      fontSize: 9, fontFace: BODY_FONT, color: WHITE, bold: true, align: "center", valign: "middle", margin: 0
    });
  });

  // ============================================================
  // SLIDE 11: INVESTMENT & INFRASTRUCTURE (redesigned, corrected costs)
  // ============================================================
  let s11 = pres.addSlide();
  s11.background = { color: WARM_WHITE };

  s11.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.2, fill: { color: MIDNIGHT } });
  s11.addText("Investment & Infrastructure", {
    x: 0.6, y: 0.15, w: 8.8, h: 0.6,
    fontSize: 36, fontFace: HEADER_FONT, color: WHITE, bold: true, margin: 0
  });
  s11.addText("Your complete monthly cost picture \u2014 no hidden fees, no surprises.", {
    x: 0.6, y: 0.72, w: 8.8, h: 0.3,
    fontSize: 13, fontFace: BODY_FONT, color: ACCENT_GOLD, italic: true, margin: 0
  });

  const infra = [
    { label: "AI Platform", tool: "Claude API", desc: "Powers the AI assistant + content generation. Usage-based pricing.", cost: "$30\u201350/mo", type: "Variable", icon: iconBrainT },
    { label: "Website Hosting", tool: "Next.js + Vercel Pro", desc: "Modern, fast, SEO-optimized. Includes $20 usage credit.", cost: "$20\u201330/mo", type: "Fixed+", icon: iconGlobeT },
    { label: "Email Platform", tool: "Mailchimp Essentials*", desc: "Email sequences by audience. Scales with contacts. *If needed.", cost: "$13\u201345/mo", type: "Variable*", icon: iconEnvelopeT },
    { label: "CRM", tool: "HubSpot (Free Tier)*", desc: "Contact database & tracking. *If not already using a CRM.", cost: "$0/mo", type: "Fixed*", icon: iconUsersT },
    { label: "Database", tool: "Supabase Pro", desc: "Backend for AI assistant, forms, analytics. Usage overages possible.", cost: "$25\u201340/mo", type: "Fixed+", icon: iconRobotT },
    { label: "Automation", tool: "N8N Cloud Starter", desc: "Workflows: lead routing, email triggers, notifications. 2,500 exec/mo.", cost: "$24/mo", type: "Fixed", icon: iconCogW },
    { label: "Domain", tool: "lodge360.com", desc: "Clean, memorable, professional URL.", cost: "~$20/yr", type: "Fixed", icon: iconGlobeT },
  ];

  // Table-style rows
  // Header row
  s11.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.45, w: 9, h: 0.38, fill: { color: DEEP_TEAL } });
  s11.addText("Category", { x: 0.7, y: 1.45, w: 1.3, h: 0.38, fontSize: 10, fontFace: BODY_FONT, color: WHITE, bold: true, valign: "middle", margin: 0 });
  s11.addText("Tool", { x: 2.1, y: 1.45, w: 1.8, h: 0.38, fontSize: 10, fontFace: BODY_FONT, color: WHITE, bold: true, valign: "middle", margin: 0 });
  s11.addText("Purpose", { x: 3.95, y: 1.45, w: 3.1, h: 0.38, fontSize: 10, fontFace: BODY_FONT, color: WHITE, bold: true, valign: "middle", margin: 0 });
  s11.addText("Type", { x: 7.1, y: 1.45, w: 0.9, h: 0.38, fontSize: 10, fontFace: BODY_FONT, color: WHITE, bold: true, valign: "middle", align: "center", margin: 0 });
  s11.addText("Est. Cost", { x: 8.05, y: 1.45, w: 1.4, h: 0.38, fontSize: 10, fontFace: BODY_FONT, color: WHITE, bold: true, valign: "middle", align: "right", margin: 0 });

  infra.forEach((item, i) => {
    const yBase = 1.88 + i * 0.45;
    const bgColor = i % 2 === 0 ? "F0F0F0" : WHITE;
    s11.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yBase, w: 9, h: 0.42, fill: { color: bgColor } });
    s11.addText(item.label, {
      x: 0.7, y: yBase, w: 1.3, h: 0.42,
      fontSize: 10, fontFace: HEADER_FONT, color: MIDNIGHT, bold: true, valign: "middle", margin: 0
    });
    s11.addText(item.tool, {
      x: 2.1, y: yBase, w: 1.8, h: 0.42,
      fontSize: 10, fontFace: BODY_FONT, color: DEEP_TEAL, valign: "middle", margin: 0
    });
    s11.addText(item.desc, {
      x: 3.95, y: yBase, w: 3.1, h: 0.42,
      fontSize: 8.5, fontFace: BODY_FONT, color: "666666", valign: "middle", margin: 0
    });
    const typeColor = item.type.includes("Variable") ? "B8860B" : DEEP_TEAL;
    s11.addText(item.type.replace("*", ""), {
      x: 7.1, y: yBase, w: 0.9, h: 0.42,
      fontSize: 8, fontFace: BODY_FONT, color: typeColor, bold: true, valign: "middle", align: "center", margin: 0
    });
    s11.addText(item.cost, {
      x: 8.05, y: yBase, w: 1.4, h: 0.42,
      fontSize: 10, fontFace: BODY_FONT, color: MIDNIGHT, bold: true, valign: "middle", align: "right", margin: 0
    });
  });

  // Total bar
  s11.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.05, w: 9, h: 0.4, fill: { color: MIDNIGHT } });
  s11.addText("Estimated range: ~$112\u2013189/mo  +  ~$20/yr domain", {
    x: 0.7, y: 5.05, w: 8.6, h: 0.4,
    fontSize: 12, fontFace: BODY_FONT, color: WHITE, bold: true, valign: "middle", align: "center", margin: 0
  });

  s11.addText("Fixed+ = base price with possible usage overages  |  Variable = scales with usage/contacts  |  *Dependent on what you already have in place", {
    x: 0.5, y: 5.5, w: 9, h: 0.25,
    fontSize: 7.5, fontFace: BODY_FONT, color: "999999", italic: true, margin: 0
  });

  // ============================================================
  // SLIDE 12: TIMELINE
  // ============================================================
  let s12 = pres.addSlide();
  s12.background = { color: MIDNIGHT };

  s12.addText("What\u2019s Needed to Get There", {
    x: 0.6, y: 0.35, w: 8.8, h: 0.6,
    fontSize: 36, fontFace: HEADER_FONT, color: WHITE, bold: true, margin: 0
  });
  s12.addText("Here\u2019s what each phase involves, how long it takes, and how hands-on you need to be.", {
    x: 0.6, y: 0.95, w: 8.8, h: 0.35,
    fontSize: 13, fontFace: BODY_FONT, color: ACCENT_GOLD, italic: true, margin: 0
  });

  const phases = [
    { title: "Discovery", time: "~1\u20132 weeks", desc: "Answer questionnaire. Share photos, materials, and content inputs. Audit contact database together. Align on tone and chatbot personality.", involvement: "Medium", icon: iconSearchW },
    { title: "Build", time: "~2\u20133 weeks", desc: "Site design, development, and content writing. Chatbot training. Email template and segment setup. Zack handles the heavy lifting.", involvement: "Low", icon: iconRocketW },
    { title: "Review", time: "~1 week", desc: "You review everything. Feedback rounds. Content and chatbot polish. Test email sequences.", involvement: "Medium", icon: iconCheckW },
    { title: "Launch", time: "~1 week", desc: "Site goes live. Email campaigns activated. Analytics configured. Ongoing support continues.", involvement: "Low", icon: iconStarW },
  ];

  phases.forEach((p, i) => {
    const xBase = 0.5 + i * 2.35;
    const yBase = 1.55;
    const cw = 2.15;
    const ch = 3.55;

    s12.addShape(pres.shapes.RECTANGLE, { x: xBase, y: yBase, w: cw, h: ch, fill: { color: CARD_BG } });

    // Icon in circle
    s12.addShape(pres.shapes.OVAL, { x: xBase + (cw - 0.5) / 2, y: yBase + 0.2, w: 0.5, h: 0.5, fill: { color: DEEP_TEAL } });
    s12.addImage({ data: p.icon, x: xBase + (cw - 0.3) / 2, y: yBase + 0.3, w: 0.3, h: 0.3 });

    s12.addText(p.title, {
      x: xBase + 0.1, y: yBase + 0.8, w: cw - 0.2, h: 0.3,
      fontSize: 16, fontFace: HEADER_FONT, color: WHITE, bold: true, align: "center", margin: 0
    });
    s12.addText(p.time, {
      x: xBase + 0.1, y: yBase + 1.12, w: cw - 0.2, h: 0.25,
      fontSize: 11, fontFace: BODY_FONT, color: ACCENT_GOLD, bold: true, align: "center", margin: 0
    });
    s12.addText(p.desc, {
      x: xBase + 0.15, y: yBase + 1.5, w: cw - 0.3, h: 1.3,
      fontSize: 10, fontFace: BODY_FONT, color: BODY_LIGHT, align: "left", margin: 0, valign: "top"
    });

    // Involvement badge
    const invColor = p.involvement === "Low" ? DEEP_TEAL : ACCENT_GOLD;
    s12.addShape(pres.shapes.RECTANGLE, { x: xBase + 0.25, y: yBase + 3.0, w: cw - 0.5, h: 0.35, fill: { color: invColor } });
    s12.addText("Your involvement: " + p.involvement, {
      x: xBase + 0.25, y: yBase + 3.0, w: cw - 0.5, h: 0.35,
      fontSize: 9, fontFace: BODY_FONT, color: WHITE, bold: true, align: "center", valign: "middle", margin: 0
    });
  });

  // ============================================================
  // SLIDE 13: SITE STRUCTURE — SOLUTION FORWARD
  // ============================================================
  let s13 = pres.addSlide();
  s13.background = { color: WARM_WHITE };

  s13.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.2, fill: { color: MIDNIGHT } });
  s13.addText("Your Site at a Glance", {
    x: 0.6, y: 0.15, w: 8.8, h: 0.6,
    fontSize: 36, fontFace: HEADER_FONT, color: WHITE, bold: true, margin: 0
  });
  s13.addText("Solution first. Depth as proof. Resources earn respect \u2014 they don\u2019t lead.", {
    x: 0.6, y: 0.72, w: 8.8, h: 0.3,
    fontSize: 13, fontFace: BODY_FONT, color: ACCENT_GOLD, italic: true, margin: 0
  });

  // HOME box (top center)
  const homeX = 3.75, homeY = 1.45, boxW = 2.5, boxH = 0.55;
  s13.addShape(pres.shapes.RECTANGLE, { x: homeX, y: homeY, w: boxW, h: boxH, fill: { color: DEEP_TEAL }, shadow: makeShadow() });
  s13.addText("Home", {
    x: homeX, y: homeY, w: boxW, h: boxH,
    fontSize: 16, fontFace: HEADER_FONT, color: WHITE, bold: true, align: "center", valign: "middle", margin: 0
  });

  // PRIMARY NAV label
  s13.addText("PRIMARY NAV", {
    x: 0.5, y: 2.15, w: 2, h: 0.2,
    fontSize: 8, fontFace: BODY_FONT, color: DEEP_TEAL, bold: true, charSpacing: 2, margin: 0
  });

  // Primary pages (solution-forward order)
  const primPages = ["Programs", "Our Story", "For Families", "Get Started"];
  const primY = 2.4;
  const primW = 2.1;
  const primH = 0.45;
  const primGap = 0.2;
  const primTotalW = primPages.length * primW + (primPages.length - 1) * primGap;
  const primStartX = (10 - primTotalW) / 2;

  // Vertical line from Home down
  const primConnMidY = homeY + boxH + (primY - homeY - boxH) / 2;
  s13.addShape(pres.shapes.LINE, { x: homeX + boxW / 2, y: homeY + boxH, w: 0, h: primConnMidY - (homeY + boxH), line: { color: DEEP_TEAL, width: 1.5 } });
  // Horizontal connector
  const primFirstCenter = primStartX + primW / 2;
  const primLastCenter = primStartX + 3 * (primW + primGap) + primW / 2;
  s13.addShape(pres.shapes.LINE, { x: primFirstCenter, y: primConnMidY, w: primLastCenter - primFirstCenter, h: 0, line: { color: DEEP_TEAL, width: 1.5 } });

  primPages.forEach((page, i) => {
    const xBase = primStartX + i * (primW + primGap);
    s13.addShape(pres.shapes.RECTANGLE, { x: xBase, y: primY, w: primW, h: primH, fill: { color: MIDNIGHT }, shadow: makeShadow() });
    s13.addText(page, { x: xBase, y: primY, w: primW, h: primH, fontSize: 12, fontFace: BODY_FONT, color: WHITE, bold: true, align: "center", valign: "middle", margin: 0 });
    s13.addShape(pres.shapes.LINE, { x: xBase + primW / 2, y: primConnMidY, w: 0, h: primY - primConnMidY, line: { color: DEEP_TEAL, width: 1.5 } });
  });

  // SUPPORTING DEPTH label
  s13.addText("SUPPORTING DEPTH  (footer / secondary)", {
    x: 0.5, y: 3.05, w: 5, h: 0.2,
    fontSize: 8, fontFace: BODY_FONT, color: ACCENT_GOLD, bold: true, charSpacing: 2, margin: 0
  });

  const secPages = ["Resources & Blog", "FAQ", "Community"];
  const secY = 3.35;
  const secW = 2.6;
  const secH = 0.42;
  const secGap = 0.2;
  const secTotalW = secPages.length * secW + (secPages.length - 1) * secGap;
  const secStartX = (10 - secTotalW) / 2;

  secPages.forEach((page, i) => {
    const xBase = secStartX + i * (secW + secGap);
    s13.addShape(pres.shapes.RECTANGLE, { x: xBase, y: secY, w: secW, h: secH, fill: { color: CARD_BG }, shadow: makeShadow() });
    s13.addText(page, { x: xBase, y: secY, w: secW, h: secH, fontSize: 11, fontFace: BODY_FONT, color: BODY_LIGHT, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  // SITE-WIDE ELEMENTS
  s13.addText("SITE-WIDE ELEMENTS", {
    x: 0.5, y: 4.05, w: 3, h: 0.2,
    fontSize: 8, fontFace: BODY_FONT, color: DEEP_TEAL, bold: true, charSpacing: 2, margin: 0
  });

  const siteWide = [
    { label: "AI Assistant", desc: "Available throughout. Answers questions with your voice and values. Routes warm leads.", icon: iconRobotT },
    { label: "Trust Signals", desc: "NARR certified. 18+ years (since 2008). Delray Drug Task Force. Woven into every page naturally.", icon: iconShieldT },
    { label: "Contact Paths", desc: "Phone, form, AI chat on every page. Multiple ways to connect, zero pressure.", icon: iconPhoneT },
  ];

  siteWide.forEach((feat, i) => {
    const xBase = 0.5 + i * 3.15;
    const cw = 2.9;
    const fY = 4.35;
    s13.addShape(pres.shapes.RECTANGLE, { x: xBase, y: fY, w: cw, h: 1.1, fill: { color: WHITE }, shadow: makeShadow() });
    s13.addImage({ data: feat.icon, x: xBase + 0.12, y: fY + 0.12, w: 0.3, h: 0.3 });
    s13.addText(feat.label, {
      x: xBase + 0.5, y: fY + 0.1, w: cw - 0.65, h: 0.3,
      fontSize: 12, fontFace: HEADER_FONT, color: MIDNIGHT, bold: true, margin: 0
    });
    s13.addText(feat.desc, {
      x: xBase + 0.12, y: fY + 0.5, w: cw - 0.24, h: 0.55,
      fontSize: 9.5, fontFace: BODY_FONT, color: "555555", margin: 0, valign: "top"
    });
  });

  // ============================================================
  // SLIDE 14: NEXT STEPS (with logo)
  // ============================================================
  let s14 = pres.addSlide();
  s14.background = { color: WARM_WHITE };

  s14.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: MIDNIGHT } });
  s14.addText("Next Steps", {
    x: 0.6, y: 0.15, w: 8.8, h: 0.65,
    fontSize: 36, fontFace: HEADER_FONT, color: WHITE, bold: true, margin: 0
  });

  const steps = [
    { num: "1", title: "Review this proposal", desc: "Take your time. Note what resonates, what you\u2019d change, and any questions that come up." },
    { num: "2", title: "Answer our questions", desc: "We\u2019ve attached a short questionnaire covering programs, preferences, and the details we need to build with confidence." },
    { num: "3", title: "We\u2019ll connect when the time is right", desc: "We\u2019ll all get together, walk through your answers, agree on the plan, and start building." },
  ];

  steps.forEach((step, i) => {
    const yBase = 1.3 + i * 1.0;
    // Number circle
    s14.addShape(pres.shapes.OVAL, { x: 0.7, y: yBase + 0.08, w: 0.5, h: 0.5, fill: { color: DEEP_TEAL } });
    s14.addText(step.num, {
      x: 0.7, y: yBase + 0.08, w: 0.5, h: 0.5,
      fontSize: 20, fontFace: HEADER_FONT, color: WHITE, bold: true, align: "center", valign: "middle", margin: 0
    });
    s14.addText(step.title, {
      x: 1.4, y: yBase + 0.05, w: 7.5, h: 0.35,
      fontSize: 18, fontFace: HEADER_FONT, color: MIDNIGHT, bold: true, margin: 0
    });
    s14.addText(step.desc, {
      x: 1.4, y: yBase + 0.42, w: 7.5, h: 0.45,
      fontSize: 12, fontFace: BODY_FONT, color: "555555", margin: 0
    });
  });

  // Closing quote in dark box
  s14.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.3, w: 5.5, h: 0.8, fill: { color: MIDNIGHT } });
  s14.addText("This industry has no shortage of programs with good marketing and average results. You\u2019re the opposite \u2014 exceptional results, no marketing. Let\u2019s fix that second part.", {
    x: 0.7, y: 4.35, w: 5.1, h: 0.7,
    fontSize: 10.5, fontFace: BODY_FONT, color: WHITE, italic: true, margin: 0, valign: "middle"
  });

  // Logo on right
  s14.addImage({ data: logoClean, x: 7.2, y: 3.85, w: 1.8, h: 1.9, sizing: { type: "contain", w: 1.8, h: 1.9 } });

  // ============================================================
  // WRITE FILE
  // ============================================================
  const outPath = "/sessions/admiring-ecstatic-hypatia/mnt/the_lodge/Lodge360_Proposal.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("Presentation saved to: " + outPath);
}

buildPresentation().catch(err => { console.error(err); process.exit(1); });
