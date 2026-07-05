// ===== ADCDA Generic Presentation Builder =====
// يفصل المحتوى (content.json) عن التصميم (الهوية ثابتة في الكود).
// الاستخدام: node build.js <content.json>
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");
const FA = require("@fortawesome/free-solid-svg-icons");

function faDef(name) {
  if (!name) return FA.faInfoCircle;
  if (name === "MdStairs") return FA.faStairs || FA.faInfoCircle;
  const key = name.startsWith("Fa") ? "fa" + name.slice(2) : name;
  const d = FA[key];
  return (d && d.icon) ? d : FA.faInfoCircle;
}
function iconSvg(name, color = "0B4EA2", size = 256) {
  const I = faDef(name).icon;
  const w = I[0], h = I[1];
  let p = I[4]; if (Array.isArray(p)) p = p.join(" ");
  const svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + w + ' ' + h +
    '" width="' + size + '" height="' + size + '"><path fill="#' + color + '" d="' + p + '"/></svg>';
  return "image/svg+xml;base64," + Buffer.from(svg).toString("base64");
}

const COLORS = {
  bg: "F7F8FA", primary: "0B4EA2", primaryDark: "1F2E5A", primaryLight: "DCE7F5",
  textDark: "1F2E5A", textLight: "555555", white: "FFFFFF",
  red: "C0392B", redDark: "7A271A", redLight: "FCEEEC", green: "1E7E34", greenDark: "14532D", greenLight: "EAF4EE",
  cardBorder: "E5EBF2", warningBg: "FEF2F2", warningBorder: "FCA5A5", warningText: "991B1B"
};
const FONT = "Tajawal";
const SLIDE_W = 13.33;

async function build(content) {
  const ASSETS = path.resolve(__dirname, "assets");
  const CODE = content.code || "F.X";
  const pres = new pptxgen();
  pres.defineLayout({ name: "WIDE", width: 13.33, height: 7.5 });
  pres.layout = "WIDE";

  function addHeader(slide) {
    slide.addImage({ path: path.join(ASSETS, "logo_proud.png"), x: 0.40, y: 0.25, w: 0.85, h: 0.85 });
    slide.addImage({ path: path.join(ASSETS, "logo_adcda.png"), x: 11.10, y: 0.35, w: 1.85, h: 0.55 });
  }
  function addBgDecorations(slide) {
    slide.addShape(pres.shapes.OVAL, { x: -1.2, y: 5.0, w: 3.0, h: 3.0, fill: { color: COLORS.primaryLight, transparency: 70 }, line: { type: "none" } });
    slide.addShape(pres.shapes.OVAL, { x: 12.5, y: 5.5, w: 2.0, h: 2.0, fill: { color: COLORS.primaryLight, transparency: 70 }, line: { type: "none" } });
    slide.addShape(pres.shapes.OVAL, { x: 0.0, y: 1.8, w: 0.6, h: 0.6, fill: { color: COLORS.primaryLight, transparency: 80 }, line: { type: "none" } });
  }
  function addFooter(slide) {
    slide.addShape(pres.shapes.LINE, { x: 0.54, y: 7.07, w: 12.30, h: 0, line: { color: COLORS.cardBorder, width: 0.75 } });
    slide.addImage({ path: path.join(ASSETS, "social_icons.png"), x: 0.20, y: 7.14, w: 2.19, h: 0.32 });
    slide.addShape(pres.shapes.LINE, { x: 2.42, y: 7.15, w: 0, h: 0.32, line: { color: COLORS.cardBorder, width: 0.75 } });
    slide.addText("@civildefencead", { x: 2.50, y: 7.12, w: 1.65, h: 0.38, fontSize: 10, fontFace: FONT, color: COLORS.primaryDark, bold: true, align: "center", valign: "middle", margin: 0 });
    slide.addShape(pres.shapes.LINE, { x: 4.20, y: 7.15, w: 0, h: 0.32, line: { color: COLORS.cardBorder, width: 0.75 } });
    slide.addText("www.adcda.gov.ae", { x: 4.30, y: 7.12, w: 1.85, h: 0.38, fontSize: 10, fontFace: FONT, color: COLORS.primaryDark, bold: true, align: "center", valign: "middle", margin: 0 });
    slide.addShape(pres.shapes.LINE, { x: 6.22, y: 7.15, w: 0, h: 0.32, line: { color: COLORS.cardBorder, width: 0.75 } });
    slide.addText("عاصمة أكثر أمنا وسلامة", { x: 6.28, y: 7.17, w: 1.85, h: 0.29, fontSize: 10, fontFace: FONT, color: COLORS.textLight, align: "center", valign: "middle", margin: 0, rtlMode: true });
    slide.addShape(pres.shapes.LINE, { x: 8.20, y: 7.15, w: 0, h: 0.32, line: { color: COLORS.cardBorder, width: 0.75 } });
    slide.addText(CODE, { x: 8.30, y: 7.17, w: 0.55, h: 0.29, fontSize: 10, fontFace: FONT, color: COLORS.textLight, align: "center", valign: "middle", margin: 0 });
    slide.addShape(pres.shapes.LINE, { x: 8.95, y: 7.15, w: 0, h: 0.32, line: { color: COLORS.cardBorder, width: 0.75 } });
    slide.addText([
      { text: "الإصدار", options: { lang: "ar-AE" } }, { text: " ", options: {} },
      { text: "الأول", options: { lang: "ar-AE" } }, { text: " © 2026 ", options: {} },
      { text: "هيئة", options: { lang: "ar-AE" } }, { text: " ", options: {} },
      { text: "أبوظبي", options: { lang: "ar-AE" } }, { text: " ", options: {} },
      { text: "للدفاع", options: { lang: "ar-AE" } }, { text: " ", options: {} },
      { text: "المدني", options: { lang: "ar-AE" } }, { text: " - ", options: {} },
      { text: "جميع", options: { lang: "ar-AE" } }, { text: " ", options: {} },
      { text: "الحقوق", options: { lang: "ar-AE" } }, { text: " ", options: {} },
      { text: "محفوظة", options: { lang: "ar-AE" } }
    ], { x: 9.05, y: 7.15, w: 3.80, h: 0.32, fontSize: 8, fontFace: FONT, color: COLORS.textLight, align: "right", valign: "middle", margin: 0, rtlMode: true });
  }
  function addSlideTitle(slide, title, subtitle) {
    slide.addText(title, { x: 0.50, y: 1.40, w: 12.30, h: 0.65, fontSize: 25, fontFace: FONT, color: COLORS.primaryDark, bold: true, align: "right", valign: "middle", margin: 0, rtlMode: true });
    if (subtitle) {
      slide.addText(subtitle, { x: 0.50, y: 2.05, w: 12.30, h: 0.40, fontSize: 14, fontFace: FONT, color: COLORS.textLight, align: "right", valign: "middle", margin: 0, rtlMode: true });
      slide.addShape(pres.shapes.LINE, { x: 10.80, y: 2.55, w: 2.00, h: 0, line: { color: COLORS.primary, width: 1.75 } });
    }
  }
  function addReminderBox(slide, label, text) {
    if (!label && !text) return;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.50, y: 6.20, w: 12.30, h: 0.55, fill: { color: COLORS.primaryDark }, line: { type: "none" }, rectRadius: 0.06, shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "1F2E5A", opacity: 0.12 } });
    slide.addText([
      { text: (label || "تذكَّر") + ": ", options: { bold: true, color: "FFFFFF", lang: "ar-AE" } },
      { text: text || "", options: { color: "DCE7F5", lang: "ar-AE" } }
    ], { x: 0.70, y: 6.20, w: 11.90, h: 0.55, fontSize: 12.5, fontFace: FONT, align: "right", valign: "middle", margin: 0, rtlMode: true, lang: "ar-AE" });
  }
  async function addIconCard(slide, x, y, w, h, icon, title, desc) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: COLORS.white }, line: { color: COLORS.cardBorder, width: 0.75 }, rectRadius: 0.08, shadow: { type: "outer", blur: 8, offset: 2, angle: 90, color: "1F2E5A", opacity: 0.06 } });
    const cs = 1.0, cx = x + (w - cs) / 2, cy = y + 0.30;
    slide.addShape(pres.shapes.OVAL, { x: cx, y: cy, w: cs, h: cs, fill: { color: COLORS.primaryLight }, line: { type: "none" } });
    slide.addImage({ data: iconSvg(icon, COLORS.primary, 256), x: cx + (cs - 0.5) / 2, y: cy + (cs - 0.5) / 2, w: 0.5, h: 0.5 });
    slide.addShape(pres.shapes.LINE, { x: x + (w - 1.45) / 2, y: cy + cs + 0.20, w: 1.45, h: 0, line: { color: COLORS.cardBorder, width: 1 } });
    slide.addText(title, { x: x + 0.10, y: cy + cs + 0.30, w: w - 0.20, h: 0.45, fontSize: 14, fontFace: FONT, color: COLORS.primaryDark, bold: true, align: "center", valign: "middle", margin: 0, rtlMode: true });
    if (desc) slide.addText(desc, { x: x + 0.10, y: cy + cs + 0.75, w: w - 0.20, h: 0.55, fontSize: 11, fontFace: FONT, color: COLORS.textLight, align: "center", valign: "top", margin: 0, rtlMode: true });
  }
  async function addCompactIconCard(slide, x, y, w, h, icon, title) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: COLORS.white }, line: { color: COLORS.cardBorder, width: 0.75 }, rectRadius: 0.08, shadow: { type: "outer", blur: 8, offset: 2, angle: 90, color: "1F2E5A", opacity: 0.06 } });
    const cSize = 0.75, cX = x + 0.20, cY = y + (h - cSize) / 2;
    slide.addShape(pres.shapes.OVAL, { x: cX, y: cY, w: cSize, h: cSize, fill: { color: COLORS.primary }, line: { type: "none" } });
    slide.addImage({ data: iconSvg(icon, "FFFFFF", 256), x: cX + 0.20, y: cY + 0.20, w: 0.35, h: 0.35 });
    const textX = x + cSize + 0.35, textW = w - cSize - 0.55;
    slide.addText(title, { x: textX, y, w: textW, h, fontSize: 13, fontFace: FONT, color: COLORS.primaryDark, bold: true, align: "right", valign: "middle", margin: 0, rtlMode: true });
  }
  async function addCompactIconCardWithDesc(slide, x, y, w, h, icon, title, desc) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: COLORS.white }, line: { color: COLORS.cardBorder, width: 0.75 }, rectRadius: 0.08, shadow: { type: "outer", blur: 8, offset: 2, angle: 90, color: "1F2E5A", opacity: 0.06 } });
    const cSize = 0.85, cX = x + 0.20, cY = y + (h - cSize) / 2;
    slide.addShape(pres.shapes.OVAL, { x: cX, y: cY, w: cSize, h: cSize, fill: { color: COLORS.primary }, line: { type: "none" } });
    slide.addImage({ data: iconSvg(icon, "FFFFFF", 256), x: cX + 0.20, y: cY + 0.20, w: 0.45, h: 0.45 });
    const textX = x + cSize + 0.35, textW = w - cSize - 0.55;
    slide.addText(title, { x: textX, y: y + 0.15, w: textW, h: 0.45, fontSize: 14, fontFace: FONT, color: COLORS.primaryDark, bold: true, align: "right", valign: "middle", margin: 0, rtlMode: true });
    slide.addText(desc, { x: textX, y: y + 0.62, w: textW, h: 0.65, fontSize: 11, fontFace: FONT, color: COLORS.textLight, align: "right", valign: "top", margin: 0, rtlMode: true });
  }
  function addNumberedItem(slide, x, y, w, num, text, color = COLORS.primary) {
    const h = 0.50;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: COLORS.white }, line: { color: COLORS.cardBorder, width: 0.75 }, rectRadius: 0.05 });
    const ds = 0.32;
    slide.addShape(pres.shapes.OVAL, { x: x + w - ds - 0.10, y: y + (h - ds) / 2, w: ds, h: ds, fill: { color }, line: { type: "none" } });
    slide.addText(String(num), { x: x + w - ds - 0.10, y: y + (h - ds) / 2, w: ds, h: ds, fontSize: 11, fontFace: FONT, color: COLORS.white, bold: true, align: "center", valign: "middle", margin: 0 });
    slide.addText(text, { x: x + 0.15, y, w: w - ds - 0.40, h, fontSize: 13, fontFace: FONT, color: COLORS.primaryDark, bold: true, align: "right", valign: "middle", margin: 0, rtlMode: true });
  }
  function addNumberedItemH(slide, x, y, w, h, num, text, color = COLORS.primary) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: COLORS.white }, line: { color: COLORS.cardBorder, width: 0.75 }, rectRadius: 0.05 });
    const ds = Math.min(0.32, h - 0.06);
    slide.addShape(pres.shapes.OVAL, { x: x + w - ds - 0.10, y: y + (h - ds) / 2, w: ds, h: ds, fill: { color }, line: { type: "none" } });
    slide.addText(String(num), { x: x + w - ds - 0.10, y: y + (h - ds) / 2, w: ds, h: ds, fontSize: 10, fontFace: FONT, color: COLORS.white, bold: true, align: "center", valign: "middle", margin: 0 });
    slide.addText(text, { x: x + 0.15, y, w: w - ds - 0.40, h, fontSize: 12, fontFace: FONT, color: COLORS.primaryDark, bold: true, align: "right", valign: "middle", margin: 0, rtlMode: true });
  }
  function newSlide() { const s = pres.addSlide(); s.background = { color: COLORS.bg }; addHeader(s); addBgDecorations(s); return s; }
  function rem(s, r) { if (r) addReminderBox(s, r.label, r.text); }

  // ===== Slide type renderers =====
  async function renderCover(c) {
    const slide = pres.addSlide(); slide.background = { color: COLORS.bg }; addHeader(slide); addBgDecorations(slide);
    const cx = 6.665;
    slide.addShape(pres.shapes.OVAL, { x: cx - 1.10, y: 1.20, w: 2.20, h: 2.20, fill: { color: COLORS.primaryLight, transparency: 50 }, line: { type: "none" } });
    slide.addShape(pres.shapes.OVAL, { x: cx - 0.85, y: 1.45, w: 1.70, h: 1.70, fill: { color: COLORS.primary }, line: { type: "none" } });
    slide.addImage({ data: iconSvg(c.icon || "FaBriefcase", "FFFFFF", 512), x: cx - 0.55, y: 1.75, w: 1.10, h: 1.10 });
    slide.addText(c.title, { x: 0.50, y: 4.00, w: 12.30, h: 0.85, fontSize: 38, fontFace: FONT, color: COLORS.primaryDark, bold: true, align: "center", valign: "middle", margin: 0, rtlMode: true });
    if (c.subtitle) slide.addText(c.subtitle, { x: 0.50, y: 4.95, w: 12.30, h: 0.45, fontSize: 18, fontFace: FONT, color: COLORS.textLight, align: "center", valign: "middle", margin: 0, rtlMode: true });
    slide.addShape(pres.shapes.LINE, { x: 6.165, y: 5.45, w: 1.0, h: 0, line: { color: COLORS.primary, width: 2 } });
    if (c.quote) slide.addText('"' + c.quote + '"', { x: 1.50, y: 5.75, w: 10.30, h: 0.50, fontSize: 13, fontFace: FONT, color: COLORS.primaryDark, italic: true, align: "center", valign: "middle", margin: 0, rtlMode: true });
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.665, y: 6.30, w: 2.0, h: 0.48, fill: { color: COLORS.primary }, line: { type: "none" }, rectRadius: 0.08, shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "1F2E5A", opacity: 0.15 } });
    slide.addText([{ text: "للطوارئ  ", options: { lang: "ar-AE" } }, { text: "999", options: {} }], { x: 5.665, y: 6.30, w: 2.0, h: 0.48, fontSize: 15, fontFace: FONT, color: COLORS.white, bold: true, align: "center", valign: "middle", margin: 0, rtlMode: true });
    addFooter(slide);
  }
  async function renderCards(c) { // ≤4: صف واحد كبير | >4: شبكة 3×2 أوسع وأوضح
    const slide = newSlide(); addSlideTitle(slide, c.title, c.subtitle);
    const cards = c.cards;
    if (cards.length > 4) {
      const cw = 4.05, ch = 1.58, gx = 0.10, gy = 0.16;
      const sx = (SLIDE_W - (3 * cw + 2 * gx)) / 2;
      for (let i = 0; i < cards.length; i++) {
        const r = Math.floor(i / 3), col = i % 3;
        await addCompactIconCardWithDesc(slide, sx + col * (cw + gx), 2.78 + r * (ch + gy), cw, ch, cards[i].icon, cards[i].title, cards[i].desc);
      }
    } else {
      const cw = 2.85, ch = 2.45, gx = 0.15;
      const sx = (SLIDE_W - (cards.length * cw + (cards.length - 1) * gx)) / 2;
      for (let i = 0; i < cards.length; i++) await addIconCard(slide, sx + i * (cw + gx), 2.85, cw, ch, cards[i].icon, cards[i].title, cards[i].desc);
    }
    rem(slide, c.reminder); addFooter(slide);
  }
  async function renderCompact(c) { // items icon+title, layout 4-3 / 4-4 / 3x2
    const slide = newSlide(); addSlideTitle(slide, c.title, c.subtitle);
    const it = c.items; const layout = c.layout || "4-3";
    if (layout === "3x2") {
      const cw = 4.10, ch = 1.50, gx = 0.10, gy = 0.20;
      const sx = (SLIDE_W - (3 * cw + 2 * gx)) / 2;
      for (let i = 0; i < it.length; i++) { const r = Math.floor(i / 3), col = i % 3; await addCompactIconCard(slide, sx + col * (cw + gx), 2.85 + r * (ch + gy), cw, ch, it[i].icon, it[i].title); }
    } else if (layout === "4-4") {
      const cw = 3.05, ch = 1.45, gx = 0.10; const sx = (SLIDE_W - (4 * cw + 3 * gx)) / 2;
      for (let i = 0; i < 4 && i < it.length; i++) await addCompactIconCard(slide, sx + i * (cw + gx), 2.85, cw, ch, it[i].icon, it[i].title);
      for (let i = 4; i < it.length; i++) await addCompactIconCard(slide, sx + (i - 4) * (cw + gx), 4.45, cw, ch, it[i].icon, it[i].title);
    } else { // 4-3
      const cw = 2.85, ch = 1.55, gx = 0.10; const sx = (SLIDE_W - (4 * cw + 3 * gx)) / 2;
      for (let i = 0; i < 4 && i < it.length; i++) await addCompactIconCard(slide, sx + i * (cw + gx), 2.85, cw, ch, it[i].icon, it[i].title);
      const sx2 = (SLIDE_W - (3 * cw + 2 * gx)) / 2;
      for (let i = 4; i < it.length; i++) await addCompactIconCard(slide, sx2 + (i - 4) * (cw + gx), 4.50, cw, ch, it[i].icon, it[i].title);
    }
    rem(slide, c.reminder); addFooter(slide);
  }
  async function renderCardsDesc(c) { // 3x2 with desc
    const slide = newSlide(); addSlideTitle(slide, c.title, c.subtitle);
    const cards = c.cards; const cw = 4.10, ch = 1.50, gx = 0.10, gy = 0.15; const sx = (SLIDE_W - (3 * cw + 2 * gx)) / 2;
    for (let i = 0; i < cards.length; i++) { const r = Math.floor(i / 3), col = i % 3; await addCompactIconCardWithDesc(slide, sx + col * (cw + gx), 2.85 + r * (ch + gy), cw, ch, cards[i].icon, cards[i].title, cards[i].desc); }
    rem(slide, c.reminder); addFooter(slide);
  }
  async function renderNumbered(c) { // 8 steps, two columns
    const slide = newSlide(); addSlideTitle(slide, c.title, c.subtitle);
    const steps = c.steps; const iw = 6.05, ih = 0.50, gy = 0.13, sy = 2.85; const half = Math.ceil(steps.length / 2);
    for (let i = 0; i < half; i++) addNumberedItem(slide, 6.78, sy + i * (ih + gy), iw, i + 1, steps[i]);
    for (let i = half; i < steps.length; i++) addNumberedItem(slide, 0.50, sy + (i - half) * (ih + gy), iw, i + 1, steps[i]);
    rem(slide, c.reminder); addFooter(slide);
  }
  async function renderStairs(c) { // green icon + horizontal numbered points
    const slide = newSlide(); addSlideTitle(slide, c.title, c.subtitle);
    const lcX = 1.50, lcY = 3.20, lcW = 2.50;
    slide.addShape(pres.shapes.OVAL, { x: lcX, y: lcY, w: lcW, h: lcW, fill: { color: COLORS.greenLight }, line: { type: "none" } });
    slide.addImage({ data: iconSvg(c.icon || "MdStairs", "1B873F", 512), x: lcX + 0.50, y: lcY + 0.50, w: 1.50, h: 1.50 });
    slide.addShape(pres.shapes.OVAL, { x: lcX + lcW - 0.55, y: lcY + lcW - 0.65, w: 0.65, h: 0.65, fill: { color: COLORS.green }, line: { color: COLORS.white, width: 3 } });
    slide.addImage({ data: iconSvg("FaCheck", "FFFFFF", 256), x: lcX + lcW - 0.40, y: lcY + lcW - 0.50, w: 0.35, h: 0.35 });
    const points = c.points; const pw = 7.50, ph = 0.40, pg = 0.08, psy = 2.95;
    for (let i = 0; i < points.length; i++) addNumberedItemH(slide, 5.30, psy + i * (ph + pg), pw, ph, i + 1, points[i]);
    rem(slide, c.reminder); addFooter(slide);
  }
  async function renderEmergency(c) {
    const slide = newSlide(); addSlideTitle(slide, c.title, c.subtitle);
    const eX = 1.10, eY = 2.85, eW = 2.00;
    slide.addShape(pres.shapes.OVAL, { x: eX, y: eY, w: eW, h: eW, fill: { color: COLORS.red }, line: { type: "none" } });
    slide.addShape(pres.shapes.OVAL, { x: eX + 0.18, y: eY + 0.18, w: eW - 0.36, h: eW - 0.36, fill: { color: COLORS.white }, line: { type: "none" } });
    slide.addText("999", { x: eX, y: eY + 0.40, w: eW, h: 0.70, fontSize: 40, fontFace: FONT, color: COLORS.red, bold: true, align: "center", valign: "middle", margin: 0 });
    slide.addText("EMERGENCY", { x: eX, y: eY + 1.13, w: eW, h: 0.30, fontSize: 9, fontFace: FONT, color: COLORS.red, bold: true, align: "center", valign: "middle", margin: 0, charSpacing: 1 });
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.50, y: 3.05, w: 9.35, h: 1.20, fill: { color: COLORS.white }, line: { color: COLORS.cardBorder, width: 0.75 }, rectRadius: 0.08, shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "1F2E5A", opacity: 0.06 } });
    slide.addText((c.exampleLabel || "مثال على بلاغ واضح:"), { x: 3.70, y: 3.10, w: 9.00, h: 0.40, fontSize: 13, fontFace: FONT, color: COLORS.textLight, bold: true, align: "right", valign: "middle", margin: 0, rtlMode: true });
    slide.addText('"' + (c.exampleText || "") + '"', { x: 3.70, y: 3.50, w: 9.00, h: 0.65, fontSize: 14, fontFace: FONT, color: COLORS.primaryDark, bold: true, italic: true, align: "right", valign: "middle", margin: 0, rtlMode: true });
    const it = c.items; const cw = 2.95, ch = 1.20, gx = 0.10; const sx = (SLIDE_W - (4 * cw + 3 * gx)) / 2;
    for (let i = 0; i < it.length; i++) await addCompactIconCard(slide, sx + i * (cw + gx), 4.85, cw, ch, it[i].icon, it[i].title);
    rem(slide, c.reminder); addFooter(slide);
  }
  async function renderDoDont(c) {
    const slide = newSlide(); addSlideTitle(slide, c.title, c.subtitle);
    const colW = 6.10, colH = 4.10, colY = 2.85, itStartY = colY + 0.70, itH = 0.42, itGap = 0.08;
    // افعل (يمين)
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.73, y: colY, w: colW, h: colH, fill: { color: COLORS.white }, line: { color: COLORS.cardBorder, width: 0.75 }, rectRadius: 0.08 });
    slide.addShape(pres.shapes.RECTANGLE, { x: 6.73, y: colY, w: colW, h: 0.55, fill: { color: COLORS.greenLight }, line: { type: "none" } });
    slide.addText([{ text: "افعل  ", options: { color: COLORS.green, bold: true } }, { text: "✓", options: { color: COLORS.green, bold: true } }], { x: 6.93, y: colY, w: colW - 0.40, h: 0.55, fontSize: 18, fontFace: FONT, align: "right", valign: "middle", margin: 0, rtlMode: true });
    for (let i = 0; i < c.do.length; i++) {
      slide.addShape(pres.shapes.OVAL, { x: 6.93 + colW - 0.65, y: itStartY + i * (itH + itGap) + 0.05, w: 0.32, h: 0.32, fill: { color: COLORS.white }, line: { color: COLORS.green, width: 1.5 } });
      slide.addText(String(i + 1), { x: 6.93 + colW - 0.65, y: itStartY + i * (itH + itGap) + 0.05, w: 0.32, h: 0.32, fontSize: 10, fontFace: FONT, color: COLORS.green, bold: true, align: "center", valign: "middle", margin: 0 });
      slide.addText(c.do[i], { x: 6.93, y: itStartY + i * (itH + itGap), w: colW - 0.85, h: itH, fontSize: 13, fontFace: FONT, color: COLORS.primaryDark, align: "right", valign: "middle", margin: 0, rtlMode: true });
    }
    // لا تفعل (يسار)
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.50, y: colY, w: colW, h: colH, fill: { color: COLORS.white }, line: { color: COLORS.cardBorder, width: 0.75 }, rectRadius: 0.08 });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.50, y: colY, w: colW, h: 0.55, fill: { color: COLORS.redLight }, line: { type: "none" } });
    slide.addText([{ text: "لا تفعل  ", options: { color: COLORS.red, bold: true } }, { text: "✕", options: { color: COLORS.red, bold: true } }], { x: 0.70, y: colY, w: colW - 0.40, h: 0.55, fontSize: 18, fontFace: FONT, align: "right", valign: "middle", margin: 0, rtlMode: true });
    for (let i = 0; i < c.dont.length; i++) {
      slide.addShape(pres.shapes.OVAL, { x: 0.70 + colW - 0.65, y: itStartY + i * (itH + itGap) + 0.05, w: 0.32, h: 0.32, fill: { color: COLORS.white }, line: { color: COLORS.red, width: 1.5 } });
      slide.addText(String(i + 1), { x: 0.70 + colW - 0.65, y: itStartY + i * (itH + itGap) + 0.05, w: 0.32, h: 0.32, fontSize: 10, fontFace: FONT, color: COLORS.red, bold: true, align: "center", valign: "middle", margin: 0 });
      slide.addText(c.dont[i], { x: 0.70, y: itStartY + i * (itH + itGap), w: colW - 0.85, h: itH, fontSize: 13, fontFace: FONT, color: COLORS.primaryDark, align: "right", valign: "middle", margin: 0, rtlMode: true });
    }
    addFooter(slide);
  }
  async function renderScenarios(c) {
    const slide = newSlide(); addSlideTitle(slide, c.title, c.subtitle);
    const sc = c.scenarios; const cw = 6.10, ch = 1.85, gy = 0.15;
    for (let i = 0; i < sc.length; i++) {
      const col = i % 2, row = Math.floor(i / 2); const x = col === 0 ? 6.73 : 0.50; const y = 2.85 + row * (ch + gy);
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, fill: { color: COLORS.white }, line: { color: COLORS.cardBorder, width: 0.75 }, rectRadius: 0.08, shadow: { type: "outer", blur: 8, offset: 2, angle: 90, color: "1F2E5A", opacity: 0.06 } });
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.20, y: y + 0.18, w: 1.20, h: 0.42, fill: { color: COLORS.primary }, line: { type: "none" }, rectRadius: 0.05 });
      slide.addText("سيناريو " + (sc[i].num || (i + 1)), { x: x + 0.20, y: y + 0.18, w: 1.20, h: 0.42, fontSize: 12, fontFace: FONT, color: COLORS.white, bold: true, align: "center", valign: "middle", margin: 0, rtlMode: true });
      slide.addText(sc[i].q, { x: x + 1.50, y: y + 0.15, w: cw - 1.70, h: 0.45, fontSize: 14, fontFace: FONT, color: COLORS.primaryDark, bold: true, align: "right", valign: "middle", margin: 0, rtlMode: true });
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.20, y: y + 0.75, w: cw - 0.40, h: 0.95, fill: { color: COLORS.greenLight }, line: { color: COLORS.green, width: 0.5 }, rectRadius: 0.05 });
      slide.addText([{ text: "التصرف الصحيح: ", options: { bold: true, color: COLORS.green, lang: "ar-AE" } }, { text: sc[i].a, options: { color: COLORS.primaryDark, lang: "ar-AE" } }], { x: x + 0.35, y: y + 0.78, w: cw - 0.70, h: 0.90, fontSize: 12, fontFace: FONT, align: "right", valign: "middle", margin: 0, rtlMode: true, lang: "ar-AE" });
    }
    addFooter(slide);
  }
  async function renderQuiz(c) {
    const slide = newSlide(); addSlideTitle(slide, c.title, c.subtitle);
    const qa = c.qa; const iw = 12.30, ih = 0.62, gy = 0.04, sy = 2.85;
    for (let i = 0; i < qa.length; i++) {
      const y = sy + i * (ih + gy);
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.50, y, w: iw, h: ih, fill: { color: COLORS.white }, line: { color: COLORS.cardBorder, width: 0.75 }, rectRadius: 0.05 });
      slide.addShape(pres.shapes.OVAL, { x: 0.50 + iw - 0.55, y: y + 0.11, w: 0.40, h: 0.40, fill: { color: COLORS.primary }, line: { type: "none" } });
      slide.addImage({ data: iconSvg("FaCheck", "FFFFFF", 256), x: 0.50 + iw - 0.48, y: y + 0.18, w: 0.26, h: 0.26 });
      slide.addText([{ text: "س" + (i + 1) + ": ", options: { bold: true, color: COLORS.primary, lang: "ar-AE" } }, { text: qa[i].q, options: { color: COLORS.primaryDark, bold: true, lang: "ar-AE" } }], { x: 0.70, y: y + 0.05, w: iw - 0.85, h: 0.30, fontSize: 13, fontFace: FONT, align: "right", valign: "middle", margin: 0, rtlMode: true, lang: "ar-AE" });
      slide.addText([{ text: "الإجابة: ", options: { bold: true, color: COLORS.green, lang: "ar-AE" } }, { text: qa[i].a, options: { color: COLORS.textLight, lang: "ar-AE" } }], { x: 0.70, y: y + 0.32, w: iw - 0.85, h: 0.27, fontSize: 11, fontFace: FONT, align: "right", valign: "middle", margin: 0, rtlMode: true, lang: "ar-AE" });
    }
    addFooter(slide);
  }
  async function renderClosing(c) {
    const slide = pres.addSlide(); slide.background = { color: COLORS.bg }; addHeader(slide); addBgDecorations(slide);
    const cx = 6.665;
    slide.addShape(pres.shapes.OVAL, { x: cx - 1.10, y: 1.20, w: 2.20, h: 2.20, fill: { color: COLORS.primaryLight, transparency: 50 }, line: { type: "none" } });
    slide.addShape(pres.shapes.OVAL, { x: cx - 0.80, y: 1.40, w: 1.60, h: 1.60, fill: { color: COLORS.primary }, line: { type: "none" } });
    slide.addImage({ data: iconSvg(c.icon || "FaBriefcase", "FFFFFF", 512), x: cx - 0.50, y: 1.70, w: 1.00, h: 1.00 });
    slide.addText(c.title || "شكرًا لكم", { x: 0.50, y: 3.80, w: 12.30, h: 0.85, fontSize: 44, fontFace: FONT, color: COLORS.primaryDark, bold: true, align: "center", valign: "middle", margin: 0, rtlMode: true });
    if (c.subtitle) slide.addText(c.subtitle, { x: 0.50, y: 4.70, w: 12.30, h: 0.45, fontSize: 17, fontFace: FONT, color: COLORS.textLight, align: "center", valign: "middle", margin: 0, rtlMode: true });
    slide.addShape(pres.shapes.LINE, { x: 6.165, y: 5.20, w: 1.0, h: 0, line: { color: COLORS.primary, width: 2 } });
    if (c.points) slide.addText(c.points.join(" • "), { x: 0.50, y: 5.45, w: 12.30, h: 0.40, fontSize: 14, fontFace: FONT, color: COLORS.primaryDark, bold: true, align: "center", valign: "middle", margin: 0, rtlMode: true });
    if (c.quote) {
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.50, y: 5.95, w: 8.30, h: 0.75, fill: { color: "EEF3FA" }, line: { color: COLORS.primary, width: 0.75 }, rectRadius: 0.05 });
      slide.addText('"' + c.quote + '"', { x: 2.50, y: 5.95, w: 8.30, h: 0.75, fontSize: 13, fontFace: FONT, color: COLORS.primary, bold: true, italic: true, align: "center", valign: "middle", margin: 0, rtlMode: true });
    }
    addFooter(slide);
  }


  async function renderSection(c) {
    const slide = pres.addSlide(); slide.background = { color: COLORS.bg }; addHeader(slide); addBgDecorations(slide);
    const cx = 6.665;
    slide.addShape(pres.shapes.OVAL, { x: cx - 0.85, y: 1.75, w: 1.70, h: 1.70, fill: { color: COLORS.primaryLight, transparency: 55 }, line: { type: "none" } });
    slide.addShape(pres.shapes.OVAL, { x: cx - 0.62, y: 1.98, w: 1.24, h: 1.24, fill: { color: COLORS.primary }, line: { type: "none" } });
    slide.addImage({ data: iconSvg(c.icon || "FaShieldAlt", "FFFFFF", 512), x: cx - 0.38, y: 2.22, w: 0.76, h: 0.76 });
    if (c.label) slide.addText(c.label, { x: 0.50, y: 3.65, w: 12.30, h: 0.45, fontSize: 16, fontFace: FONT, color: COLORS.red, bold: true, align: "center", valign: "middle", margin: 0, rtlMode: true });
    slide.addText(c.title, { x: 0.50, y: 4.15, w: 12.30, h: 0.80, fontSize: 30, fontFace: FONT, color: COLORS.primaryDark, bold: true, align: "center", valign: "middle", margin: 0, rtlMode: true });
    if (c.subtitle) slide.addText(c.subtitle, { x: 0.50, y: 5.05, w: 12.30, h: 0.45, fontSize: 15, fontFace: FONT, color: COLORS.textLight, align: "center", valign: "middle", margin: 0, rtlMode: true });
    addFooter(slide);
  }
  async function renderList(c) {
    const slide = newSlide(); addSlideTitle(slide, c.title, c.subtitle);
    let topY = 2.75;
    if (c.intro) {
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.50, y: topY, w: 12.30, h: 0.85, fill: { color: COLORS.primaryDark }, line: { type: "none" }, rectRadius: 0.06, shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "1F2E5A", opacity: 0.12 } });
      slide.addText(c.intro, { x: 0.75, y: topY, w: 11.80, h: 0.85, fontSize: 13.5, fontFace: FONT, color: COLORS.white, align: "right", valign: "middle", margin: 0, rtlMode: true });
      topY += 1.05;
    }
    const pts = c.points || [];
    const limitY = c.reminder ? 6.10 : 6.90;
    const step = Math.min(0.52, (limitY - topY - 0.30) / Math.max(1, pts.length));
    const cardH = 0.25 + pts.length * step;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.50, y: topY, w: 12.30, h: cardH, fill: { color: COLORS.white }, line: { color: COLORS.cardBorder, width: 0.75 }, rectRadius: 0.08, shadow: { type: "outer", blur: 8, offset: 2, angle: 90, color: "1F2E5A", opacity: 0.06 } });
    for (let i = 0; i < pts.length; i++) {
      const y = topY + 0.15 + i * step;
      slide.addShape(pres.shapes.OVAL, { x: 12.28, y: y + 0.17, w: 0.14, h: 0.14, fill: { color: COLORS.primary }, line: { type: "none" } });
      slide.addText(pts[i], { x: 0.80, y, w: 11.35, h: 0.48, fontSize: 14, fontFace: FONT, color: COLORS.primaryDark, align: "right", valign: "middle", margin: 0, rtlMode: true });
    }
    rem(slide, c.reminder); addFooter(slide);
  }


  async function renderRefs(c) {
    const slide = newSlide(); addSlideTitle(slide, c.title || "المصادر بصيغة APA", c.subtitle || "قائمة المراجع الرسمية");
    const refs = c.refs || [];
    const topY = 2.80, step = 0.58;
    const cardH = Math.min(3.85, 0.30 + refs.length * step);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.50, y: topY, w: 12.30, h: cardH, fill: { color: COLORS.white }, line: { color: COLORS.cardBorder, width: 0.75 }, rectRadius: 0.08, shadow: { type: "outer", blur: 8, offset: 2, angle: 90, color: "1F2E5A", opacity: 0.06 } });
    for (let i = 0; i < refs.length; i++) {
      slide.addText(refs[i], { x: 0.85, y: topY + 0.18 + i * step, w: 11.60, h: step - 0.05, fontSize: 10.5, fontFace: FONT, color: COLORS.primaryDark, align: "left", valign: "middle", margin: 0 });
    }
    rem(slide, c.reminder); addFooter(slide);
  }

  const renderers = { section: renderSection, list: renderList, refs: renderRefs, cover: renderCover, cards: renderCards, compact: renderCompact, cardsDesc: renderCardsDesc, numbered: renderNumbered, stairs: renderStairs, emergency: renderEmergency, dodont: renderDoDont, scenarios: renderScenarios, quiz: renderQuiz, closing: renderClosing };

  if (content.cover) await renderCover(content.cover);
  for (const s of (content.slides || [])) {
    const fn = renderers[s.type];
    if (!fn) { console.error("Unknown slide type:", s.type); continue; }
    await fn(s);
  }
  const outPath = content.output || ("./" + CODE.replace(/\./g, "") + "_ADCDA.pptx");
  await pres.writeFile({ fileName: outPath });
  console.log("✅ Saved:", outPath, "| slides:", (content.cover ? 1 : 0) + (content.slides || []).length);
}

const file = process.argv[2] || "content.json";
const content = JSON.parse(fs.readFileSync(file, "utf8"));
build(content).catch(e => { console.error(e); process.exit(1); });
