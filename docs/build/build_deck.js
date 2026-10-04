// Build docs/str_term_validator_overview.pptx   (run: node docs/build/build_deck.js)
const path = require("path");
const pptxgen = require("pptxgenjs");
const NAVY = "1E2761", ICE = "CADCFC", WHITE = "FFFFFF";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Neeraj Khandelwal, Inquestica Consulting";
pres.title = "How str_term_validator.py Works";
const slide = pres.addSlide();
slide.background = { color: NAVY };

slide.addText("How str_term_validator.py Works", { x: 0.6, y: 0.32, w: 12.1, h: 0.7, margin: 0,
  fontFace: "Cambria", fontSize: 32, bold: true, color: WHITE, isTextBox: true });
slide.addText("Turning raw proposed search terms into validated, self-verified dtSearch queries for a Relativity Search Terms Report",
  { x: 0.6, y: 1.02, w: 12.1, h: 0.4, margin: 0, fontFace: "Calibri", fontSize: 13.5, color: ICE, isTextBox: true });

const steps = [
  ["Read Terms", "Reads .xlsx / .csv / .txt, one term per line — each line is one dtSearch query"],
  ["Detect Issues", "40+ checks: dtSearch syntax, STR limits, metadata field:value, noise words"],
  ["Auto-Repair", "Fixes every repairable defect; extracts metadata to the saved search"],
  ["Self-Verify", "Re-validates the revised term until stable: CLEAN / ADVISORY / NEEDS REVIEW"],
  ["Report", "Color-coded Excel report, extracted conditions, clean STR-ready term list"],
];
const colW = 2.3, gap = 0.1;
steps.forEach((s, i) => {
  const xBase = 0.7 + i * (colW + gap);
  const cx = xBase + colW / 2 - 0.4;
  slide.addShape(pres.ShapeType.ellipse, { x: cx, y: 1.8, w: 0.8, h: 0.8, fill: { color: ICE } });
  slide.addText(String(i + 1), { x: cx, y: 1.8, w: 0.8, h: 0.8, margin: 0, align: "center", valign: "middle",
    fontFace: "Calibri", fontSize: 22, bold: true, color: NAVY, isTextBox: true });
  slide.addText(s[0], { x: xBase, y: 2.72, w: colW, h: 0.32, margin: 0, align: "center",
    fontFace: "Calibri", fontSize: 13, bold: true, color: WHITE, isTextBox: true });
  slide.addText(s[1], { x: xBase, y: 3.08, w: colW, h: 0.95, margin: 0, align: "center",
    fontFace: "Calibri", fontSize: 10, color: ICE, isTextBox: true });
  if (i < steps.length - 1) {
    slide.addText("\u203A", { x: xBase + colW + gap / 2 - 0.25, y: 1.95, w: 0.5, h: 0.5, margin: 0,
      align: "center", fontFace: "Calibri", fontSize: 22, bold: true, color: ICE, isTextBox: true });
  }
});

slide.addText("KEY DESIGN CHOICES", { x: 0.6, y: 4.42, w: 8.0, h: 0.4, margin: 0,
  fontFace: "Calibri", fontSize: 18, bold: true, color: WHITE, isTextBox: true });
const cards = [
  ["A", "Fixpoint Repair", "The revised term is re-checked until nothing changes \u2014 so a second validation run stays clean."],
  ["B", "Metadata Extraction", "custodian:, date:, ext: aren't dtSearch \u2014 conditions move to the searchable-set saved search, listed on their own sheet."],
  ["C", "Flagged Assumptions", "Every default the script chooses (w/5 distance, w/1 for w/0) is labeled ASSUMPTION APPLIED for case-team sign-off."],
  ["D", "Severity Model", "ERROR must be fixed \u00B7 WARNING needs a recorded decision \u00B7 INFO is a note. NEEDS REVIEW terms never reach the clean list."],
];
cards.forEach((c, i) => {
  const x = i % 2 === 0 ? 0.6 : 6.8;
  const y = i < 2 ? 4.92 : 6.02;
  slide.addShape(pres.ShapeType.roundRect, { x, y, w: 5.9, h: 0.95, rectRadius: 0.08,
    fill: { color: WHITE, transparency: 90 } });
  slide.addShape(pres.ShapeType.ellipse, { x: x + 0.2, y: y + 0.22, w: 0.5, h: 0.5, fill: { color: ICE } });
  slide.addText(c[0], { x: x + 0.2, y: y + 0.22, w: 0.5, h: 0.5, margin: 0, align: "center", valign: "middle",
    fontFace: "Calibri", fontSize: 16, bold: true, color: NAVY, isTextBox: true });
  slide.addText(c[1], { x: x + 0.85, y: y + 0.12, w: 4.9, h: 0.3, margin: 0,
    fontFace: "Calibri", fontSize: 13, bold: true, color: WHITE, isTextBox: true });
  slide.addText(c[2], { x: x + 0.85, y: y + 0.42, w: 4.9, h: 0.5, margin: 0,
    fontFace: "Calibri", fontSize: 10, color: ICE, isTextBox: true });
});
slide.addText("Input: proposed term list (.xlsx / .csv / .txt)   \u2022   Output: validation report (.xlsx) + extracted conditions + clean STR-ready term list (.txt)",
  { x: 0.6, y: 7.08, w: 12.1, h: 0.3, margin: 0, fontFace: "Calibri", fontSize: 10, color: ICE, isTextBox: true });

pres.writeFile({ fileName: path.join(__dirname, "..", "str_term_validator_overview.pptx") })
  .then((f) => console.log("wrote", f));
