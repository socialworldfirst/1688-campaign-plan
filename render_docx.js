// content.json -> Word file. node render_docx.js content.json out.docx
const fs = require("fs");
const path = require("path");
const d = require("docx");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, AlignmentType, HeadingLevel, LevelFormat, BorderStyle, Header, Footer,
  PageNumber, TabStopType,
} = d;

const src = process.argv[2] || path.join(__dirname, "content.json");
const out = process.argv[3] || path.join(__dirname, "pack.docx");
const doc = JSON.parse(fs.readFileSync(src, "utf8"));

const FONT = { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei", cs: "Arial" };
const GREY = "8A9199", INK = "1F2323", MUTED = "5B626A", ROSE = "FF0051", LINE = "E6E8EB";
const PAGE_W = 11906, MARGIN = 1134, TEXT_W = PAGE_W - 2 * MARGIN; // 9638 DXA

// ---- inline: **bold** and [brackets]
function runs(s, base = {}) {
  const parts = s.split(/(\*\*[^*]+\*\*|\[[^\]]+\])/g).filter(Boolean);
  return parts.map((p) => {
    if (/^\*\*[^*]+\*\*$/.test(p)) return new TextRun({ ...base, text: p.slice(2, -2), bold: true });
    if (/^\[[^\]]+\]$/.test(p)) return new TextRun({ ...base, text: p, color: GREY, italics: true });
    return new TextRun({ ...base, text: p });
  });
}
const P = (s, opts = {}, base = {}) => new Paragraph({ ...opts, children: runs(s, base) });
const lead = (s) => P(s, { style: "Lead" });
const cue = (s) => P(s, { style: "Cue" });
const label = (s) => new Paragraph({ style: "Label", children: [new TextRun({ text: s.toUpperCase() })] });
const line = (s) => P(s, { style: "Line" });
const zh = (s) => P(s, { style: "ZH" });
const note = (s) => new Paragraph({
  style: "Note",
  border: { left: { style: BorderStyle.SINGLE, size: 12, color: ROSE, space: 8 } },
  children: runs(s),
});
const bullet = (s, style) => new Paragraph({ style: style || undefined, numbering: { reference: "bullets", level: 0 }, children: runs(s) });
const h1 = (s) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, children: [new TextRun(s)] });
const h2 = (s) => new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, children: [new TextRun(s)] });
const h3 = (s) => new Paragraph({ heading: HeadingLevel.HEADING_3, keepNext: true, children: [new TextRun(s)] });

const border = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border };
const margins = { top: 100, bottom: 100, left: 120, right: 120 };

function cellOf(text, width, opts = {}) {
  const paras = String(text).split("\n").map((t) => P(t, {}, { size: opts.size || 19, bold: !!opts.bold, color: opts.color || INK }));
  return new TableCell({ width: { size: width, type: WidthType.DXA }, margins, shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill, color: "auto" } : undefined, children: paras });
}
function table3(cols, rows) {
  const W = [2200, 1500, TEXT_W - 3700];
  const header = new TableRow({ tableHeader: true, children: cols.map((c, i) => cellOf(c, W[i], { bold: true, fill: "F2F2F2", size: 17, color: MUTED })) });
  const body = rows.map((r) => {
    const x = r.cells;
    const first = (r.audience === "internal" ? "(internal) " : "") + x[0];
    return new TableRow({ cantSplit: true, children: [cellOf(first, W[0], { bold: true }), cellOf(x[1], W[1], { color: MUTED }), cellOf(x[2], W[2])] });
  });
  return new Table({ width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: W, borders, rows: [header, ...body] });
}
function table2(rows) {
  const W = [2300, TEXT_W - 2300];
  const body = rows.map((r) => {
    const lab = (r.audience === "internal" ? "(internal) " : "") + r.label;
    return new TableRow({ cantSplit: true, children: [cellOf(lab, W[0], { bold: true }), cellOf(r.value, W[1])] });
  });
  return new Table({ width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: W, borders, rows: body });
}

function renderBlock(b) {
  const internal = b.audience === "internal";
  const els = [];
  switch (b.t) {
    case "p": els.push(lead(b.text)); break;
    case "h3": els.push(h3(b.text + (internal ? " · internal" : ""))); break;
    case "note": els.push(note(b.text)); break;
    case "list": b.items.forEach((i) => els.push(bullet(i))); break;
    case "rows": els.push(table2(b.rows)); els.push(new Paragraph({ spacing: { after: 60 } })); break;
    case "table": els.push(table3(b.cols, b.rows)); els.push(new Paragraph({ spacing: { after: 60 } })); break;
    case "script":
      els.push(h3(`${b.id} · ${b.title}${internal ? " · internal" : ""}`));
      els.push(cue(b.meta));
      b.lines.forEach((l) => els.push(line(l)));
      (b.notes || []).forEach((n) => els.push(bullet(n, "Cue")));
      break;
    case "sheet":
      els.push(h2(`${b.id} · ${b.title}`));
      els.push(cue(b.track + (b.time ? ` · ${b.time}` : "")));
      if (b.coldopen) { els.push(label("On screen")); els.push(cue(`“${b.coldopen}”` + (b.source ? ` · ${b.source}` : ""))); }
      els.push(label("Aria asks")); els.push(P(b.asks, { style: "Ask" }));
      els.push(label("Ary reads")); b.lines.forEach((l) => els.push(line(l)));
      if (b.zh) { els.push(label("中文提示")); els.push(zh(b.zh)); }
      if (b.fills && b.fills.length) { els.push(label("1688 fills in")); b.fills.forEach((f) => els.push(bullet(f, "Cue"))); }
      break;
    default: throw new Error("block " + b.t);
  }
  return els;
}

const children = [];
children.push(new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun(doc.title)] }));
children.push(P(doc.sub, { style: "Lead" }));
children.push(P(doc.meta, { style: "Cue" }));
children.push(new Paragraph({ spacing: { after: 200 } }));
children.push(P("Contents", { style: "Label" }));
doc.sections.forEach((s) => children.push(P(`${s.num}  ${s.title}${s.audience === "internal" ? " · internal" : ""}`, { style: "Lead" })));

doc.sections.forEach((s) => {
  children.push(h1(`${s.num} · ${s.title}${s.audience === "internal" ? " · internal" : ""}`));
  children.push(lead(s.lead));
  s.blocks.forEach((b) => renderBlock(b).forEach((e) => children.push(e)));
});

const document = new Document({
  creator: "WorldFirst",
  title: doc.title,
  styles: {
    default: { document: { run: { font: FONT, size: 21, color: INK } } },
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal", run: { size: 52, bold: true, font: FONT, color: INK }, paragraph: { spacing: { after: 200 } } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 34, bold: true, font: FONT, color: INK }, paragraph: { spacing: { before: 120, after: 160 }, keepNext: true, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, font: FONT, color: INK }, paragraph: { spacing: { before: 360, after: 80 }, keepNext: true, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 22, bold: true, font: FONT, color: ROSE }, paragraph: { spacing: { before: 300, after: 80 }, keepNext: true, outlineLevel: 2 } },
      { id: "Lead", name: "Lead", basedOn: "Normal", run: { size: 21, font: FONT, color: MUTED }, paragraph: { spacing: { after: 140, line: 300 } } },
      { id: "Line", name: "Line", basedOn: "Normal", run: { size: 26, font: FONT, color: INK }, paragraph: { spacing: { after: 90, line: 320 }, keepLines: true } },
      { id: "Ask", name: "Ask", basedOn: "Normal", run: { size: 22, italics: true, font: FONT, color: INK }, paragraph: { spacing: { after: 100 } } },
      { id: "Cue", name: "Cue", basedOn: "Normal", run: { size: 18, italics: true, font: FONT, color: GREY }, paragraph: { spacing: { after: 100 } } },
      { id: "Label", name: "Label", basedOn: "Normal", run: { size: 16, bold: true, font: FONT, color: ROSE, characterSpacing: 20 }, paragraph: { spacing: { before: 140, after: 40 }, keepNext: true } },
      { id: "ZH", name: "ZH", basedOn: "Normal", run: { size: 20, font: FONT, color: MUTED }, paragraph: { spacing: { after: 100 } } },
      { id: "Note", name: "Note", basedOn: "Normal", run: { size: 19, font: FONT, color: MUTED }, paragraph: { spacing: { before: 120, after: 160 }, indent: { left: 160 } } },
    ],
  },
  numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 260 }, spacing: { after: 60 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: `${doc.title} · ${doc.date}`, size: 16, color: GREY, font: FONT })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 16, color: GREY, font: FONT })] })] }) },
    children,
  }],
});

Packer.toBuffer(document).then((buf) => { fs.writeFileSync(out, buf); console.log("wrote", out, buf.length, "bytes"); });
