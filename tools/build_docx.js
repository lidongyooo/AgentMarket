#!/usr/bin/env node
// 把 book/ 下的 Markdown 书稿排版成 Word 文档（封面 + 自动目录 + 章节样式）。
// 用法：node tools/build_docx.js [输出路径]
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow,
  TableCell, WidthType, BorderStyle, ShadingType, LevelFormat, TableOfContents, Header,
  Footer, PageNumber, TableLayoutType, LineRuleType,
} = require("docx");
const { estimatePageNumbersWith } = require("docx/layout");

const ROOT = path.resolve(__dirname, "..");
const BOOK_DIR = path.join(ROOT, "book");
const OUT = path.resolve(process.argv[2] || path.join(ROOT, "会干活的AI怎么卖-全本.docx"));

// ---------- 字体与版式 ----------
const SONG = "SimSun"; // 宋体：正文
const HEI = "SimHei"; // 黑体：标题
const KAI = "KaiTi"; // 楷体：引文
const LATIN = "Times New Roman";
const LATIN_SANS = "Arial";
const font = (eastAsia, latin = LATIN) => ({ ascii: latin, hAnsi: latin, cs: latin, eastAsia, hint: "eastAsia" });

const PAGE_W = 11906; // A4
const PAGE_H = 16838;
const MARGIN_LR = 1701; // 3 cm
const MARGIN_TB = 1440; // 2.54 cm
const TEXT_W = PAGE_W - 2 * MARGIN_LR;
const BODY_SIZE = 22; // 11 pt
const INDENT_2CH = 2 * BODY_SIZE * 10; // 首行缩进两个字符（twips）

// ---------- 行内格式 ----------
function curlyQuotes(s) {
  const n = (s.match(/"/g) || []).length;
  if (n === 0 || n % 2 !== 0) return s;
  let open = true;
  return s.replace(/"/g, () => {
    const q = open ? "“" : "”";
    open = !open;
    return q;
  });
}

function runs(text, extra = {}) {
  const out = [];
  const src = curlyQuotes(text);
  const re = /\*\*(.+?)\*\*/g;
  let last = 0;
  let m;
  while ((m = re.exec(src))) {
    if (m.index > last) out.push(new TextRun({ text: src.slice(last, m.index), ...extra }));
    out.push(new TextRun({ text: m[1], ...extra, bold: true }));
    last = re.lastIndex;
  }
  if (last < src.length) out.push(new TextRun({ text: src.slice(last), ...extra }));
  return out;
}

// ---------- Markdown 解析 ----------
function parseBlocks(lines) {
  const blocks = [];
  let para = null;
  const flush = () => {
    if (para) blocks.push(para);
    para = null;
  };
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].replace(/\s+$/, "");
    if (line.trim() === "") { flush(); continue; }
    let m;
    if ((m = line.match(/^(#{1,6})\s+(.*)$/))) {
      flush();
      blocks.push({ type: "heading", depth: m[1].length, text: m[2].trim() });
    } else if (/^---+$/.test(line.trim())) {
      flush();
      blocks.push({ type: "hr" });
    } else if (line.startsWith(">")) {
      flush();
      const inner = [];
      while (i < lines.length && lines[i].startsWith(">")) {
        inner.push(lines[i].replace(/^>\s?/, ""));
        i++;
      }
      i--;
      blocks.push({ type: "quote", blocks: parseBlocks(inner) });
    } else if (line.trim().startsWith("|")) {
      flush();
      const rows = [];
      while (i < lines.length && lines[i].trim().startsWith("|")) {
        const cells = lines[i].trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map((c) => c.trim());
        if (!cells.every((c) => /^:?-{2,}:?$/.test(c))) rows.push(cells);
        i++;
      }
      i--;
      blocks.push({ type: "table", rows });
    } else if ((m = line.match(/^-\s+\[[ xX]\]\s+(.*)$/))) {
      flush();
      blocks.push({ type: "check", text: m[1] });
    } else if ((m = line.match(/^[-*]\s+(.*)$/))) {
      flush();
      blocks.push({ type: "bullet", text: m[1] });
    } else if ((m = line.match(/^(\d+)\.\s+(.*)$/))) {
      flush();
      blocks.push({ type: "ordered", num: m[1], text: m[2] });
    } else if (para) {
      para.parts.push(line.trim());
    } else {
      para = { type: "para", parts: [line.trim()] };
    }
  }
  flush();
  return blocks;
}

// ---------- 渲染 ----------
function textWidth(s) {
  let w = 0;
  for (const ch of s.replace(/\*\*/g, "")) w += /[⺀-鿿＀-￯　-〿“”]/.test(ch) ? 2 : 1;
  return w;
}

function renderTable(rows, width) {
  const cols = Math.max(...rows.map((r) => r.length));
  const norm = rows.map((r) => [...r, ...Array(cols - r.length).fill("")]);
  const allBlank = norm.slice(1).every((r) => r.every((c) => c === ""));
  const weights = Array.from({ length: cols }, (_, j) => {
    const w = Math.max(...norm.map((r) => textWidth(r[j])));
    const blank = norm.slice(1).every((r) => r[j] === "");
    if (allBlank) return Math.max(textWidth(norm[0][j]), 6) + 4 + (j === 0 ? 14 : 0);
    return blank ? 24 : Math.min(Math.max(w, 6), 36) + 4;
  });
  const total = weights.reduce((a, b) => a + b, 0);
  const widths = weights.map((w) => Math.floor((width * w) / total));
  widths[cols - 1] += width - widths.reduce((a, b) => a + b, 0);
  const line = { style: BorderStyle.SINGLE, size: 4, color: "9AA5B1" };
  return new Table({
    width: { size: width, type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED,
    borders: { top: line, bottom: line, left: line, right: line, insideHorizontal: line, insideVertical: line },
    rows: norm.map((r, i) => new TableRow({
      tableHeader: i === 0,
      cantSplit: true,
      children: r.map((c, j) => new TableCell({
        width: { size: widths[j], type: WidthType.DXA },
        shading: i === 0 ? { fill: "E3EAF2", type: ShadingType.CLEAR, color: "auto" } : undefined,
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        children: [new Paragraph({ style: "TableText", children: runs(c, i === 0 ? { bold: true } : {}) })],
      })),
    })),
  });
}

function renderBlocks(blocks, ctx, out, inQuote = false) {
  blocks.forEach((b, idx) => {
    switch (b.type) {
      case "heading": {
        out.push(renderHeading(b, ctx));
        break;
      }
      case "hr": {
        // 章节标题前的分隔线由分页代替；正文中的分隔线画成一条细线。
        const next = blocks.slice(idx + 1).find((x) => x.type !== "hr");
        if (!next || (next.type === "heading" && next.depth === 1) || out.length === 0) break;
        out.push(new Paragraph({
          spacing: { before: 120, after: 240 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "C3CBD3", space: 1 } },
          children: [],
        }));
        break;
      }
      case "quote": {
        const inner = [];
        renderBlocks(b.blocks, ctx, inner, true);
        out.push(...inner);
        break;
      }
      case "table":
        out.push(renderTable(b.rows, inQuote ? TEXT_W - 420 : TEXT_W));
        out.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
        break;
      case "bullet":
        out.push(new Paragraph({ style: inQuote ? "QuoteText" : "ListText", numbering: { reference: "bullets", level: 0 }, children: runs(b.text) }));
        break;
      case "check":
        out.push(new Paragraph({ style: "ListText", indent: { left: 640, hanging: 400 }, children: runs("□ " + b.text) }));
        break;
      case "ordered":
        out.push(new Paragraph({ style: inQuote ? "QuoteText" : "ListText", indent: { left: 640, hanging: 400 }, children: runs(`${b.num}. ${b.text}`) }));
        break;
      case "para": {
        const children = [];
        b.parts.forEach((part, k) => {
          const r = runs(part);
          if (k > 0 && /^\*\*[^*]+\*\*$/.test(b.parts[k - 1])) r.unshift(new TextRun({ break: 1 }));
          children.push(...r);
        });
        out.push(new Paragraph({ style: inQuote ? "QuoteText" : "BodyText", children }));
        break;
      }
      default:
        break;
    }
  });
}

function renderHeading(b, ctx) {
  const text = curlyQuotes(b.text);
  if (b.depth === 1 || ctx.forceTop) {
    ctx.forceTop = false;
    if (/^第[一二三四五六七八九十]+部分/.test(text)) {
      ctx.unit = "part";
      const cut = text.indexOf("　");
      const num = text.slice(0, cut);
      const title = text.slice(cut + 1).trim();
      return new Paragraph({
        heading: HeadingLevel.HEADING_1,
        pageBreakBefore: true,
        alignment: AlignmentType.CENTER,
        spacing: { before: 3600, after: 360, line: 600, lineRule: LineRuleType.AUTO },
        children: [
          new TextRun({ text: num, size: 32, color: "5B6B7C" }),
          new TextRun({ text: title, size: 40, break: 1 }),
        ],
      });
    }
    if (/^第\s*\d+\s*章/.test(text)) {
      ctx.unit = "chapter";
      const cut = text.indexOf("　");
      return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        pageBreakBefore: true,
        alignment: AlignmentType.CENTER,
        spacing: { before: 480, after: 600, line: 480, lineRule: LineRuleType.AUTO },
        children: [
          new TextRun({ text: text.slice(0, cut), size: 26, color: "5B6B7C" }),
          new TextRun({ text: text.slice(cut + 1).trim(), break: 1 }),
        ],
      });
    }
    ctx.unit = "front";
    const cut = text.indexOf("　");
    const children = cut < 0
      ? [new TextRun(text)]
      : [
        new TextRun({ text: text.slice(0, cut), size: 28, color: "5B6B7C" }),
        new TextRun({ text: text.slice(cut + 1).trim(), size: 36, break: 1 }),
      ];
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      pageBreakBefore: true,
      alignment: AlignmentType.CENTER,
      spacing: { before: 480, after: 600, line: 480, lineRule: LineRuleType.AUTO },
      children,
    });
  }
  // 章节内部：## → 3 级，### → 4 级，#### → 5 级（不进目录）
  const level = { 2: HeadingLevel.HEADING_3, 3: HeadingLevel.HEADING_4 }[b.depth] || HeadingLevel.HEADING_5;
  return new Paragraph({ heading: level, children: [new TextRun(text)] });
}

// ---------- 封面与前置部分 ----------
function cover() {
  const p = (text, opts = {}) => new Paragraph({ alignment: AlignmentType.CENTER, ...opts, children: [new TextRun({ text, ...opts.run })] });
  return [
    p("会干活的 AI 怎么卖", { spacing: { before: 3800, after: 360 }, run: { font: font(HEI, LATIN_SANS), size: 64, bold: true, color: "1F3A5F" } }),
    p("AI Agent 中国市场营销全书", { spacing: { after: 720 }, run: { font: font(HEI, LATIN_SANS), size: 36, color: "3D5A80" } }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      indent: { left: 2400, right: 2400 },
      border: { top: { style: BorderStyle.SINGLE, size: 8, color: "3D5A80", space: 1 } },
      spacing: { after: 360 },
      children: [],
    }),
    p("写给 Claude、Codex 类智能体产品的", { spacing: { after: 80 }, run: { font: font(KAI), size: 26, color: "444444" } }),
    p("创始人、产品、市场与销售团队", { spacing: { after: 4600 }, run: { font: font(KAI), size: 26, color: "444444" } }),
    p("2026 年 10 月", { run: { font: font(SONG), size: 24, color: "666666" } }),
  ];
}

function readingNotes(ctx) {
  // 00-目录.md 里手写的目录由自动目录代替，只保留"阅读说明"之后的内容。
  const src = fs.readFileSync(path.join(BOOK_DIR, "00-目录.md"), "utf8").split("\n");
  const start = src.findIndex((l) => l.startsWith("## 阅读说明"));
  const blocks = parseBlocks(src.slice(start)).filter((b) => b.type !== "hr");
  const out = [];
  blocks.forEach((b, i) => {
    if (b.type === "heading") {
      out.push(i === 0
        ? new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, children: [new TextRun(curlyQuotes(b.text))] })
        : new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(curlyQuotes(b.text))] }));
    } else {
      renderBlocks([b], ctx, out);
    }
  });
  return out;
}

// ---------- 样式 ----------
const headingStyle = (id, name, outlineLevel, size, extra = {}) => ({
  id, name, basedOn: "Normal", next: "BodyText", quickFormat: true,
  run: { font: font(HEI, LATIN_SANS), size, bold: true, color: extra.color || "1F3A5F" },
  paragraph: { outlineLevel, keepNext: true, keepLines: true, spacing: { line: 300, lineRule: LineRuleType.AUTO, ...extra.spacing }, alignment: extra.alignment },
});

const styles = {
  default: {
    document: { run: { font: font(SONG), size: BODY_SIZE }, paragraph: { spacing: { line: 360, lineRule: LineRuleType.AUTO } } },
  },
  paragraphStyles: [
    headingStyle("Heading1", "Heading 1", 0, 40, { spacing: { before: 600, after: 480 }, alignment: AlignmentType.CENTER }),
    headingStyle("Heading2", "Heading 2", 1, 36, { spacing: { before: 600, after: 600 }, alignment: AlignmentType.CENTER }),
    headingStyle("Heading3", "Heading 3", 2, 30, { spacing: { before: 480, after: 200 } }),
    headingStyle("Heading4", "Heading 4", 3, 26, { spacing: { before: 320, after: 140 }, color: "2E4A6B" }),
    headingStyle("Heading5", "Heading 5", 4, 23, { spacing: { before: 240, after: 100 }, color: "2E4A6B" }),
    {
      id: "BodyText", name: "Body Text", basedOn: "Normal", quickFormat: true,
      paragraph: { indent: { firstLine: INDENT_2CH }, spacing: { after: 120, line: 360, lineRule: LineRuleType.AUTO }, alignment: AlignmentType.JUSTIFIED },
    },
    {
      id: "ListText", name: "List Text", basedOn: "Normal",
      paragraph: { spacing: { after: 80, line: 340, lineRule: LineRuleType.AUTO }, alignment: AlignmentType.JUSTIFIED },
    },
    {
      id: "QuoteText", name: "Quote Text", basedOn: "Normal",
      run: { font: font(KAI), color: "3A3A3A" },
      paragraph: {
        indent: { left: 420, right: 210 },
        spacing: { before: 60, after: 120, line: 340, lineRule: LineRuleType.AUTO },
        border: { left: { style: BorderStyle.SINGLE, size: 18, color: "8FA3B8", space: 10 } },
      },
    },
    {
      id: "TableText", name: "Table Text", basedOn: "Normal",
      run: { size: 19 },
      paragraph: { spacing: { before: 0, after: 0, line: 300, lineRule: LineRuleType.AUTO } },
    },
    {
      id: "TOC1", name: "toc 1", basedOn: "Normal", next: "Normal",
      run: { font: font(HEI, LATIN_SANS), size: 24, bold: true },
      paragraph: { spacing: { before: 160, after: 40, line: 300, lineRule: LineRuleType.AUTO } },
    },
    {
      id: "TOC2", name: "toc 2", basedOn: "Normal", next: "Normal",
      run: { size: 22 },
      paragraph: { indent: { left: 440 }, spacing: { before: 0, after: 20, line: 300, lineRule: LineRuleType.AUTO } },
    },
  ],
};

const numbering = {
  config: [{
    reference: "bullets",
    levels: [{
      level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 640, hanging: 320 } } },
    }],
  }],
};

// ---------- 页码估算（用 Word 自带字体测量，让目录在打开前就有页码） ----------
function pageEstimator() {
  const dir = "/Applications/Microsoft Word.app/Contents/Resources/DFonts";
  const files = ["Simsun.ttc", "SimHei.ttf", "Kaiti.ttf", "times.ttf", "timesbd.ttf", "arial.ttf", "arialbd.ttf"];
  const fonts = files.filter((f) => fs.existsSync(path.join(dir, f))).map((f) => ({ data: fs.readFileSync(path.join(dir, f)) }));
  return estimatePageNumbersWith({ fonts, guess: true });
}

// ---------- 组装 ----------
function build() {
  const ctx = { unit: null };
  const body = [];

  body.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 360 },
    children: [new TextRun({ text: "目　录", font: font(HEI, LATIN_SANS), size: 40, bold: true, color: "1F3A5F" })],
  }));
  body.push(new TableOfContents("目录", { hyperlink: true, headingStyleRange: "1-2", beginDirty: true }));
  body.push(...readingNotes(ctx));

  const files = fs.readdirSync(BOOK_DIR).filter((f) => f.endsWith(".md") && !f.startsWith("00-")).sort();
  for (const f of files) {
    const lines = fs.readFileSync(path.join(BOOK_DIR, f), "utf8").split("\n");
    renderBlocks(parseBlocks(lines), ctx, body);
  }

  const header = new Header({
    children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "C3CBD3", space: 4 } },
      children: [new TextRun({ text: "会干活的 AI 怎么卖 · AI Agent 中国市场营销全书", size: 17, color: "7A8794" })],
    })],
  });
  const footer = new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: ["— ", PageNumber.CURRENT, " —"], size: 18, color: "555555" })],
    })],
  });
  const page = { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN_TB, bottom: MARGIN_TB, left: MARGIN_LR, right: MARGIN_LR, header: 720, footer: 720 } };

  return new Document({
    title: "会干活的 AI 怎么卖：AI Agent 中国市场营销全书",
    subject: "AI Agent 中国市场营销",
    creator: "",
    features: { updateFields: true },
    pageNumbers: pageEstimator(),
    styles,
    numbering,
    sections: [
      { properties: { page }, children: cover() },
      {
        properties: { page: { ...page, pageNumbers: { start: 1 } } },
        headers: { default: header },
        footers: { default: footer },
        children: body,
      },
    ],
  });
}

Packer.toBuffer(build()).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("已生成：" + OUT + `（${(buf.length / 1024).toFixed(0)} KB）`);
});
