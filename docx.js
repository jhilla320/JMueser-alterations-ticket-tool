// Hand-rolled OOXML .docx generator (via JSZip) so the tool never needs a
// server. Takes plain ticket data (not HTML) so the printed/exported layout
// can be designed independently from the on-screen preview.

const INK = "1C1F18";
const BRASS = "8A6A22";
const OXBLOOD = "7A2B22";
const SLATE = "6B6459";
const RULE = "D8CFBE";

// Word can't load web fonts, so we pick the closest widely-installed
// equivalents to the on-screen Newsreader/Inter/IBM Plex Mono system:
// a serif for display headings, a clean sans for everything else.
const FONT_DISPLAY = "Georgia";
const FONT_BODY = "Calibri";

function escapeXml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function run(text, { bold, italic, size, color, caps, font = FONT_BODY } = {}) {
  const props = [
    `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>`,
    bold ? "<w:b/>" : "",
    italic ? "<w:i/>" : "",
    caps ? "<w:caps/>" : "",
    size ? `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>` : "",
    color ? `<w:color w:val="${color}"/>` : "",
  ].join("");
  const rPr = props ? `<w:rPr>${props}</w:rPr>` : "";
  const preserve = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : "";
  return `<w:r>${rPr}<w:t${preserve}>${escapeXml(text)}</w:t></w:r>`;
}

function paragraph(runs, { before = 0, after = 120, border, align, shading, indent } = {}) {
  const borderXml = border ? `<w:pBdr><w:${border.side || "top"} w:val="single" w:sz="${border.sz || 6}" w:space="4" w:color="${border.color || RULE}"/></w:pBdr>` : "";
  const shadingXml = shading ? `<w:shd w:val="clear" w:color="auto" w:fill="${shading}"/>` : "";
  const jc = align ? `<w:jc w:val="${align}"/>` : "";
  const indXml = indent ? `<w:ind w:left="${indent}"/>` : "";
  const pPr = `<w:pPr><w:spacing w:before="${before}" w:after="${after}" w:line="288" w:lineRule="auto"/>${jc}${borderXml}${shadingXml}${indXml}</w:pPr>`;
  return `<w:p>${pPr}${(Array.isArray(runs) ? runs : [runs]).join("")}</w:p>`;
}

function tableCell(content, { width, shading, borderColor = RULE } = {}) {
  const tcBorders = `<w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="${borderColor}"/></w:tcBorders>`;
  const shd = shading ? `<w:shd w:val="clear" w:color="auto" w:fill="${shading}"/>` : "";
  const tcPr = `<w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${tcBorders}${shd}<w:tcMar><w:top w:w="60" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar></w:tcPr>`;
  return `<w:tc>${tcPr}${content}</w:tc>`;
}

function twoColTable(rows, { labelWidth = 3200, valueWidth = 7600 } = {}) {
  const tblPr = `<w:tblPr><w:tblW w:w="${labelWidth + valueWidth}" w:type="dxa"/><w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr>`;
  const grid = `<w:tblGrid><w:gridCol w:w="${labelWidth}"/><w:gridCol w:w="${valueWidth}"/></w:tblGrid>`;
  const body = rows
    .map(([label, value]) => {
      const labelCell = tableCell(paragraph(run(label, { bold: true, size: 18, color: SLATE, caps: true }), { after: 0 }), { width: labelWidth });
      const valueCell = tableCell(paragraph(run(value, { size: 20, color: INK }), { after: 0 }), { width: valueWidth });
      return `<w:tr>${labelCell}${valueCell}</w:tr>`;
    })
    .join("");
  return `<w:tbl>${tblPr}${grid}${body}</w:tbl>`;
}

function garmentSectionXml(section) {
  const parts = [];
  const INDENT = 260; // left indent (twips) for everything under the title, showing hierarchy
  parts.push(
    paragraph(run(section.label, { bold: true, size: 28, color: INK, caps: true }), {
      before: 200,
      after: 100,
      border: { side: "top", sz: 8, color: BRASS },
    }),
  );
  section.adjustments.forEach((line) => {
    parts.push(paragraph(run(line, { size: 20, color: INK }), { after: 60, indent: INDENT }));
  });
  section.attributes.forEach((a) => {
    parts.push(paragraph(run(`${a.label}: ${a.value}`, { size: 20, color: INK }), { after: 60, indent: INDENT }));
  });
  if (section.notes) {
    parts.push(paragraph(run("Notes", { bold: true, size: 18, color: SLATE, caps: true }), { before: 120, after: 40, indent: INDENT }));
    section.notes.split("\n").filter(Boolean).forEach((line) => {
      parts.push(paragraph(run(`•  ${line}`, { size: 20, color: INK }), { after: 40, indent: INDENT }));
    });
  }
  return parts.join("");
}

function buildDocumentXml(ticket) {
  const body = [];

  body.push(paragraph(run("J. MUESER", { bold: true, size: 40, color: INK, caps: true, font: FONT_DISPLAY }), { after: 20 }));
  body.push(paragraph(run("Alterations Ticket", { size: 20, color: SLATE, caps: true }), { after: 160, border: { side: "bottom", sz: 12, color: BRASS } }));

  if (ticket.rush) {
    body.push(paragraph(run("★  RUSH", { bold: true, size: 28, color: OXBLOOD }), { after: 80 }));
  }

  body.push(paragraph(run("DUE", { bold: true, size: 15, color: SLATE, caps: true }), { after: 20 }));
  body.push(paragraph(run(ticket.dueDate, { bold: true, size: 26, color: INK }), { after: 160 }));

  body.push(paragraph(run(ticket.customerName, { bold: true, size: 26, color: INK, font: FONT_DISPLAY }), { after: 160 }));

  body.push(paragraph(run("TAILOR", { bold: true, size: 15, color: SLATE, caps: true }), { after: 20 }));
  body.push(paragraph(run(ticket.tailor, { bold: true, size: 26, color: INK }), { after: 160 }));

  body.push(paragraph(run("SALESPERSON", { bold: true, size: 15, color: SLATE, caps: true }), { after: 20 }));
  body.push(paragraph(run(ticket.salesperson, { bold: true, size: 26, color: INK }), { after: 200 }));

  ticket.garmentSections.forEach((section) => {
    body.push(garmentSectionXml(section));
  });

  if (ticket.balanceDisplay) {
    body.push(paragraph(run(ticket.balanceDisplay, { bold: true, size: 20, color: INK }), { before: 200, after: 40 }));
  }
  body.push(paragraph(run(`Created ${ticket.createdDisplay}`, { size: 15, color: SLATE }), { before: ticket.balanceDisplay ? 40 : 200, after: 60 }));

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${body.join("")}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

export async function buildDocxBlob(ticket) {
  if (!window.JSZip) {
    throw new Error("DOCX generator failed to load. Please refresh and try again.");
  }

  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );
  zip.folder("_rels").file(
    ".rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.folder("word").file("document.xml", buildDocumentXml(ticket));

  return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}
