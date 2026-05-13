const fs = require("fs");
const path = require("path");
const {
  File,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  TableBorders,
  VerticalAlignTable,
  TableLayoutType,
  BorderStyle,
  HeightRule,
  LineRuleType,
  TableAnchorType,
  RelativeHorizontalPosition,
  RelativeVerticalPosition,
} = require("docx");
const { getCell } = require("./excel");
const { formatDateTr, formatBarcode, sanitizeFileName } = require("./letter");

const ROOT = path.join(__dirname, "..");
const LOGO_PATH = path.join(ROOT, "arcon_kozmetik_logo.jpeg");
/** letter.js ile aynı logo boyutu (ImageRun transformation). */
const LOGO_WIDTH = 174;
const LOGO_HEIGHT = 110;

let logoBuf = null;
function getLogo() {
  if (!logoBuf) logoBuf = fs.readFileSync(LOGO_PATH);
  return logoBuf;
}

const BODY_TWIPS = 9026;
/** Logo sağ padding (twips); letter.js ile aynı. */
const LOGO_PADDING_RIGHT_TWIPS = 420;
const COL_PRODUCT = [2000, 5200, 1826];

const NONE = {
  top: { style: BorderStyle.NONE, size: 0, color: "auto" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "auto" },
  left: { style: BorderStyle.NONE, size: 0, color: "auto" },
  right: { style: BorderStyle.NONE, size: 0, color: "auto" },
};

const THIN = {
  style: BorderStyle.SINGLE,
  size: 1,
  color: "000000",
};

const CELL_BORDER_GRID = {
  top: THIN,
  bottom: THIN,
  left: THIN,
  right: THIN,
};

const FONT = "Times New Roman";
const SIZE_BODY = 24;
const RUN_BODY = { font: FONT, size: SIZE_BODY, color: "000000" };

/** Veri satırlarından sonra tabloda yer alan boş satır sayısı (elle yazım için). */
const TABLE_EMPTY_ROWS = 12;

/** Tablo satırı minimum yüksekliği (twips). */
const TABLE_ROW_HEIGHT = 400;
const CELL_LINE = 264;
const LINE_COMPACT = 240;
/** `ATLEAST`: minimum satır yüksekliği; uzun ürün adı iki satıra düşerse satır büyür. */
const ROW_HEIGHT_RULE = HeightRule.ATLEAST;

function metaLine(label, value) {
  return new Paragraph({
    spacing: { after: 40, line: LINE_COMPACT },
    style: "Normal",
    run: { font: FONT, color: "000000" },
    children: [
      new TextRun({ ...RUN_BODY, text: `${label}: `, bold: true }),
      new TextRun({ ...RUN_BODY, text: String(value ?? "").trim() }),
    ],
  });
}

function headerTitleLogoTable() {
  const logo = getLogo();
  return new Table({
    layout: TableLayoutType.FIXED,
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: [5200, BODY_TWIPS - 5200],
    borders: TableBorders.NONE,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 5200, type: WidthType.DXA },
            borders: NONE,
            verticalAlign: VerticalAlignTable.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { after: 60, line: LINE_COMPACT },
                children: [
                  new TextRun({
                    text: "📄 TESLİM TUTANAĞI",
                    bold: true,
                    font: FONT,
                    size: 26,
                    color: "000000",
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: { size: BODY_TWIPS - 5200, type: WidthType.DXA },
            borders: NONE,
            verticalAlign: VerticalAlignTable.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                spacing: { after: 60, line: LINE_COMPACT },
                indent: { end: LOGO_PADDING_RIGHT_TWIPS },
                children: [
                  new ImageRun({
                    data: logo,
                    transformation: { width: LOGO_WIDTH, height: LOGO_HEIGHT },
                    type: "jpg",
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function cellParagraph(text, opts = {}) {
  const { bold = false } = opts;
  const t = String(text ?? "").trim() || "\u00a0";
  return new Paragraph({
    spacing: { before: 0, after: 0, line: CELL_LINE, lineRule: LineRuleType.EXACT },
    children: [new TextRun({ ...RUN_BODY, text: t, bold })],
  });
}

function emptyCellParagraph() {
  return new Paragraph({
    spacing: { before: 0, after: 0, line: CELL_LINE, lineRule: LineRuleType.EXACT },
    children: [new TextRun({ ...RUN_BODY, text: "\u00a0" })],
  });
}

function productTable(dataRows) {
  const headerRow = new TableRow({
    tableHeader: true,
    height: { value: TABLE_ROW_HEIGHT, rule: ROW_HEIGHT_RULE },
    children: [
      new TableCell({
        width: { size: COL_PRODUCT[0], type: WidthType.DXA },
        borders: CELL_BORDER_GRID,
        verticalAlign: VerticalAlignTable.CENTER,
        children: [cellParagraph("Ürün Kodu", { bold: true })],
      }),
      new TableCell({
        width: { size: COL_PRODUCT[1], type: WidthType.DXA },
        borders: CELL_BORDER_GRID,
        verticalAlign: VerticalAlignTable.CENTER,
        children: [cellParagraph("Ürün Adı", { bold: true })],
      }),
      new TableCell({
        width: { size: COL_PRODUCT[2], type: WidthType.DXA },
        borders: CELL_BORDER_GRID,
        verticalAlign: VerticalAlignTable.CENTER,
        children: [cellParagraph("Açıklama", { bold: true })],
      }),
    ],
  });

  const bodyRows = dataRows.map(
    (row) =>
      new TableRow({
        height: { value: TABLE_ROW_HEIGHT, rule: ROW_HEIGHT_RULE },
        children: [
          new TableCell({
            width: { size: COL_PRODUCT[0], type: WidthType.DXA },
            borders: CELL_BORDER_GRID,
            verticalAlign: VerticalAlignTable.CENTER,
            children: [cellParagraph(formatBarcode(getCell(row, "BARKODU")))],
          }),
          new TableCell({
            width: { size: COL_PRODUCT[1], type: WidthType.DXA },
            borders: CELL_BORDER_GRID,
            verticalAlign: VerticalAlignTable.CENTER,
            children: [cellParagraph(getCell(row, "STOK ADI", "STOK ADI "))],
          }),
          new TableCell({
            width: { size: COL_PRODUCT[2], type: WidthType.DXA },
            borders: CELL_BORDER_GRID,
            verticalAlign: VerticalAlignTable.CENTER,
            children: [cellParagraph(durumOrAciklama(row))],
          }),
        ],
      }),
  );

  const filler = Array.from({ length: TABLE_EMPTY_ROWS }, () =>
    new TableRow({
      height: { value: TABLE_ROW_HEIGHT, rule: ROW_HEIGHT_RULE },
      children: [
        new TableCell({
          width: { size: COL_PRODUCT[0], type: WidthType.DXA },
          borders: CELL_BORDER_GRID,
          verticalAlign: VerticalAlignTable.CENTER,
          children: [emptyCellParagraph()],
        }),
        new TableCell({
          width: { size: COL_PRODUCT[1], type: WidthType.DXA },
          borders: CELL_BORDER_GRID,
          verticalAlign: VerticalAlignTable.CENTER,
          children: [emptyCellParagraph()],
        }),
        new TableCell({
          width: { size: COL_PRODUCT[2], type: WidthType.DXA },
          borders: CELL_BORDER_GRID,
          verticalAlign: VerticalAlignTable.CENTER,
          children: [emptyCellParagraph()],
        }),
      ],
    }),
  );

  return new Table({
    layout: TableLayoutType.FIXED,
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: COL_PRODUCT,
    borders: TableBorders.NONE,
    rows: [headerRow, ...bodyRows, ...filler],
  });
}

function durumOrAciklama(row) {
  const d = getCell(
    row,
    "İNCELEMDEN SONRA DURUM",
    "İNCELEMDEN SONRA DURUM ",
    "INCELEMDEN SONRA DURUM",
  );
  if (String(d).trim()) return String(d).trim();
  return String(getCell(row, "Açıklama", "Açıklama ") ?? "").trim();
}

/** İmza satırlarındaki nokta alanı (tek satır). */
const SIGNATURE_DOTS = ".".repeat(46);

/**
 * Paragraf çerçevesi Word'da satır/yaslama hatalarına yol açıyordu.
 * Kenarlıksız tablo + sayfa altına float: her satır ayrı hücrede, sola hizalı.
 */
function signaturesFloatingTable() {
  const L = AlignmentType.LEFT;
  const cell = (paragraphs) =>
    new TableCell({
      width: { size: BODY_TWIPS, type: WidthType.DXA },
      borders: NONE,
      children: paragraphs,
    });
  const row = (paragraphs) => new TableRow({ children: [cell(paragraphs)] });

  const line = (children, spacingAfter = 60) =>
    new Paragraph({
      alignment: L,
      spacing: { before: 0, after: spacingAfter, line: 276, lineRule: LineRuleType.AT_LEAST },
      children,
    });

  const sep = new Paragraph({
    alignment: L,
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
    },
    spacing: { before: 0, after: 120, line: 240 },
    children: [new TextRun({ ...RUN_BODY, text: "\u00a0" })],
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: [BODY_TWIPS],
    layout: TableLayoutType.FIXED,
    borders: TableBorders.NONE,
    alignment: L,
    float: {
      horizontalAnchor: TableAnchorType.MARGIN,
      verticalAnchor: TableAnchorType.MARGIN,
      relativeHorizontalPosition: RelativeHorizontalPosition.LEFT,
      relativeVerticalPosition: RelativeVerticalPosition.BOTTOM,
      leftFromText: 180,
      bottomFromText: 200,
    },
    rows: [
      row([sep]),
      row([
        line([
          new TextRun({ ...RUN_BODY, text: "👤 " }),
          new TextRun({ ...RUN_BODY, text: "Teslim Eden", bold: true }),
        ]),
      ]),
      row([
        line([
          new TextRun({ ...RUN_BODY, text: "Adı Soyadı: ", bold: true }),
          new TextRun({ ...RUN_BODY, text: SIGNATURE_DOTS }),
        ]),
      ]),
      row([
        line([
          new TextRun({ ...RUN_BODY, text: "İmzası: ", bold: true }),
          new TextRun({ ...RUN_BODY, text: SIGNATURE_DOTS }),
        ], 100),
      ]),
      row([line([new TextRun({ ...RUN_BODY, text: "\u00a0" })], 40)]),
      row([
        line([
          new TextRun({ ...RUN_BODY, text: "👤 " }),
          new TextRun({ ...RUN_BODY, text: "Teslim Alan", bold: true }),
        ]),
      ]),
      row([
        line([
          new TextRun({ ...RUN_BODY, text: "Adı Soyadı: ", bold: true }),
          new TextRun({ ...RUN_BODY, text: SIGNATURE_DOTS }),
        ]),
      ]),
      row([
        line(
          [
            new TextRun({ ...RUN_BODY, text: "İmzası: ", bold: true }),
            new TextRun({ ...RUN_BODY, text: SIGNATURE_DOTS }),
          ],
          0,
        ),
      ]),
    ],
  });
}

function buildDocument(rows, options = {}) {
  if (!rows?.length) throw new Error("Teslim tutanağı için en az bir satır gerekli.");
  const today = options.date || new Date();
  const first = rows[0];
  const takipFromRow = String(
    getCell(first, "Takip No", "TAKİP NO", "TAKIP NO", "Takip Numarası", "Takip Numarasi") ?? "",
  ).trim();
  const belgeNo =
    String(options.takipNo ?? "").trim() || takipFromRow || "BELGESIZ";
  const magaza = String(getCell(first, "Mağaza", "Mağaza ", "MAĞAZA") ?? "").trim();
  const tarih = formatDateTr(today);
  const koli = rows.length;

  const intro = new Paragraph({
    spacing: { after: 100, line: LINE_COMPACT },
    children: [
      new TextRun({ ...RUN_BODY, text: "Tarafımıza inceleme amacıyla ulaştırılan, aşağıda bilgileri yer alan ürün(ler), gerekli test ve kontrol süreçlerinden geçirilmiş olup, değerlendirme işlemleri tamamlandıktan sonra tarafınıza " }),
      new TextRun({ ...RUN_BODY, text: `toplam ${koli} koli`, bold: true }),
      new TextRun({
        ...RUN_BODY,
        text: " olarak teslim edilmiştir.",
      }),
    ],
  });

  const pageProps = {
    margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 },
  };

  const children = [
    headerTitleLogoTable(),
    metaLine("Tarih", tarih),
    metaLine("Belge No", belgeNo),
    metaLine("Teslim Edilen Mağaza", trUpperStore(magaza)),
    new Paragraph({ spacing: { after: 60, line: LINE_COMPACT } }),
    new Paragraph({
      spacing: { after: 60, line: LINE_COMPACT },
      children: [new TextRun({ ...RUN_BODY, text: "Sayın Yetkili," })],
    }),
    intro,
    productTable(rows),
    signaturesFloatingTable(),
  ];

  const doc = new File({
    title: belgeNo,
    creator: "Arcon Kozmetik — Teslim Tutanağı",
    sections: [
      {
        properties: { page: pageProps },
        children,
      },
    ],
  });

  const fileNameBase = sanitizeFileName(belgeNo) || "teslim-tutanagi";
  return { doc, fileNameBase };
}

function trUpperStore(s) {
  return String(s ?? "")
    .trim()
    .toLocaleUpperCase("tr-TR");
}

async function rowsToBuffer(rows, options) {
  const { doc } = buildDocument(rows, options);
  return Packer.toBuffer(doc);
}

module.exports = {
  buildDocument,
  rowsToBuffer,
};
