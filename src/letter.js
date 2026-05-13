const fs = require("fs");
const path = require("path");
const {
  File,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  AlignmentType,
  Footer,
  Table,
  TableRow,
  TableCell,
  WidthType,
  TableBorders,
  VerticalAlignTable,
  TableLayoutType,
  BorderStyle,
} = require("docx");
const { getCell } = require("./excel");

const ROOT = path.join(__dirname, "..");

const LOGO_PATH = path.join(ROOT, "arcon_kozmetik_logo.jpeg");
/** Sağ kenara yakın logonun sağında nefes (twips). */
const LOGO_PADDING_RIGHT_TWIPS = 420;
let cachedBuffers = null;
function getMediaBuffers() {
  if (!cachedBuffers) {
    cachedBuffers = {
      logo: fs.readFileSync(LOGO_PATH),
    };
  }
  return cachedBuffers;
}

function formatDateTr(date) {
  const d = date instanceof Date ? date : new Date();
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  return `${day}.${month}.${year}`;
}

function trUpper(s) {
  return String(s ?? "")
    .trim()
    .toLocaleUpperCase("tr-TR");
}

function formatBarcode(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value.toLocaleString("en-US", { useGrouping: false, maximumFractionDigits: 0 });
  }
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  if (/e\+/i.test(raw)) {
    const n = Number(raw);
    if (Number.isFinite(n)) {
      return n.toLocaleString("en-US", { useGrouping: false, maximumFractionDigits: 0 });
    }
  }
  return raw;
}

function isFragranceCase(stockName, aciklama) {
  const combined = `${stockName} ${aciklama}`.toLowerCase();
  if (
    /(koku|kalıc|kalic|vapo|vapı|vapor|parfüm|parfum|edp|edt|edc|spray|şişe|kolonya|frag)/i.test(
      aciklama,
    )
  ) {
    return true;
  }
  if (/\b(edp|edt|edc|parfüm|parfum|spray|kolonya)\b/i.test(combined)) return true;
  if (/\d+\s*ml\b/i.test(stockName)) return true;
  return false;
}

/** Standart sayfa gövdesi genişliği (twips); tablo sütunları toplamı buna yakın olmalı. */
const BODY_WIDTH_TWIPS = 9026;

const CONTACT_ADDRESS_LINE1 = "Ahi Evran Cad. 42 Maslak No:6 A Kule Kat:10 D:1";
const CONTACT_ADDRESS_LINE2 = "34398, Maslak / İstanbul +90 212 290 28 50(pbx)";
const CONTACT_WEBSITE = "arconkozmetik.com.tr";
const CONTACT_FONT = "Arial";
const CONTACT_RUN = { font: CONTACT_FONT, color: "000000", size: 20 };

const CLOSING = [
  "Arcon Kozmetik iade işlemlerinde Tüketiciyi Koruma Kanunu çerçevesinde hareket eder. Üründen kaynaklanan bir problem olmaması nedeni ile bu ürünü iade alamayacağımızı bildiririz.",
  "",
  "Bilginizi rica eder ve iyi çalışmalar dileriz.",
  "",
  "Saygılarımızla,",
  "",
  "",
  "ARCON KOZMETİK",
];

function buildMainParagraphs(musteri, stokAdi, aciklama, fragrance) {
  const cust = trUpper(musteri);
  const product = String(stokAdi ?? "").trim();

  if (fragrance) {
    const p1Parts = [
      "Müşteriniz SN. ",
      cust,
      " adına test talep edilen ",
      product,
      " ürün ile ilgili olarak, gerek cilt gerek koku kartında yapılan test sonucunda ürünün yapısında, kokusunda, kalıcılığında; vaporizatör, şişe ve ambalajında herhangi bir soruna rastlanmamıştır.",
    ];
    const p2 =
      "Bir parfümün cilde tatbik edilmesinden sonra kokunun kalıcılığı ortalama 3-4 saat kadardır.";
    const p3 =
      "Parfümünüzün kalıcılığını arttırmak için ürünlerin vücut kremlerini veya kokusuz cilt nemlendirici kullanarak süreyi uzatabilirsiniz.";
    return [p1Parts, "", p2, "", p3];
  }

  const p1Parts = [
    "Müşteriniz SN. ",
    cust,
    " adına test talep edilen ",
    product,
    " ilgili olarak, yapılan test sonucunda ürünün yapısında, dokusunda ve ambalajında herhangi bir soruna rastlanmamıştır.",
  ];
  return [p1Parts];
}

function textParagraph(text, opts = {}) {
  const { bold = false, alignment = AlignmentType.LEFT } = opts;
  const runs = [];
  runs.push(
    new TextRun({
      text,
      bold,
      font: "Times New Roman",
      size: 24,
      color: "000000",
    }),
  );
  return new Paragraph({
    alignment,
    spacing: { after: 120, line: 276 },
    style: "Normal",
    run: {
      font: "Times New Roman",
      color: "000000",
    },
    children: runs,
  });
}

function textPartsParagraph(parts, opts = {}) {
  const { alignment = AlignmentType.LEFT } = opts;
  const runs = parts.map((part) =>
    new TextRun({
      text: part.text,
      bold: Boolean(part.bold),
      font: "Times New Roman",
      size: 24,
      color: "000000",
    }),
  );
  return new Paragraph({
    alignment,
    spacing: { after: 120, line: 276 },
    style: "Normal",
    run: {
      font: "Times New Roman",
      color: "000000",
    },
    children: runs,
  });
}

/** Solda iki satır adres, sağda dikey ortada kalın web sitesi; kenarlıksız tablo (çakışma yok). */
function contactFooterTable() {
  const colLeft = Math.round(BODY_WIDTH_TWIPS * 0.72);
  const colRight = BODY_WIDTH_TWIPS - colLeft;
  const cellPara = (spacingAfter) => ({
    spacing: { before: 0, after: spacingAfter, line: 276 },
    style: "Normal",
  });

  return new Table({
    layout: TableLayoutType.FIXED,
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: [colLeft, colRight],
    borders: TableBorders.NONE,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: colLeft, type: WidthType.DXA },
            verticalAlign: VerticalAlignTable.CENTER,
            margins: { top: 80, bottom: 80, right: 120 },
            borders: {
              top: { style: BorderStyle.NONE, size: 0, color: "auto" },
              bottom: { style: BorderStyle.NONE, size: 0, color: "auto" },
              left: { style: BorderStyle.NONE, size: 0, color: "auto" },
              right: { style: BorderStyle.NONE, size: 0, color: "auto" },
            },
            children: [
              new Paragraph({
                ...cellPara(40),
                children: [new TextRun({ ...CONTACT_RUN, text: CONTACT_ADDRESS_LINE1 })],
              }),
              new Paragraph({
                ...cellPara(0),
                children: [new TextRun({ ...CONTACT_RUN, text: CONTACT_ADDRESS_LINE2 })],
              }),
            ],
          }),
          new TableCell({
            width: { size: colRight, type: WidthType.DXA },
            verticalAlign: VerticalAlignTable.CENTER,
            margins: { top: 80, bottom: 80, left: 120 },
            borders: {
              top: { style: BorderStyle.NONE, size: 0, color: "auto" },
              bottom: { style: BorderStyle.NONE, size: 0, color: "auto" },
              left: { style: BorderStyle.NONE, size: 0, color: "auto" },
              right: { style: BorderStyle.NONE, size: 0, color: "auto" },
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                ...cellPara(0),
                children: [new TextRun({ ...CONTACT_RUN, text: CONTACT_WEBSITE, bold: true })],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function buildDocument(row, options = {}) {
  const today = options.date || new Date();
  const media = getMediaBuffers();

  const barkod = formatBarcode(getCell(row, "BARKODU"));
  const stokAdi = String(getCell(row, "STOK ADI") ?? "").trim();
  const magaza = String(getCell(row, "Mağaza", "Mağaza ") ?? "").trim();
  const musteri = String(getCell(row, "Müşteri", "Müşteri ") ?? "").trim();
  const aciklama = String(getCell(row, "Açıklama", "Açıklama ") ?? "").trim();

  const title = `${barkod} - ${stokAdi}`;
  const fragrance = isFragranceCase(stokAdi, aciklama);
  const mainBlocks = buildMainParagraphs(musteri, stokAdi, aciklama, fragrance);
  const tarih = formatDateTr(today);

  const mainParagraphs = [];
  for (const block of mainBlocks) {
    if (block === "") mainParagraphs.push(new Paragraph({ spacing: { after: 120, line: 276 } }));
    else if (Array.isArray(block)) {
      mainParagraphs.push(
        textPartsParagraph([
          { text: block[0] },
          { text: block[1], bold: true },
          { text: block[2] },
          { text: block[3], bold: true },
          { text: block[4] },
        ]),
      );
    } else mainParagraphs.push(textParagraph(block));
  }

  const logoParagraph = new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 200, line: 276 },
    indent: { end: LOGO_PADDING_RIGHT_TWIPS },
    style: "Normal",
    run: {
      color: "000000",
    },
    children: [
      new ImageRun({
        data: media.logo,
        transformation: { width: 160, height: 102 },
        type: "jpg",
      }),
    ],
  });

  const pageProps = {
    margin: {
      top: 1440,
      right: 1440,
      bottom: 1440,
      left: 1440,
      /** Alt bilgi alanı (twips); iki satır + tablo için biraz geniş. */
      footer: 1020,
    },
  };

  const defaultFooter = new Footer({
    children: [contactFooterTable()],
  });

  const bodyChildren = [
    logoParagraph,
    textParagraph(tarih, { bold: true }),
    textParagraph(trUpper(magaza), { bold: true }),
    textParagraph("Sayın İlgili,"),
    new Paragraph({ spacing: { after: 120, line: 276 } }),
    ...mainParagraphs,
    new Paragraph({ spacing: { after: 120, line: 276 } }),
    ...CLOSING.map((line) => textParagraph(line)),
  ];

  const doc = new File({
    title,
    creator: "Arcon Kozmetik — İade Yazısı",
    sections: [
      {
        properties: {
          page: pageProps,
        },
        footers: {
          default: defaultFooter,
        },
        children: bodyChildren,
      },
    ],
  });

  return { doc, fileNameBase: title };
}

async function rowToBuffer(row, options) {
  const { doc } = buildDocument(row, options);
  return Packer.toBuffer(doc);
}

function sanitizeFileName(name) {
  return String(name)
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 200);
}

module.exports = {
  buildDocument,
  rowToBuffer,
  sanitizeFileName,
  formatDateTr,
  formatBarcode,
};
