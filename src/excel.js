const XLSX = require("xlsx");

/**
 * Satır sözlüğünden sütun değerini okur (başlıktaki fazla boşlukları tolere eder).
 */
function normHeader(s) {
  return String(s).trim().toLocaleLowerCase("tr-TR");
}

function getCell(row, ...names) {
  const keys = Object.keys(row);
  const aliases = [...new Set(names.map((n) => normHeader(n)))];
  for (const key of keys) {
    const k = normHeader(key);
    if (aliases.includes(k)) return row[key];
  }
  return "";
}

/**
 * Excel bazen kullanılmayan milyonlarca satırı !ref'e yazar; tüm aralığı sheet_to_json
 * ile okumak bellek/CPU'yu öldürür (nginx 502 / zaman aşımı). Sadece !ref'in ilk satırındaki
 * hücreleri okuyarak başlık ararız.
 */
function firstSheetRowCellStrings(sheet) {
  if (!sheet["!ref"]) return [];
  const d = XLSX.utils.decode_range(sheet["!ref"]);
  const r = d.s.r;
  const out = [];
  for (let c = d.s.c; c <= d.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = sheet[addr];
    let v = "";
    if (cell) {
      if (cell.w != null) v = String(cell.w);
      else if (cell.v != null) v = String(cell.v);
    }
    out.push(v);
  }
  return out;
}

function findDataSheet(workbook) {
  for (const name of workbook.SheetNames) {
    const sheet = workbook.Sheets[name];
    if (!sheet["!ref"]) continue;
    const headerRow = firstSheetRowCellStrings(sheet);
    const flat = headerRow.map((c) => String(c).toLocaleUpperCase("tr-TR")).join(" ");
    if (flat.includes("İNCELEMDEN") && flat.includes("DURUM") && flat.includes("MAĞAZA")) {
      return name;
    }
  }
  return workbook.SheetNames[0];
}

/**
 * Excel bazen P1048084 gibi tek bir hücreyle !ref'i milyon satıra uzatır; tüm anahtarların
 * max satırını almak sheet_to_json'i yine patlatır. Çok yüksekte seyrek (tekil) satırları at.
 */
function trimmedMaxDataRow(rowCounts) {
  const descending = [...rowCounts.keys()].sort((a, b) => b - a);
  const TRUST_ROW_LTE = 12_000;
  const MIN_CELLS_IF_ABOVE = 4;
  const MIN_GAP_TO_NEXT_TO_DROP_ORPHAN = 3_000;

  for (let i = 0; i < descending.length; i++) {
    const r = descending[i];
    const cnt = rowCounts.get(r) ?? 0;
    if (r <= TRUST_ROW_LTE) return r;
    if (cnt >= MIN_CELLS_IF_ABOVE) return r;
    const nextLower = descending[i + 1];
    if (nextLower === undefined) return r;
    if (r - nextLower < MIN_GAP_TO_NEXT_TO_DROP_ORPHAN) return r;
  }
  return descending[descending.length - 1] ?? 0;
}

/** Gerçek dolu hücrelere göre sheet_to_json aralığı; dev sahte !ref ve sondaki hayalet hücreyi tolere eder. */
function rangeForDataRows(sheet) {
  if (!sheet["!ref"]) return undefined;
  const decoded = XLSX.utils.decode_range(sheet["!ref"]);
  const rowCounts = new Map();
  let minR = decoded.s.r;
  let minC = decoded.s.c;
  let maxC = decoded.s.c;
  let any = false;
  for (const k of Object.keys(sheet)) {
    if (k[0] === "!") continue;
    any = true;
    try {
      const { r, c } = XLSX.utils.decode_cell(k);
      rowCounts.set(r, (rowCounts.get(r) || 0) + 1);
      if (r < minR) minR = r;
      if (c < minC) minC = c;
      if (c > maxC) maxC = c;
    } catch (_) {
      /* ignore */
    }
  }
  if (!any) return decoded;
  const maxR = Math.min(trimmedMaxDataRow(rowCounts), decoded.e.r);
  return { s: { r: minR, c: minC }, e: { r: maxR, c: maxC } };
}

function parseWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const sheetName = findDataSheet(workbook);
  const sheet = workbook.Sheets[sheetName];
  const range = rangeForDataRows(sheet);
  const rows = XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    raw: false,
    range,
  });
  return { sheetName, rows };
}

function isRedStatus(value) {
  const s = String(value ?? "")
    .trim()
    .toLowerCase();
  return s === "red" || s.includes("red");
}

function filterRedRows(rows) {
  return rows.filter((row) => {
    const durum = getCell(
      row,
      "İNCELEMDEN SONRA DURUM",
      "İNCELEMDEN SONRA DURUM ",
      "INCELEMDEN SONRA DURUM",
    );
    return isRedStatus(durum);
  });
}

function getTakipNo(row) {
  return String(
    getCell(
      row,
      "Takip No",
      "TAKİP NO",
      "TAKIP NO",
      "Takip Numarası",
      "Takip Numarasi",
      "Takip numarası",
    ) ?? "",
  ).trim();
}

/** Aynı takip numarasına sahip satırları gruplar (Teslim Tutanağı: tek belge, çok ürün). */
function groupRowsByTakipNo(rows) {
  const map = new Map();
  for (const row of rows) {
    const t = getTakipNo(row);
    if (!t) continue;
    if (!map.has(t)) map.set(t, []);
    map.get(t).push(row);
  }
  return [...map.entries()].map(([takipNo, groupRows]) => ({ takipNo, rows: groupRows }));
}

module.exports = {
  getCell,
  findDataSheet,
  parseWorkbook,
  filterRedRows,
  getTakipNo,
  groupRowsByTakipNo,
};
