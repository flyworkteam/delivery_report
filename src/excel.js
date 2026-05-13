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

function findDataSheet(workbook) {
  for (const name of workbook.SheetNames) {
    const sheet = workbook.Sheets[name];
    if (!sheet["!ref"]) continue;
    const matrix = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      raw: false,
    });
    const headerRow = matrix[0] || [];
    const flat = headerRow.map((c) => String(c).toLocaleUpperCase("tr-TR")).join(" ");
    if (flat.includes("İNCELEMDEN") && flat.includes("DURUM") && flat.includes("MAĞAZA")) {
      return name;
    }
  }
  return workbook.SheetNames[0];
}

function parseWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const sheetName = findDataSheet(workbook);
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
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
