/**
 * Tek seferlik CLI test: örnek Excel ile RED satırlarını sayar ve ilk belgeyi yazar.
 */
const fs = require("fs");
const path = require("path");
const { parseWorkbook, filterRedRows } = require("./excel");
const { rowToBuffer } = require("./letter");

const xlsxPath =
  process.argv[2] ||
  path.join(__dirname, "..", "Müşteri İadeleri 02.06.25 Ten sonra.xlsx");

async function main() {
  const buf = fs.readFileSync(xlsxPath);
  const { sheetName, rows } = parseWorkbook(buf);
  const reds = filterRedRows(rows);
  console.log("Sayfa:", sheetName);
  console.log("Toplam satır:", rows.length);
  console.log("RED satır:", reds.length);
  if (!reds.length) return;
  const out = path.join(__dirname, "..", "_test-first-red.docx");
  const firstBuf = await rowToBuffer(reds[0], { date: new Date() });
  fs.writeFileSync(out, firstBuf);
  console.log("Örnek çıktı:", out, "(" + firstBuf.length + " byte)");
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
