const path = require("path");
const fs = require("fs");
const os = require("os");
const express = require("express");
const multer = require("multer");
const archiverImport = require("archiver");
const {
  parseWorkbook,
  filterRedRows,
  getCell,
  groupRowsByTakipNo,
} = require("./excel");
const { rowToBuffer, sanitizeFileName, formatBarcode } = require("./letter");
const { rowsToBuffer } = require("./teslimTutanagi");
function createZipArchive() {
  if (typeof archiverImport === "function") {
    return archiverImport("zip", { zlib: { level: 9 } });
  }
  if (typeof archiverImport?.default === "function") {
    return archiverImport.default("zip", { zlib: { level: 9 } });
  }
  if (typeof archiverImport?.ZipArchive === "function") {
    return new archiverImport.ZipArchive({ zlib: { level: 9 } });
  }
  throw new Error("archiver modulu zip olusturacak sekilde yuklenemedi");
}

const upload = multer({
  storage: multer.diskStorage({
    destination: (_req, _file, cb) => {
      const dir = fs.mkdtempSync(path.join(os.tmpdir(), "arcon-iade-"));
      cb(null, dir);
    },
    filename: (_req, file, cb) => {
      cb(null, file.originalname || "upload.xlsx");
    },
  }),
  limits: { fileSize: 80 * 1024 * 1024 },
});

const app = express();
const PUBLIC = path.join(__dirname, "..", "public");

app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

app.use(express.static(PUBLIC));

app.get("/health", (_req, res) => {
  res.json({ ok: true });
});

app.post("/api/generate", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: "Excel dosyası gerekli (file alanı)." });
  }

  const module = String(req.body?.module ?? "iade")
    .trim()
    .toLowerCase();

  let buffer;
  try {
    buffer = fs.readFileSync(req.file.path);
  } catch (e) {
    return res.status(500).json({ error: "Dosya okunamadı." });
  } finally {
    try {
      fs.rmSync(path.dirname(req.file.path), { recursive: true, force: true });
    } catch (_) {
      /* ignore */
    }
  }

  let parsed;
  try {
    parsed = parseWorkbook(buffer);
  } catch (e) {
    return res.status(400).json({ error: "Excel açılamadı. Geçerli bir .xlsx yükleyin." });
  }

  const today = new Date();
  const archive = createZipArchive();
  archive.on("error", (err) => {
    if (!res.headersSent) res.status(500).end(String(err));
  });

  if (module === "teslim" || module === "teslim-tutanagi") {
    const groups = groupRowsByTakipNo(parsed.rows);
    if (groups.length === 0) {
      return res.status(422).json({
        error:
          "Takip No dolu satır bulunamadı. Sütun adı 'Takip No', 'TAKİP NO' veya 'Takip Numarası' olmalıdır.",
        sheetName: parsed.sheetName,
        rowCount: parsed.rows.length,
      });
    }

    res.setHeader("Content-Type", "application/zip");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="teslim-tutanaklari-${Date.now()}.zip"`,
    );
    archive.pipe(res);

    const dupCount = new Map();
    for (const { takipNo, rows } of groups) {
      const buf = await rowsToBuffer(rows, { date: today, takipNo });
      const base = sanitizeFileName(takipNo) || "teslim-tutanagi";
      const n = (dupCount.get(base) ?? 0) + 1;
      dupCount.set(base, n);
      const finalName =
        n === 1 ? `${base}.docx` : `${base} (${n - 1}).docx`;
      archive.append(buf, { name: finalName });
    }
    await archive.finalize();
    return;
  }

  const redRows = filterRedRows(parsed.rows);
  if (redRows.length === 0) {
    return res.status(422).json({
      error:
        "İNCELEMDEN SONRA DURUM sütununda 'RED' bulunamadı veya veri sayfası seçilemedi.",
      sheetName: parsed.sheetName,
      rowCount: parsed.rows.length,
    });
  }

  res.setHeader("Content-Type", "application/zip");
  res.setHeader(
    "Content-Disposition",
    `attachment; filename="iade-yazilari-${Date.now()}.zip"`,
  );
  archive.pipe(res);

  const dupCount = new Map();

  for (const row of redRows) {
    const buf = await rowToBuffer(row, { date: today });
    const barkod = formatBarcode(getCell(row, "BARKODU"));
    const stok = String(getCell(row, "STOK ADI") ?? "").trim();
    const base =
      sanitizeFileName(`${barkod || "barkod"} - ${stok || "urun"}`) || "belge";
    const n = (dupCount.get(base) ?? 0) + 1;
    dupCount.set(base, n);
    const finalName =
      n === 1 ? `${base}.docx` : `${base} (${n - 1}).docx`;
    archive.append(buf, { name: finalName });
  }

  await archive.finalize();
});

const PORT = 3033;
app.listen(PORT, () => {
  console.log(`Belge sunucusu http://localhost:${PORT} (iade + teslim tutanağı)`);
});
