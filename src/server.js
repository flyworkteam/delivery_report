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

function logLine(level, message, meta) {
  const suffix =
    meta !== undefined
      ? ` ${typeof meta === "object" ? JSON.stringify(meta) : String(meta)}`
      : "";
  console.log(`[${new Date().toISOString()}] [${level}] ${message}${suffix}`);
}
const logger = {
  info: (message, meta) => logLine("INFO", message, meta),
  warn: (message, meta) => logLine("WARN", message, meta),
  error: (message, meta) => logLine("ERROR", message, meta),
};

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

app.use((req, res, next) => {
  const t0 = Date.now();
  res.on("finish", () => {
    logger.info("http", {
      method: req.method,
      path: req.originalUrl || req.url,
      status: res.statusCode,
      ms: Date.now() - t0,
    });
  });
  next();
});

app.use(express.static(PUBLIC));

app.get("/health", (_req, res) => {
  res.json({ ok: true });
});

app.post("/api/generate", upload.single("file"), async (req, res) => {
  const tStart = Date.now();
  if (!req.file) {
    logger.warn("generate: excel yok (file alanı boş)");
    return res.status(400).json({ error: "Excel dosyası gerekli (file alanı)." });
  }

  const module = String(req.body?.module ?? "iade")
    .trim()
    .toLowerCase();

  let buffer;
  try {
    buffer = fs.readFileSync(req.file.path);
    logger.info("generate: dosya okundu", {
      module,
      originalName: req.file.originalname,
      bytes: buffer.length,
    });
  } catch (e) {
    logger.error("generate: dosya okunamadı", { message: e?.message });
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
    logger.info("generate: excel açıldı", {
      module,
      sheetName: parsed.sheetName,
      rowCount: parsed.rows.length,
    });
  } catch (e) {
    logger.warn("generate: excel parse hatası", { message: e?.message });
    return res.status(400).json({ error: "Excel açılamadı. Geçerli bir .xlsx yükleyin." });
  }

  const today = new Date();
  const archive = createZipArchive();
  archive.on("error", (err) => {
    logger.error("generate: zip archiver hatası", { message: err?.message });
    if (!res.headersSent) res.status(500).end(String(err));
  });

  if (module === "teslim" || module === "teslim-tutanagi") {
    const groups = groupRowsByTakipNo(parsed.rows);
    if (groups.length === 0) {
      logger.warn("generate: teslim — takip no grubu yok", {
        sheetName: parsed.sheetName,
        rowCount: parsed.rows.length,
      });
      return res.status(422).json({
        error:
          "Takip No dolu satır bulunamadı. Sütun adı 'Takip No', 'TAKİP NO' veya 'Takip Numarası' olmalıdır.",
        sheetName: parsed.sheetName,
        rowCount: parsed.rows.length,
      });
    }

    logger.info("generate: teslim zip üretiliyor", { docCount: groups.length });

    res.setHeader("Content-Type", "application/zip");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="teslim-tutanaklari-${Date.now()}.zip"`,
    );
    archive.pipe(res);

    const dupCount = new Map();
    for (const { takipNo, rows } of groups) {
      const buf = await rowsToBuffer(rows, { date: today, takipNo });
      const first = rows[0];
      const magaza = String(
        getCell(first, "Mağaza", "Mağaza ", "MAĞAZA") ?? "",
      ).trim();
      const takipPart = sanitizeFileName(takipNo) || "teslim-tutanagi";
      const magazaPart = magaza ? sanitizeFileName(magaza) : "";
      const base =
        magazaPart && magazaPart.length > 0
          ? `${takipPart} - ${magazaPart}`
          : takipPart;
      const n = (dupCount.get(base) ?? 0) + 1;
      dupCount.set(base, n);
      const finalName =
        n === 1 ? `${base}.docx` : `${base} (${n - 1}).docx`;
      archive.append(buf, { name: finalName });
    }
    await archive.finalize();
    logger.info("generate: teslim tamam", {
      docCount: groups.length,
      ms: Date.now() - tStart,
    });
    return;
  }

  const redRows = filterRedRows(parsed.rows);
  if (redRows.length === 0) {
    logger.warn("generate: iade — RED satır yok", {
      sheetName: parsed.sheetName,
      rowCount: parsed.rows.length,
    });
    return res.status(422).json({
      error:
        "İNCELEMDEN SONRA DURUM sütununda 'RED' bulunamadı veya veri sayfası seçilemedi.",
      sheetName: parsed.sheetName,
      rowCount: parsed.rows.length,
    });
  }

  logger.info("generate: iade zip üretiliyor", { docCount: redRows.length });

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
    const magaza = String(
      getCell(row, "Mağaza", "Mağaza ", "MAĞAZA") ?? "",
    ).trim();
    const nameParts = [barkod || "barkod", stok || "urun"];
    if (magaza) nameParts.push(magaza);
    const base = sanitizeFileName(nameParts.join(" - ")) || "belge";
    const n = (dupCount.get(base) ?? 0) + 1;
    dupCount.set(base, n);
    const finalName =
      n === 1 ? `${base}.docx` : `${base} (${n - 1}).docx`;
    archive.append(buf, { name: finalName });
  }

  await archive.finalize();
  logger.info("generate: iade tamam", {
    docCount: redRows.length,
    ms: Date.now() - tStart,
  });
});

app.use((err, req, res, next) => {
  logger.error("işlenmeyen hata", {
    path: req.originalUrl || req.url,
    message: err?.message,
    stack: err?.stack,
  });
  if (res.headersSent) return next(err);
  res.status(500).json({ error: "Sunucu hatası." });
});

const PORT = 3033;
app.listen(PORT, () => {
  logger.info(`Belge sunucusu dinliyor`, { port: PORT });
});
