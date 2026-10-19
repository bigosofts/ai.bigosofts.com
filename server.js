require("dotenv").config({ override: true });

// --- CPanel / CloudLinux NPROC Thread Limit Fix ---
// Native C++ modules like sharp (libvips) and canvas (cairo/glib) will try to
// spawn threads equal to the CPU core count on startup. On shared hosting (cPanel),
// this hits the NPROC limit and causes "glib: Error creating thread: Resource temporarily unavailable".
// We MUST set these before requiring any other module.
process.env.UV_THREADPOOL_SIZE = "1";
process.env.VIPS_CONCURRENCY = "1";
process.env.OMP_NUM_THREADS = "1";
process.env.OPENBLAS_NUM_THREADS = "1";

const express = require("express");
const cors = require("cors");
const { v4: uuidv4 } = require("uuid");
const fs = require("fs");
const path = require("path");
const { execFile, execFileSync } = require("child_process");
const multer = require("multer");
const sharp = require("sharp");
sharp.cache(false);
sharp.concurrency(1);
const {
  Document,
  Packer,
  Paragraph: OriginalParagraph,
  TextRun: OriginalTextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
  HeadingLevel,
  TabStopType,
  TabStopPosition,
  Header,
  Footer,
  PageNumber,
  ImageRun,
  ShadingType,
  TableLayoutType,
} = require("docx");
const docx = require("docx");

class SafeParagraph extends OriginalParagraph {
  constructor(options) {
    if (options && Array.isArray(options.children)) {
      options.children = options.children.flat(Infinity);
    }
    super(options);
  }
}

class SafeTextRun extends OriginalTextRun {
  constructor(options) {
    const defaultFont = {
      ascii: "Arial",
      cs: "Kalpurush",
      eastAsia: "Kalpurush",
      hAnsi: "Arial",
      hint: "default"
    };
    if (options) {
      if (options.font) {
        if (typeof options.font === "string") {
          options.font = {
            ascii: options.font,
            cs: "Kalpurush",
            eastAsia: "Kalpurush",
            hAnsi: options.font,
            hint: "default"
          };
        } else {
          options.font = { ...defaultFont, ...options.font };
        }
      } else {
        options.font = defaultFont;
      }
    }
    super(options);
  }
}

const Paragraph = SafeParagraph;
const TextRun = SafeTextRun;
docx.Paragraph = SafeParagraph;
docx.TextRun = SafeTextRun;


const app = express();
app.use(cors());
app.use(express.json({ limit: "50mb" }));

// Serve static files with no-cache for CSS to ensure updates are always fetched
app.use(
  express.static("public", {
    maxAge: 0,
    etag: false,
  }),
);

// Ensure data directory and authorized users file
const DATA_DIR = path.join(__dirname, "data");
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR);
const AUTH_USERS_FILE = path.join(DATA_DIR, "authorized_users.json");
if (!fs.existsSync(AUTH_USERS_FILE))
  fs.writeFileSync(AUTH_USERS_FILE, "[]", "utf8");

// In-memory OTP store: { email: { otp, name, age, page, expiresAt } }
const otpStore = new Map();
const OTP_TTL_MS = 10 * 60 * 1000; // 10 minutes

const PORT = process.env.PORT || 3000;
const OUTPUT_DIR = path.join(__dirname, "output");
const UPLOADS_DIR = path.join(__dirname, "uploads");
const VALID_PALETTES = new Set([
  "forestMoss",
  "oceanBreeze",
  "coralEnergy",
  "purpleMajesty",
  "cherryBlossom",
  "sunsetGlow",
  "arcticBreeze",
  "lavenderDream",
  "earthyTerracotta",
  "goldenHour",
  "midnightGold",
  "darkCharcoal",
  "cyberNeon",
  "emeraldAbyss",
  "crimsonForge",
]);
const GENERATED_DOCUMENT_TTL_MS = 30 * 60 * 1000;

// Ensure directories exist
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR);
if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR);

// Serve uploaded images as static files
app.use("/uploads", express.static(UPLOADS_DIR));

// ----------------------------------------------------------------------------
// Multer Configuration for Image Uploads
// ----------------------------------------------------------------------------
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOADS_DIR),
  filename: (req, file, cb) => {
    // Use sessionId (from header or generate one) + UUID to namespace uploads per user session
    const sessionId = (req.headers["x-session-id"] || "anon")
      .replace(/[^a-zA-Z0-9-]/g, "")
      .slice(0, 36);
    const ext = path.extname(file.originalname);
    const uniqueName = `img_${sessionId}_${uuidv4()}${ext}`;
    cb(null, uniqueName);
  },
});

const upload = multer({
  storage,
  limits: { fileSize: 20 * 1024 * 1024 }, // 20MB limit (increased for PDFs)
  fileFilter: (req, file, cb) => {
    const allowed = /jpeg|jpg|png|gif|svg|webp|bmp|pdf/;
    const extOk = allowed.test(path.extname(file.originalname).toLowerCase());
    const mimeOk =
      allowed.test(file.mimetype.split("/")[1]) ||
      file.mimetype === "application/pdf";
    if (extOk || mimeOk) {
      cb(null, true);
    } else {
      cb(new Error("Only image and PDF files are allowed"));
    }
  },
});

// Periodic cleanup: remove stale uploads and completed presentations after 30 minutes.
const GENERATED_PRESENTATION_TTL_MS = 30 * 60 * 1000;

setInterval(
  () => {
    try {
      const now = Date.now();
      const files = fs.readdirSync(UPLOADS_DIR);
      files.forEach((f) => {
        try {
          const fp = path.join(UPLOADS_DIR, f);
          const stat = fs.statSync(fp);
          if (now - stat.mtimeMs > 30 * 60 * 1000) {
            fs.unlinkSync(fp);
            console.log(`ЁЯз╣ Stale upload removed: ${f}`);
          }
        } catch (e) {
          console.error(`Error deleting upload file ${f}:`, e.message);
        }
      });

      fs.readdirSync(OUTPUT_DIR).forEach((f) => {
        try {
          if (
            !/^(Presentation_[a-f0-9-]+\.(pptx|debug\.js|plan\.json)|Document_[a-f0-9-]+\.(docx|repro\.docx|debug\.js|plan\.json|raw_plan\.txt))$/i.test(
              f,
            )
          )
            return;
          const fp = path.join(OUTPUT_DIR, f);
          const stat = fs.statSync(fp);
          if (stat.isFile() && now - stat.mtimeMs > GENERATED_DOCUMENT_TTL_MS) {
            fs.unlinkSync(fp);
            console.log(`Expired presentation/document removed: ${f}`);
          }
        } catch (e) {
          console.error(`Error deleting output file ${f}:`, e.message);
        }
      });
    } catch (e) {
      /* silent */
    }
  },
  10 * 60 * 1000,
).unref(); // run every 10 minutes; don't retain imported test processes

// ----------------------------------------------------------------------------
// API: Upload Image
// ----------------------------------------------------------------------------
app.post("/api/upload-image", upload.single("image"), (req, res) => {
  try {
    if (!req.file) {
      return res
        .status(400)
        .json({ success: false, error: "No image file provided" });
    }

    const filePath = req.file.path;
    const staticUrl = `/uploads/${req.file.filename}`;

    res.json({
      success: true,
      filePath: filePath,
      staticUrl: staticUrl,
      fileName: req.file.filename,
    });
  } catch (error) {
    console.error("Upload error:", error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ----------------------------------------------------------------------------
// API: Request access (send OTP + admin notification)
// ----------------------------------------------------------------------------
app.post("/api/request-access", async (req, res) => {
  try {
    const { name, age, email, page } = req.body || {};
    if (!name || !email || !page) {
      return res.status(400).json({ success: false, error: "Missing fields" });
    }

    // Generate 6-digit OTP
    const otp = String(Math.floor(100000 + Math.random() * 900000));
    const expiresAt = Date.now() + OTP_TTL_MS;
    otpStore.set(email + "|" + page, { otp, name, age, page, expiresAt });

    // Send admin email
    try {
      const adminHtml = `<p>User trying to access ${page}:</p><ul><li><strong>Name:</strong> ${name}</li><li><strong>Email:</strong> ${email}</li><li><strong>Age:</strong> ${age || "-"}</li></ul>`;
      await fetch("https://mailer.bigosofts.com/send-mail", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          to: "bigosofts@gmail.com",
          subject: `[BigOSofts] Access attempt: ${page} — from ${name}`,
          text: `User ${name} <${email}> (age: ${age || "-"}) is attempting to access ${page}.`,
          html: adminHtml,
        }),
      });
    } catch (adminErr) {
      console.error("Admin notification email failed to send:", adminErr.message);
    }

    // Always log the OTP to the server console for local testing and debugging
    console.log(`\n🔑 [OTP FALLBACK] Generated OTP for ${email} on ${page}: ${otp}\n`);

    // Send client email with OTP
    let clientMailSent = false;
    try {
      const clientHtml = `<p>Hi ${name},</p><p>Your one-time code to access BigOSofts ${page} is:</p><h2 style="letter-spacing:4px">${otp}</h2><p>This code expires in 10 minutes. If you didn't request this, ignore this email.</p>`;
      await fetch("https://mailer.bigosofts.com/send-mail", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          to: email,
          subject: "Your BigOSofts access code",
          text: `Hi ${name}, your access code is ${otp}`,
          html: clientHtml,
        }),
      });
      clientMailSent = true;
    } catch (clientErr) {
      console.error("Client OTP email failed to send:", clientErr.message);
    }

    if (!clientMailSent) {
      return res.json({
        success: true,
        message: "OTP generated (Check server console: mailer service connection failed)",
        devFallback: true
      });
    }

    return res.json({ success: true, message: "OTP sent" });
  } catch (err) {
    console.error("request-access error:", err);
    return res.status(500).json({ success: false, error: err.message });
  }
});

// ----------------------------------------------------------------------------
// API: Verify OTP and grant access
// ----------------------------------------------------------------------------
app.post("/api/verify-otp", async (req, res) => {
  try {
    const { email, otp, page } = req.body || {};
    if (!email || !otp || !page) {
      return res.status(400).json({ success: false, error: "Missing fields" });
    }
    const key = email + "|" + page;
    const record = otpStore.get(key);
    if (!record) {
      return res.status(400).json({ success: false, error: "No OTP found" });
    }
    if (Date.now() > record.expiresAt) {
      otpStore.delete(key);
      return res.status(400).json({ success: false, error: "OTP expired" });
    }
    if (record.otp !== String(otp).trim()) {
      return res.status(400).json({ success: false, error: "Invalid OTP" });
    }

    // Append to authorized users file
    const users = JSON.parse(fs.readFileSync(AUTH_USERS_FILE, "utf8") || "[]");
    users.push({
      name: record.name,
      email,
      age: record.age || null,
      page,
      grantedAt: new Date().toISOString(),
    });
    fs.writeFileSync(AUTH_USERS_FILE, JSON.stringify(users, null, 2), "utf8");

    // cleanup
    otpStore.delete(key);
    return res.json({ success: true });
  } catch (err) {
    console.error("verify-otp error:", err);
    return res.status(500).json({ success: false, error: err.message });
  }
});

// ----------------------------------------------------------------------------
// API: Upload File (Image or PDF)
// ----------------------------------------------------------------------------
app.post("/api/upload-file", upload.single("file"), (req, res) => {
  try {
    if (!req.file) {
      return res
        .status(400)
        .json({ success: false, error: "No file provided" });
    }
    const filePath = req.file.path;
    const staticUrl = `/uploads/${req.file.filename}`;
    res.json({
      success: true,
      filePath: filePath,
      staticUrl: staticUrl,
      fileName: req.file.filename,
      mimeType: req.file.mimetype,
    });
  } catch (error) {
    console.error("Upload error:", error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ----------------------------------------------------------------------------
// PDF & Document Data Extraction Engine (Hybrid: Native Text + High-Res Vision)
// ----------------------------------------------------------------------------

/**
 * Extracts native digital selectable text from a PDF file using pdfjs-dist.
 * Returns { isNativeText: boolean, numPages: number, pages: Array<{ pageNum: number, text: string }>, totalChars: number }
 */
async function extractPdfText(filePath, maxPages = 50) {
  try {
    const { getDocument } = await import("pdfjs-dist/legacy/build/pdf.mjs");
    const standardFontPath = path
      .join(__dirname, "node_modules", "pdfjs-dist", "standard_fonts", "/")
      .replace(/\\/g, "/");

    const data = new Uint8Array(fs.readFileSync(filePath));
    const loadingTask = getDocument({
      data,
      standardFontDataUrl: standardFontPath,
      disableFontFace: true,
      useSystemFonts: true,
    });
    const pdfDoc = await loadingTask.promise;
    const numPages = Math.min(pdfDoc.numPages, maxPages);

    let totalChars = 0;
    const pages = [];

    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
      const page = await pdfDoc.getPage(pageNum);
      const textContent = await page.getTextContent();

      let lastY = null;
      let pageText = "";
      for (const item of textContent.items) {
        if (!item.str) continue;
        if (lastY !== null && Math.abs(item.transform[5] - lastY) > 5) {
          pageText += "\n";
        } else if (pageText && !pageText.endsWith(" ") && !pageText.endsWith("\n")) {
          pageText += " ";
        }
        pageText += item.str;
        lastY = item.transform[5];
      }
      pageText = pageText.trim();
      totalChars += pageText.length;
      pages.push({ pageNum, text: pageText });
    }

    const averageChars = pages.length > 0 ? totalChars / pages.length : 0;
    // Native digital text PDF if average page has more than 40 characters
    const isNativeText = averageChars > 40;

    return {
      isNativeText,
      numPages: pdfDoc.numPages,
      pages,
      totalChars,
    };
  } catch (err) {
    console.warn("⚠️ Native PDF text extraction warning:", err.message);
    return { isNativeText: false, error: err.message, pages: [] };
  }
}

/**
 * Prepares extraction units (native text pages or high-res images) from uploaded files.
 */
async function prepareExtractionUnits(filePaths) {
  const units = [];

  for (const filePath of filePaths) {
    if (!fs.existsSync(filePath)) continue;

    const ext = path.extname(filePath).toLowerCase();
    const fileName = path.basename(filePath);

    if (ext === ".pdf") {
      // 1. First attempt instant native digital text extraction
      const textResult = await extractPdfText(filePath, 50);

      if (textResult.isNativeText && textResult.pages.length > 0) {
        console.log(`📄 [Fast-Path] Found digital text in ${fileName} (${textResult.pages.length} pages, ${textResult.totalChars} chars). Skipping heavy image rasterization.`);
        for (const p of textResult.pages) {
          units.push({
            type: "text",
            fileName,
            pageNum: p.pageNum,
            text: p.text,
          });
        }
        continue;
      }

      // 2. Fallback: Scanned / Image-only PDF -> Render high-res pages for Vision OCR
      console.log(`📷 [Vision OCR] Scanned PDF detected in ${fileName}. Rendering high-res pages for AI vision...`);
      const { pdf } = await import("pdf-to-img");
      const standardFontPath = path
        .join(__dirname, "node_modules", "pdfjs-dist", "standard_fonts", "/")
        .replace(/\\/g, "/");

      const document = await pdf(filePath, {
        scale: 1.5, // 150-200 DPI equivalent for clean OCR
        standardFontDataUrl: standardFontPath,
      });

      let pageCount = 0;
      for await (const pageBuffer of document) {
        pageCount++;
        if (pageCount > 50) break;

        try {
          // Optimized for AI text extraction: 1600px max, quality 80%, with sharpening for OCR
          const enhanced = await sharp(pageBuffer)
            .resize({
              width: 1600,
              height: 1600,
              fit: "inside",
              withoutEnlargement: true,
            })
            .normalize()
            .sharpen({ sigma: 0.8 })
            .jpeg({ quality: 80 })
            .toBuffer();

          units.push({
            type: "image",
            fileName,
            pageNum: pageCount,
            base64: enhanced.toString("base64"),
          });
        } catch (sharpErr) {
          console.warn(`Sharp processing failed on page ${pageCount}, using raw buffer:`, sharpErr.message);
          units.push({
            type: "image",
            fileName,
            pageNum: pageCount,
            base64: pageBuffer.toString("base64"),
          });
        }
      }
    } else {
      // Image file (PNG, JPG, WEBP, etc.)
      try {
        const enhanced = await sharp(filePath)
          .resize({
            width: 1600,
            height: 1600,
            fit: "inside",
            withoutEnlargement: true,
          })
          .normalize()
          .sharpen({ sigma: 0.8 })
          .jpeg({ quality: 80 })
          .toBuffer();

        units.push({
          type: "image",
          fileName,
          pageNum: 1,
          base64: enhanced.toString("base64"),
        });
      } catch (sharpErr) {
        console.warn(`Sharp processing failed on image, using raw buffer:`, sharpErr.message);
        const rawBuffer = fs.readFileSync(filePath);
        units.push({
          type: "image",
          fileName,
          pageNum: 1,
          base64: rawBuffer.toString("base64"),
        });
      }
    }
  }

  return units;
}

/**
 * Executes AI data extraction for a batch of units (text or images)
 * with seamless Markdown formatting and syntax verification.
 */
async function executeBatchExtraction(units, batchIndex, totalBatches, userInstruction = "", previousTail = "") {
  const textUnits = units.filter((u) => u.type === "text");
  const imageUnits = units.filter((u) => u.type === "image");
  const base64Images = imageUnits.map((u) => u.base64);

  let userContent = "";

  if (userInstruction && userInstruction.trim()) {
    userContent += `User Extraction Goal:\n${userInstruction.trim()}\n\n`;
  }

  if (previousTail && previousTail.trim()) {
    userContent += `--- CONTINUATION CONTEXT ---\nThe previous section concluded with:\n"...${previousTail.slice(-300).trim()}..."\nContinue extracting seamlessly from this point without repeating content.\n-----------------------------\n\n`;
  }

  if (textUnits.length > 0) {
    userContent += `--- DOCUMENT TEXT CONTENT (Batch ${batchIndex} of ${totalBatches}) ---\n`;
    for (const u of textUnits) {
      userContent += `\n### [${u.fileName} - Page ${u.pageNum}]\n${u.text}\n`;
    }
    userContent += `\n----------------------------------------------------\n`;
  }

  if (base64Images.length > 0) {
    userContent += `\nPlease also extract and structure all information accurately from the ${base64Images.length} attached document image(s).\n`;
  }

  userContent += `\nExtract all information faithfully. Format the output cleanly using Markdown (tables, bullet points, headings, and bold text). Return raw Markdown only.`;

  const systemPrompt = `You are a world-class Data Extractor AI. Your job is to extract all data and information accurately from the provided document content and/or images.
CRITICAL RULES:
1. Never invent, alter, or falsify names, dates, numbers, financial amounts, or facts.
2. Fix only obvious spelling mistakes or grammatical typos in extracted text.
3. Structure complex or tabular data into well-formatted Markdown tables.
4. Use clear headings, bullet points, and numbered lists where appropriate.
5. Return ONLY clean, valid raw Markdown.`;

  const rawResult = await callAI(systemPrompt, userContent, base64Images);
  if (!rawResult || !rawResult.trim()) {
    throw new Error(`AI returned an empty response for Batch ${batchIndex}.`);
  }

  // Validate Markdown syntax and heal if broken
  const validatedResult = await fixCodeSyntaxWithAI({
    code: rawResult,
    type: "markdown",
    validatorFn: (code) => checkMarkdownSyntax(code),
  });

  return validatedResult;
}

// ----------------------------------------------------------------------------
// API: Extract Data (Single Request)
// ----------------------------------------------------------------------------
app.post("/api/extract-data", async (req, res) => {
  try {
    const { prompt, filePaths } = req.body;
    if (!prompt || !filePaths || filePaths.length === 0) {
      return res.status(400).json({
        success: false,
        error: "Prompt and at least one file are required.",
      });
    }

    const units = await prepareExtractionUnits(filePaths);
    if (units.length === 0) {
      throw new Error("No valid document pages or images could be processed.");
    }

    const result = await executeBatchExtraction(units, 1, 1, prompt, "");
    res.json({ success: true, result });
  } catch (error) {
    console.error("Extraction error:", error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// In-memory store for batch processing results: { batchId: { results: [], pending: true } }
const batchStore = new Map();

// Cleanup expired batches every 5 minutes
setInterval(
  () => {
    const now = Date.now();
    for (const [batchId, batch] of batchStore.entries()) {
      if (now - batch.createdAt > 30 * 60 * 1000) {
        batchStore.delete(batchId);
        console.log(`Batch ${batchId} expired and cleaned up`);
      }
    }
  },
  5 * 60 * 1000,
);

// ----------------------------------------------------------------------------
// API: Extract Data - Hybrid Progressive (Batch Processing)
// First 5 items processed immediately, remaining batched asynchronously
// ----------------------------------------------------------------------------
app.post("/api/extract-batch", async (req, res) => {
  try {
    const { prompt, filePaths } = req.body;
    if (!prompt || !filePaths || filePaths.length === 0) {
      return res.status(400).json({
        success: false,
        error: "Prompt and at least one file are required.",
      });
    }

    const allUnits = await prepareExtractionUnits(filePaths);
    if (allUnits.length === 0) {
      throw new Error("No valid document pages or images could be processed.");
    }

    const batchId = uuidv4();
    const batchSize = 5;
    const totalBatches = Math.ceil(allUnits.length / batchSize);

    batchStore.set(batchId, {
      results: [],
      pending: true,
      totalBatches,
      currentBatch: 1,
      createdAt: Date.now(),
    });

    // Process first batch immediately for instant user feedback
    const firstBatchUnits = allUnits.slice(0, batchSize);
    const firstResult = await executeBatchExtraction(
      firstBatchUnits,
      1,
      totalBatches,
      prompt,
      ""
    );

    batchStore.get(batchId).results.push({
      batch: 1,
      result: firstResult,
      status: "complete",
    });

    // Process remaining batches in background
    if (allUnits.length > batchSize) {
      processRemainingBatches(batchId, allUnits, batchSize, prompt, firstResult);
    } else {
      batchStore.get(batchId).pending = false;
    }

    res.json({
      success: true,
      batchId,
      totalBatches,
      result: firstResult,
    });
  } catch (error) {
    console.error("Batch extraction error:", error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Helper: Process remaining batches asynchronously with sliding context
async function processRemainingBatches(batchId, allUnits, batchSize, userInstruction, initialTail = "") {
  try {
    const batchData = batchStore.get(batchId);
    if (!batchData) return;

    let previousTail = initialTail;

    for (let i = batchSize; i < allUnits.length; i += batchSize) {
      const batchUnits = allUnits.slice(i, i + batchSize);
      if (batchUnits.length > 0) {
        const batchNumber = Math.floor(i / batchSize) + 1;
        batchData.currentBatch = batchNumber;
        console.log(`[Batch Engine] Starting Batch ${batchNumber} of ${batchData.totalBatches} (${batchUnits.length} items)...`);
        
        const result = await executeBatchExtraction(
          batchUnits,
          batchNumber,
          batchData.totalBatches,
          userInstruction,
          previousTail
        );

        batchData.results.push({
          batch: batchNumber,
          result,
          status: "complete",
        });
        previousTail = result;
        console.log(`[Batch Engine] Batch ${batchNumber} completed successfully.`);
      }
    }

    batchData.pending = false;
    console.log(`[Batch Engine] All ${batchData.totalBatches} batches complete for session ${batchId}.`);
  } catch (error) {
    console.error("Background batch processing error:", error);
    if (batchStore.has(batchId)) {
      batchStore.get(batchId).error = error.message;
      batchStore.get(batchId).pending = false;
    }
  }
}

// API: Poll for batch results
app.get("/api/get-batch-result/:batchId", (req, res) => {
  try {
    const { batchId } = req.params;
    const batch = batchStore.get(batchId);

    if (!batch) {
      return res.status(404).json({ success: false, error: "Batch session not found or expired" });
    }

    res.json({
      success: true,
      batchId,
      results: batch.results,
      totalBatches: batch.totalBatches,
      currentBatch: batch.currentBatch,
      pending: batch.pending,
      error: batch.error || null,
    });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// ----------------------------------------------------------------------------
// API: Generate Presentation
// ----------------------------------------------------------------------------
app.post("/api/generate", async (req, res) => {
  const uploadedFilePaths = []; // Track files to cleanup
  const streamProgress = (req.get("accept") || "").includes(
    "application/x-ndjson",
  );

  // The browser requests newline-delimited JSON so it can display milestones
  // from the real generation pipeline as they occur. Keep normal JSON replies
  // available for any older clients that do not request the stream.
  if (streamProgress) {
    res.setHeader("Content-Type", "application/x-ndjson; charset=utf-8");
    res.setHeader("Cache-Control", "no-cache, no-transform");
    res.setHeader("X-Accel-Buffering", "no");
    res.flushHeaders();
  }
  const send = (payload) => {
    if (streamProgress) res.write(`${JSON.stringify(payload)}\n`);
  };
  const sendStatus = (step, progress, title, subtitle, activity) => {
    send({ type: "status", step, progress, title, subtitle, activity });
  };
  const sendError = (status, message) => {
    if (streamProgress) {
      send({ type: "error", error: message });
      return res.end();
    }
    return res.status(status).json({ success: false, error: message });
  };

  try {
    const { content, palette, palettes, images } = req.body;
    if (!content) {
      return sendError(400, "Content is required");
    }

    // Batch mode intentionally shares one AI-authored slide plan across every
    // selected palette. This makes visual variants comparable slide-for-slide.
    const requestedPalettes =
      Array.isArray(palettes) && palettes.length
        ? palettes
        : [palette || "forestMoss"];
    const selectedPalettes = [...new Set(requestedPalettes)].filter(
      (name) => typeof name === "string" && VALID_PALETTES.has(name),
    );
    if (!selectedPalettes.length) {
      return sendError(400, "Select at least one valid color palette.");
    }
    if (selectedPalettes.length > 4) {
      return sendError(
        400,
        "You can generate up to four color variants at once.",
      );
    }
    const primaryPalette = selectedPalettes[0];

    sendStatus(
      1,
      8,
      "Preparing your content...",
      "Checking the topic, palette, and uploaded assets",
      "Preparing the presentation request",
    );

    const planId = uuidv4();
    const baseOutputPath = path.join(OUTPUT_DIR, `Presentation_${planId}.pptx`);

    // 1. Prepare the system prompt containing the skill instructions
    const skillPath = path.join(__dirname, "beautiful_pptx_skill.md");
    if (!fs.existsSync(skillPath)) {
      throw new Error(
        "beautiful_pptx_skill.md not found. This is the source of truth.",
      );
    }
    const skillContent = fs.readFileSync(skillPath, "utf8");
    sendStatus(
      1,
      18,
      "Preparing the design system...",
      "Loading the selected presentation layouts",
      "Loading layout rules and design tokens",
    );

    // Build image context for AI if images were provided
    let imageContext = "";
    if (images && images.length > 0) {
      imageContext = `\n\n# USER-PROVIDED IMAGES\nThe user has attached the following images. Use slide.addImage() to insert them where appropriate based on their descriptions.\nFor each image, use the ABSOLUTE path or URL provided.\n\n`;
      images.forEach((img, i) => {
        // For uploaded files, convert the static URL to absolute path for pptxgenjs
        let imageSrc;
        if (img.type === "upload" && img.serverPath) {
          imageSrc = img.serverPath.replace(/\\/g, "\\\\");
          uploadedFilePaths.push(img.serverPath); // Track for cleanup
        } else {
          imageSrc = img.url;
        }
        imageContext += `Image ${i + 1}:\n  - Source: "${imageSrc}"\n  - Type: ${img.type === "upload" ? "Local file (use path property)" : "URL (use path property)"}\n  - Description: "${img.description}"\n\n`;
      });
      imageContext += `To insert an image as a full slide background (without covering the text), do THIS after calling a layout function:\n`;
      imageContext += `  let s1 = addTitleSlide(pres, ...);\n`;
      imageContext += `  s1.background = { path: "<source>" }; // pptxgenjs automatically covers/centers backgrounds\n\n`;
      imageContext += `For layouts that explicitly accept an image (like addSplitInfoImage), pass the path directly as an argument.\n`;
      imageContext += `If you MUST insert a floating image, use: slide.addImage({ path: "<source>", x: ..., y: ..., w: ..., h: ..., sizing: { type: 'cover', w: ..., h: ... } });\n`;
      imageContext += `Match each image to the most relevant slide based on the description. If description says "logo" or "background" put it on the title slide, etc.\n`;
    }

    // Palette instruction
    const paletteInstruction = `\nIMPORTANT: Build this shared slide plan using the "${primaryPalette}" color palette. The server will render the same plan in ${selectedPalettes.length} selected palette variant(s). You MUST use: const C = PALETTES.${primaryPalette};\nDo NOT choose a different palette.\n`;

    const systemPrompt = `You are a world-class Presentation AI Generator. 
You will be provided with a user's presentation content and instructions.
Your goal is to generate ONLY the Javascript code using the 'pptxgenjs' library to create the presentation.
DO NOT output any markdown blocks, just raw Javascript code that can be executed directly in a Node.js environment.

# PRIORITY 1 тАФ DATA FIDELITY AND OVERFLOW
You MUST NOT drop, skip, merge, or combine any separately named user data points. You may summarize supporting descriptions, fix grammar, and polish wording, but preserve every individual item and its order.
For an explicit sequence, each source stage MUST become exactly one array element. For example, "Farm тЖТ Collection тЖТ Processing тЖТ Storage тЖТ Transport тЖТ Retailer тЖТ Consumer" requires seven separate process-flow itemsтАФnever combine Retailer and Consumer into one card.
Pass the complete ordered array to its layout function in ONE call, even when it exceeds one slide's capacity. Do NOT truncate it or split it manually. The layout functions paginate overflow automatically.

# PRIORITY 2 тАФ SOURCE OF TRUTH (beautiful_pptx_skill.md)
The following is the complete design system and reference code. You MUST strictly follow these rules:
1. ONLY use the palette and font definitions defined in the skill file.
2. ONLY call the layout functions (like addTitleSlide, addCardGrid3, etc.) described in the skill file.
3. You must select ONE color palette from ┬з2. Do NOT mix palettes.${paletteInstruction}
4. Call the tested layout functions directly; do not invent or prefix function names.
5. Refer to the AI Layout Decision Engine in the skill file to select the layout for the content type.
6. The Javascript you output will be injected into a boilerplate that already defines the layout functions.
   Therefore, DO NOT redefine the layout functions (like addTitleSlide) in your output. Just call them!
   DO NOT redefine PALETTES or FONTS. They are already in the global scope.

# BOILERPLATE TO FOLLOW
Your output MUST start with exactly this:
const pptxgen = require('pptxgenjs');
let pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'AI Presentation Generator';

const C = PALETTES.${primaryPalette}; // Selected palette

// ... THEN MAKE ALL YOUR SLIDE FUNCTION CALLS HERE ...
// Example: addTitleSlide(pres, "My Title", "My Subtitle", "John Doe", C);
// ...

// SPEAKER NOTES
// Do not generate direct addNotes(...) calls. Some layout functions auto-create
// continuation slides and return an array rather than a single Slide. If notes
// are needed, use addSpeakerNotes(result, "notes"); it is supplied by the
// execution wrapper and safely supports either return type.
// Do NOT define, redefine, or wrap addSpeakerNotes. It already exists in the
// execution wrapper. Output only the layout calls and optional note calls.

pres.writeFile({ fileName: '${baseOutputPath.replace(/\\/g, "\\\\")}' })
  .then(() => console.log('Done'))
  .catch(err => console.error(err));
${imageContext}
Here is the full skill reference for your knowledge:
=============================================================================
${skillContent}
=============================================================================

Remember: Output ONLY RAW JAVASCRIPT CODE. Do not wrap in \`\`\`javascript or any markdown.
`;

    // 2. Call the AI model
    sendStatus(
      2,
      35,
      "AI is designing the slides...",
      "Creating the presentation structure and layout code",
      "Generating slide structure and layout instructions",
    );
    let aiCode = await callAI(systemPrompt, content);

    if (!aiCode || aiCode.trim() === "") {
      throw new Error("AI returned empty response.");
    }

    // Step 2: Check code for syntax errors by node command, and auto-heal with AI if needed
    sendStatus(
      2,
      48,
      "Checking code syntax...",
      "Verifying JavaScript syntax using Node.js compiler",
      "Checking code for syntax errors",
    );
    aiCode = await fixCodeSyntaxWithAI({
      code: aiCode,
      type: "pptx",
      validatorFn: (code) => checkPptxSyntax(code, skillContent),
      onProgress: (msg) =>
        sendStatus(
          2,
          52,
          "Resolving syntax issue...",
          msg,
          "Self-healing syntax error",
        ),
    });

    sendStatus(
      2,
      55,
      "AI design is ready...",
      "Preparing the approved layout code for rendering",
      "Validating the generated slide instructions",
    );

    // 3. Render the exact same AI-authored slide plan once per selected palette.
    // The isolated Node processes make up to four variants run concurrently.
    sendStatus(
      3,
      68,
      `Building ${selectedPalettes.length} presentation variant${selectedPalettes.length === 1 ? "" : "s"}...`,
      "Applying your selected color palettes to the shared slide plan",
      "Rendering PowerPoint variants",
    );
    const files = await Promise.all(
      selectedPalettes.map(async (variantPalette) => {
        const fileId = uuidv4();
        const outputPath = path.join(OUTPUT_DIR, `Presentation_${fileId}.pptx`);
        const scriptPath = path.join(OUTPUT_DIR, `script_${fileId}.js`);
        const variantAiCode = createPaletteVariant(
          aiCode,
          variantPalette,
          outputPath,
        );
        fs.writeFileSync(
          scriptPath,
          prepareExecutionCode(variantAiCode, skillContent),
          "utf8",
        );
        try {
          await runGeneratorScript(scriptPath);
        } finally {
          if (fs.existsSync(scriptPath)) fs.unlinkSync(scriptPath);
        }
        return {
          fileId,
          palette: variantPalette,
          fileName: `Presentation_${variantPalette}.pptx`,
        };
      }),
    );
    sendStatus(
      3,
      92,
      "Finalizing your PowerPoint...",
      `${files.length} presentation variant${files.length === 1 ? "" : "s"} created successfully`,
      "Packaging completed presentations",
    );

    // Cleanup uploaded images after successful generation
    cleanupUploadedImages(uploadedFilePaths);

    const result = {
      success: true,
      fileId: files[0].fileId,
      fileName: files[0].fileName,
      files,
    };
    if (streamProgress) {
      send({ type: "result", ...result });
      res.end();
    } else {
      res.json(result);
    }
  } catch (error) {
    console.error("Error generating presentation:", error);
    // Still cleanup on error
    cleanupUploadedImages(uploadedFilePaths);
    sendError(500, error.message);
  }
});

// ----------------------------------------------------------------------------
// Generation helpers
// ----------------------------------------------------------------------------
function createPaletteVariant(aiCode, palette, outputPath) {
  // Treat the model response as a slide plan, not as a fragile API contract.
  // A valid plan can omit our palette declaration or final save call.
  const paletteLine = `const C = PALETTES.${palette}; // Selected palette`;
  let variantCode = String(aiCode || "")
    .replace(/^\s*```(?:javascript|js)?\s*$/gim, "")
    .replace(/^\s*const\s+C\s*=\s*PALETTES\.[A-Za-z0-9_]+\s*;[^\n]*\n?/gim, "")
    .trim();

  const authorLine = /pres\.author\s*=\s*[^;]+;/i;
  if (authorLine.test(variantCode)) {
    variantCode = variantCode.replace(
      authorLine,
      (match) => `${match}\n${paletteLine}`,
    );
  } else {
    variantCode = `${paletteLine}\n${variantCode}`;
  }

  // The backend owns output paths. Remove any model-provided tail and append
  // a known-good write operation so a missing/changed writeFile call cannot
  // break a request or overwrite another variant.
  const writeFileIndex = variantCode.search(
    /\b(?:await\s+)?pres\.writeFile\s*\(/i,
  );
  if (writeFileIndex !== -1) {
    variantCode = variantCode.slice(0, writeFileIndex).trimEnd();
  }
  const safeOutputPath = outputPath.replace(/\\/g, "\\\\");
  return `${variantCode}\n\npres.writeFile({ fileName: '${safeOutputPath}' })\n  .then(() => console.log('Done'))\n  .catch(err => { console.error(err); process.exitCode = 1; });\n`;
}

const MAX_CONCURRENT_GENERATORS = 4;
const MAX_QUEUED_GENERATORS = 8;
const GENERATOR_TIMEOUT_MS = 300_000;
const GENERATOR_KILL_GRACE_MS = 5_000;
const WINDOWS_TREE_KILL_TIMEOUT_MS = 5_000;
const activeGeneratorProcesses = new Set();
let activeGeneratorCount = 0;
const generatorWaitQueue = [];

function acquireGeneratorSlot() {
  if (activeGeneratorCount < MAX_CONCURRENT_GENERATORS) {
    activeGeneratorCount++;
    return Promise.resolve();
  }
  if (generatorWaitQueue.length >= MAX_QUEUED_GENERATORS) {
    return Promise.reject(new Error("The presentation renderer is busy. Please try again shortly."));
  }
  return new Promise((resolve) => generatorWaitQueue.push(resolve));
}

function releaseGeneratorSlot() {
  const next = generatorWaitQueue.shift();
  if (next) next();
  else activeGeneratorCount--;
}

function terminateGeneratorTree(child, signal) {
  if (!child || !child.pid) return;
  if (process.platform === "win32") {
    // child.kill() only targets the Node renderer on Windows. taskkill /T
    // also terminates descendants that renderer code may have started.
    execFile("taskkill", ["/PID", String(child.pid), "/T", "/F"], {
      windowsHide: true,
      timeout: WINDOWS_TREE_KILL_TIMEOUT_MS,
    }, (error) => {
      if (error && error.code !== "ESRCH" && error.code !== "ENOENT") {
        try { child.kill(signal); } catch (killError) {
          if (killError.code !== "ESRCH") console.error("Renderer termination failed:", killError.message);
        }
      }
    });
    return;
  }
  try {
    process.kill(-child.pid, signal);
  } catch (error) {
    if (error.code !== "ESRCH") {
      try { child.kill(signal); } catch (killError) {
        if (killError.code !== "ESRCH") console.error("Renderer termination failed:", killError.message);
      }
    }
  }
}

// Detached POSIX renderers run in their own process group so timeouts and
// server shutdown can clean up the renderer and any processes it starts.
process.once("exit", () => {
  for (const child of activeGeneratorProcesses) {
    terminateGeneratorTree(child, "SIGTERM");
  }
});

function runGeneratorScript(scriptPath) {
  return acquireGeneratorSlot().then(() => new Promise((resolve, reject) => {
    // Use process.execPath instead of 'node' so child processes work on
    // cPanel/Passenger where the node binary isn't in the default PATH.
    const child = execFile(
      process.execPath,
      [scriptPath],
      { detached: process.platform !== "win32" },
      (error, stdout, stderr) => {
        clearTimeout(timeout);
        clearTimeout(forceKillTimeout);
        activeGeneratorProcesses.delete(child);
        releaseGeneratorSlot();
        if (error) {
          console.error(`Execution error: ${error}\nStderr: ${stderr}`);
          reject(
            new Error(timedOut
              ? "Presentation rendering exceeded the 5-minute time limit."
              : "Failed to generate a PPTX variant during rendering."),
          );
          return;
        }
        resolve(stdout);
      },
    );
    activeGeneratorProcesses.add(child);
    let forceKillTimeout;
    let timedOut = false;
    const timeout = setTimeout(() => {
      timedOut = true;
      terminateGeneratorTree(child, "SIGTERM");
      forceKillTimeout = setTimeout(
        () => terminateGeneratorTree(child, "SIGKILL"),
        GENERATOR_KILL_GRACE_MS,
      );
      forceKillTimeout.unref();
    }, GENERATOR_TIMEOUT_MS);
    timeout.unref();
  }));
}

// ----------------------------------------------------------------------------
// Cleanup uploaded images
// ----------------------------------------------------------------------------
function cleanupUploadedImages(filePaths) {
  filePaths.forEach((fp) => {
    try {
      if (fp && fs.existsSync(fp)) {
        fs.unlinkSync(fp);
        console.log(`ЁЯз╣ Cleaned up uploaded image: ${path.basename(fp)}`);
      }
    } catch (e) {
      console.error(`Failed to cleanup ${fp}:`, e.message);
    }
  });
}

// ----------------------------------------------------------------------------
// API: Download Presentation
// ----------------------------------------------------------------------------
app.get("/api/download/:fileId", (req, res) => {
  const { fileId } = req.params;

  // Basic path traversal protection
  if (fileId.includes("/") || fileId.includes("\\") || fileId.includes("..")) {
    return res.status(400).send("Invalid file ID");
  }

  const fileName = `Presentation_${fileId}.pptx`;
  const filePath = path.join(OUTPUT_DIR, fileName);

  if (fs.existsSync(filePath)) {
    res.download(filePath, "AI_Presentation.pptx", (err) => {
      if (err) {
        console.error("Error downloading file:", err);
      }
    });
  } else {
    res.status(404).send("File not found or expired");
  }
});

// ----------------------------------------------------------------------------
// API: Generate Document
// ----------------------------------------------------------------------------
app.post("/api/generate-document", async (req, res) => {
  const uploadedFilePaths = [];
  try {
    const { content, images = [] } = req.body || {};
    if (!content || !String(content).trim()) {
      return res.status(400).json({
        success: false,
        error: "Content and instructions are required.",
      });
    }

    const skillPath = path.join(__dirname, "docx_skill.md");
    if (!fs.existsSync(skillPath)) throw new Error("docx_skill.md not found.");
    const skill = fs.readFileSync(skillPath, "utf8");
    const outputId = uuidv4();
    const outputPath = path.join(OUTPUT_DIR, `Document_${outputId}.docx`);
    const imageContext = images
      .map((image, index) => {
        if (image.type === "upload" && image.serverPath)
          uploadedFilePaths.push(image.serverPath);
        return `Image ${index + 1}: description: ${String(image.description || "No description supplied")}.`;
      })
      .join("\n");
    const systemPrompt = `You are a professional Word document architect. Turn the user's brief into a polished, faithful document plan. Choose an appropriate document archetype and layout based on the reference skill. Preserve all facts. Use supplied image DESCRIPTIONS only to decide placement; do not infer image content from filenames, URLs, upload paths, or other metadata. Do not change dates, values, names, or topic. Do not merge, drop, or reorder user data. Return ONLY valid JSON matching this schema:\n{"title":"string","subtitle":"string","documentType":"string","designPreset":"standard_business|formal_administrative_register|formal_letter|professional_cv|business_proposal","metadata":["Label: value"],"letterMeta":{"date":"string","recipientLines":["string"],"subject":"string","salutation":"string","closing":"string","signatory":"string","signatoryTitle":"string","signatoryOrg":"string"},"cvMeta":{"name":"string","title":"string","contact":[{"icon":"phone|email|link","value":"string"}],"education":[{"degree":"string","institution":"string","year":"string"}],"skills":["string"],"languages":["string"],"profile":"string","experience":[{"title":"string","company":"string","date":"string","bullets":["string"]}],"projects":[{"name":"string","description":"string","link":"string"}]},"proposalMeta":{"title":"string","subtitle":"string","preparedFor":"string","preparedBy":"string","date":"string"},"sections":[{"heading":"string","paragraphs":["string"],"bullets":["string"],"numbered":["string"],"callout":"string","signatures":[{"role":"string","name":"string"}],"table":{"headers":["string"],"rows":[["string"]]},"bigNumberList":[{"number":"string","title":"string","text":"string"}],"dividedList":[{"title":"string","subtitle":"string"}]}],"imagePlacements":[{"imageIndex":1,"afterSection":0,"caption":"string","placement":"wide"}]}\nChoose professional_cv for modern two-column CVs or resumes. When using professional_cv, you MUST populate the cvMeta object with name, title, contact, education, skills, languages, profile, experience, and projects. Use the image (if provided) as the profile photo.\nChoose formal_letter for applications, দরখাস্ত, formal letters, cover letters, requests, complaints, or any correspondence addressed to a specific person or authority. When using formal_letter, you MUST populate the letterMeta object. The body paragraphs go in sections[0].paragraphs as usual.\nChoose formal_administrative_register for attendance, beneficiary, student/patient, field-report, official-record, or grouped register data. Use tables only for genuine comparable row/column data. Keep prose concise and well structured. Do not emit markdown or code.\nChoose business_proposal for pitches, digital marketing strategies, or highly styled project plans. Populate the proposalMeta object.\n\nREFERENCE SKILL:\n${skill}\n\nUSER IMAGES (descriptions only):\n${imageContext || "None"}`;
    const rawPlan = await callAI(systemPrompt, String(content));
    if (!rawPlan || !rawPlan.trim()) {
      throw new Error("AI returned empty response for document plan.");
    }

    // Step 2: Check document plan for syntax errors by node command, and auto-heal with AI if needed
    console.log("[DOCX] Checking document plan syntax with Node.js...");
    const validatedRawPlan = await fixCodeSyntaxWithAI({
      code: rawPlan,
      type: "docx",
      validatorFn: (code) => checkDocxSyntax(code),
    });

    const plan = parseDocumentPlan(validatedRawPlan, content, images.length);
    const buffer =
      plan.designPreset === "professional_cv"
        ? await buildCVDocx(plan, images)
        : plan.designPreset === "formal_letter"
          ? await buildLetterDocx(plan, images)
          : plan.designPreset === "business_proposal"
            ? await buildProposalDocx(plan, images)
            : await buildDocxBuffer(plan, images);
    fs.writeFileSync(outputPath, buffer);

    const debugScriptPath = path.join(
      OUTPUT_DIR,
      `Document_${outputId}.debug.js`,
    );
    const planJsonPath = path.join(
      OUTPUT_DIR,
      `Document_${outputId}.plan.json`,
    );
    const rawPlanPath = path.join(
      OUTPUT_DIR,
      `Document_${outputId}.raw_plan.txt`,
    );

    fs.writeFileSync(
      debugScriptPath,
      createDocumentDebugScript(outputId, plan, images),
    );
    fs.writeFileSync(planJsonPath, JSON.stringify(plan, null, 2));
    fs.writeFileSync(rawPlanPath, String(validatedRawPlan).trim());

    cleanupUploadedImages(uploadedFilePaths);
    res.json({
      success: true,
      fileId: outputId,
      fileName: `${safeFileName(plan.title || "AI Document")}.docx`,
      debugFileName: `Document_${outputId}.debug.js`,
      planFileName: `Document_${outputId}.plan.json`,
      rawPlanFileName: `Document_${outputId}.raw_plan.txt`,
    });
  } catch (error) {
    console.error("Document generation error:", error);
    cleanupUploadedImages(uploadedFilePaths);
    res.status(500).json({
      success: false,
      error: error.message || "Unable to create the document.",
    });
  }
});

app.get("/api/download-document/:fileId", (req, res) => {
  const { fileId } = req.params;
  if (!/^[a-f0-9-]+$/i.test(fileId))
    return res.status(400).send("Invalid file ID");
  const filePath = path.join(OUTPUT_DIR, `Document_${fileId}.docx`);
  if (!fs.existsSync(filePath))
    return res.status(404).send("File not found or expired");
  res.download(filePath, "AI_Document.docx");
});

// ----------------------------------------------------------------------------
// AI Integration тАФ 3-Key Failover Pool
// ----------------------------------------------------------------------------

/**
 * Build the list of active API keys from environment variables.
 * Filters out placeholder values so the system only tries real keys.
 */
function getApiKeyPool() {
  const keys = [];
  // Dynamically load up to 50 keys (GEMINI_API_KEY_1 to GEMINI_API_KEY_50)
  for (let i = 1; i <= 50; i++) {
    const k = process.env[`GEMINI_API_KEY_${i}`];
    // Gemini API keys are usually ~39 characters. Require at least 20 to prevent accidental empty/bad values.
    if (
      k &&
      typeof k === "string" &&
      k.trim().length > 20 &&
      !k.includes("PASTE_YOUR_")
    ) {
      keys.push(k.trim());
    }
  }
  // Legacy fallback тАФ still works if someone uses OPENAI_API_KEY directly
  const fallback = process.env.OPENAI_API_KEY;
  if (fallback && typeof fallback === "string" && fallback.trim().length > 20) {
    keys.push(fallback.trim());
  }

  // Deduplicate (in case OPENAI_API_KEY duplicates another key)
  return [...new Set(keys)];
}

/**
 * Try a single API key. Returns the generated code string on success.
 * Throws an error on failure so the caller can try the next key.
 */
async function tryKeyOnce(apiKey, systemPrompt, userPrompt, base64Images = [], customBaseUrl = null, customModel = null) {
  const baseUrl = (
    customBaseUrl || process.env.OPENAI_BASE_URL ||
    "https://generativelanguage.googleapis.com/v1beta/openai"
  ).replace(/\/$/, "");
  const model =
    customModel || process.env.OPENAI_MODEL || "gemini-2.5-flash-lite-preview-06-17";

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 180_000); // 3-minute timeout per key

  let userContent = userPrompt;
  if (base64Images && base64Images.length > 0) {
    userContent = [{ type: "text", text: userPrompt }];
    base64Images.forEach((b64) => {
      userContent.push({
        type: "image_url",
        image_url: { url: `data:image/jpeg;base64,${b64}` },
      });
    });
  }

  let response;
  try {
    const isGoogle = baseUrl.includes("generativelanguage.googleapis.com");
    const headers = {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    };
    if (isGoogle) {
      headers["x-goog-api-key"] = apiKey;
    }
    const requestUrl = isGoogle
      ? `${baseUrl}/chat/completions?key=${encodeURIComponent(apiKey)}`
      : `${baseUrl}/chat/completions`;

    response = await fetch(requestUrl, {
      method: "POST",
      headers,
      body: JSON.stringify({
        model,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userContent },
        ],
        temperature: 0.2,
      }),
      signal: controller.signal,
    });
  } finally {
    clearTimeout(timeout);
  }

  if (!response.ok) {
    const errText = await response.text().catch(() => response.statusText);
    throw new Error(`HTTP ${response.status}: ${errText}`);
  }

  const data = await response.json();
  let code = data.choices[0].message.content;

  // Strip markdown wrappers if the model still added them
  code = code.replace(/^```javascript\s*/gm, "");
  code = code.replace(/^```js\s*/gm, "");
  code = code.replace(/^```\s*/gm, "");

  return code;
}

/**
 * Decide whether an error should trigger a failover to the next key.
 * Rate-limits, busy signals, timeouts, and network errors all qualify.
 */
function isFailoverError(err) {
  if (!err) return false;
  const msg = err.message || "";
  // HTTP 400, 401, 402, 403, 404, 429, 500, 502, 503, 504
  if (/HTTP (400|401|402|403|404|429|500|502|503|504)/i.test(msg)) return true;
  // Auth & key validity issues
  if (/API key not valid|UNAUTHENTICATED|ACCESS_TOKEN|unauthorized|invalid api key|forbidden/i.test(msg)) return true;
  // Timeout / network abort
  if (
    err.name === "AbortError" ||
    /timeout|ETIMEDOUT|ECONNRESET|ECONNREFUSED/i.test(msg)
  )
    return true;
  // Undici / fetch connection errors (connect timeouts, DNS lookups, socket errors)
  if (/Connect Timeout Error|UND_ERR_CONNECT_TIMEOUT|fetch failed/i.test(msg))
    return true;
  const causeCode = (err.cause && (err.cause.code || "")) || "";
  if (
    /(UND_ERR_CONNECT_TIMEOUT|UND_ERR_SOCKET|ENOTFOUND|EAI_AGAIN|ECONNRESET|ECONNREFUSED|ETIMEDOUT)/i.test(
      causeCode,
    )
  )
    return true;
  // Google-specific quota messages
  if (
    /quota|rate.?limit|too many request|resource.?exhausted|overloaded/i.test(
      msg,
    )
  )
    return true;
  return true; // Default to true in rotation pool so one bad key never halts the pool
}

/**
 * Main AI caller with automatic failover across the key pool.
 */
async function callAI(systemPrompt, userPrompt, base64Images = []) {
  let pool = getApiKeyPool();

  if (pool.length === 0 && !process.env.OPENROUTER_API_KEY) {
    throw new Error(
      "No API keys configured. Please set GEMINI_API_KEY_1 or OPENROUTER_API_KEY in your .env file."
    );
  }

  // Fisher-Yates shuffle to randomize the API key order for this specific request
  for (let i = pool.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [pool[i], pool[j]] = [pool[j], pool[i]];
  }

  let lastError;

  for (let i = 0; i < pool.length; i++) {
    // Use a masked prefix for logs since order is randomized
    const keyMask =
      typeof pool[i] === "string" && pool[i].length > 15
        ? pool[i].substring(0, 10) + "..."
        : `KeySlot_${i}`;

    try {
      console.log(`ЁЯдЦ Trying API key (${keyMask})тАж`);
      const code = await tryKeyOnce(
        pool[i],
        systemPrompt,
        userPrompt,
        base64Images,
      );
      if (i > 0) {
        console.log(`тЬЕ API (${keyMask}) succeeded after ${i} failover(s).`);
      }
      return code;
    } catch (err) {
      lastError = err;
      if (i < pool.length - 1) {
        console.warn(
          `⚠️  API (${keyMask}) failed (${err.message.slice(0, 100)}). Rotating to next key (${i + 2}/${pool.length})…`,
        );
      } else {
        console.error(
          `❌ All ${pool.length} API key(s) in pool exhausted. Last error: ${err.message}`,
        );
      }
    }
  }

  // Fallback to OpenRouter if configured
  const openRouterKey = process.env.OPENROUTER_API_KEY;
  if (openRouterKey) {
    try {
      console.log(`⚠️ All Gemini keys failed. Falling back to OpenRouter API...`);
      return await tryKeyOnce(
        openRouterKey,
        systemPrompt,
        userPrompt,
        base64Images,
        process.env.OPENROUTER_BASE_URL || "https://openrouter.ai/api/v1",
        process.env.OPENROUTER_MODEL || "z-ai/glm-5.2:free"
      );
    } catch (err) {
      console.error(`❌ OpenRouter fallback failed: ${err.message}`);
      throw new Error(`OpenRouter fallback failed: ${err.message}`);
    }
  }

  throw new Error(
    `All ${pool.length} Gemini API key(s) are currently unavailable. Please try again in a moment.`
  );
}

// ----------------------------------------------------------------------------
// Syntax Verification via Node.js Commands & Self-Healing AI Loop
// ----------------------------------------------------------------------------

/**
 * 1. Checks PPTX (pptxgenjs) JavaScript code for syntax errors using `node --check`.
 * Validates the complete executable code without running side effects.
 */
function checkPptxSyntax(aiCode, skillContent = "") {
  const tempFile = path.join(OUTPUT_DIR, `temp_syntax_pptx_${uuidv4()}.js`);
  try {
    let scriptToTest = aiCode;
    if (skillContent) {
      try {
        scriptToTest = prepareExecutionCode(aiCode, skillContent);
      } catch (prepErr) {
        return {
          valid: false,
          error: `Failed to prepare execution code: ${prepErr.message}`,
        };
      }
    }
    fs.writeFileSync(tempFile, scriptToTest, "utf8");
    execFileSync(process.execPath, ["--check", tempFile], {
      encoding: "utf8",
      timeout: 10000,
    });
    return { valid: true };
  } catch (err) {
    return {
      valid: false,
      error: (err.stderr || err.stdout || err.message || "").trim(),
    };
  } finally {
    if (fs.existsSync(tempFile)) {
      try { fs.unlinkSync(tempFile); } catch (_) {}
    }
  }
}

/**
 * 2. Checks DOCX document plan (JSON / JS object) for syntax errors using Node.js.
 * Validates structure via `node --check` and verifies object / JSON parsing.
 */
function checkDocxSyntax(rawPlan) {
  const cleaned = String(rawPlan || "")
    .replace(/^```(?:json|javascript|js)?\s*/gim, "")
    .replace(/```\s*$/gim, "")
    .trim();

  const tempJsFile = path.join(OUTPUT_DIR, `temp_syntax_docx_${uuidv4()}.js`);
  const tempJsonFile = path.join(OUTPUT_DIR, `temp_syntax_docx_${uuidv4()}.json`);

  try {
    // Check JavaScript object syntax via node --check
    fs.writeFileSync(tempJsFile, `const plan = ${cleaned};`, "utf8");
    execFileSync(process.execPath, ["--check", tempJsFile], {
      encoding: "utf8",
      timeout: 10000,
    });

    // Check JSON / Object evaluation via Node process
    fs.writeFileSync(tempJsonFile, cleaned, "utf8");
    const nodeParseScript = `
const fs = require('fs');
const content = fs.readFileSync(process.argv[1], 'utf8');
try {
  JSON.parse(content);
} catch (jsonErr) {
  try {
    new Function('return (' + content + ')');
  } catch (jsErr) {
    console.error('SyntaxError in DOCX document plan: ' + (jsonErr.message || jsErr.message));
    process.exit(1);
  }
}
process.exit(0);
`;
    execFileSync(process.execPath, ["-e", nodeParseScript, tempJsonFile], {
      encoding: "utf8",
      timeout: 10000,
    });

    return { valid: true };
  } catch (err) {
    return {
      valid: false,
      error: (err.stderr || err.stdout || err.message || "").trim(),
    };
  } finally {
    if (fs.existsSync(tempJsFile)) {
      try { fs.unlinkSync(tempJsFile); } catch (_) {}
    }
    if (fs.existsSync(tempJsonFile)) {
      try { fs.unlinkSync(tempJsonFile); } catch (_) {}
    }
  }
}

/**
 * 3. Checks HTML code (markup from Data Extractor) for syntax errors using Node.js.
 * Validates matching tags, hierarchy, unclosed attribute quotes, and unclosed tags.
 */
function checkHtmlSyntax(htmlContent) {
  const html = String(htmlContent || "");
  // If no HTML tags are found in the content, it's plain text/markdown without HTML errors
  if (!/<[a-zA-Z0-9]+[^>]*>/i.test(html)) {
    return { valid: true };
  }

  const tempFile = path.join(OUTPUT_DIR, `temp_syntax_html_${uuidv4()}.html`);
  fs.writeFileSync(tempFile, html, "utf8");

  const htmlCheckerScript = `
const fs = require('fs');
const filePath = process.argv[1];
const content = fs.readFileSync(filePath, 'utf8');

const voidTags = new Set([
  'area','base','br','col','embed','hr','img','input','link','meta','param','source','track','wbr'
]);

const stack = [];
const tagRegex = /<!--[\\s\\S]*?-->|<(\\/?)([a-zA-Z0-9\\-]+)([^>]*)(\\/?)>/g;
let match;

while ((match = tagRegex.exec(content)) !== null) {
  if (match[0].startsWith('<!--')) continue;
  const isClosing = !!match[1];
  const tagName = match[2].toLowerCase();
  const attrs = match[3] || '';
  const isSelfClosing = !!match[4] || voidTags.has(tagName);

  // Check unclosed attribute quotes
  let inDouble = false;
  let inSingle = false;
  for (let i = 0; i < attrs.length; i++) {
    const ch = attrs[i];
    if (ch === '"' && !inSingle) inDouble = !inDouble;
    else if (ch === "'" && !inDouble) inSingle = !inSingle;
  }
  if (inDouble || inSingle) {
    const line = content.substring(0, match.index).split('\\n').length;
    console.error(\`\${filePath}:\${line}\\n\${match[0]}\\nSyntaxError: Unclosed attribute quote in HTML tag <\${tagName}>\`);
    process.exit(1);
  }

  if (isClosing) {
    if (voidTags.has(tagName)) continue;
    if (stack.length === 0) {
      const line = content.substring(0, match.index).split('\\n').length;
      console.error(\`\${filePath}:\${line}\\n\${match[0]}\\nSyntaxError: Unexpected closing tag </\${tagName}> without matching opening tag\`);
      process.exit(1);
    }
    const last = stack.pop();
    if (last.tagName !== tagName) {
      const line = content.substring(0, match.index).split('\\n').length;
      console.error(\`\${filePath}:\${line}\\n\${match[0]}\\nSyntaxError: Mismatched closing tag: expected </\${last.tagName}> (opened at line \${last.line}) but found </\${tagName}>\`);
      process.exit(1);
    }
  } else if (!isSelfClosing) {
    const line = content.substring(0, match.index).split('\\n').length;
    stack.push({ tagName, index: match.index, line });
  }
}

if (stack.length > 0) {
  const unclosed = stack.map(s => \`<\${s.tagName}> (opened at line \${s.line})\`).join(', ');
  const last = stack[stack.length - 1];
  console.error(\`\${filePath}:\${last.line}\\nSyntaxError: Unclosed HTML tag(s): \${unclosed}\`);
  process.exit(1);
}

process.exit(0);
`;

  try {
    execFileSync(process.execPath, ["-e", htmlCheckerScript, tempFile], {
      encoding: "utf8",
      timeout: 10000,
    });
    return { valid: true };
  } catch (err) {
    return {
      valid: false,
      error: (err.stderr || err.stdout || err.message || "").trim(),
    };
  } finally {
    if (fs.existsSync(tempFile)) {
      try { fs.unlinkSync(tempFile); } catch (_) {}
    }
  }
}

/**
 * 4. Checks Markdown content for syntax errors.
 * Validates unclosed code fences and malformed table delimiters.
 */
function checkMarkdownSyntax(markdownContent) {
  const text = String(markdownContent || "").trim();
  if (!text) {
    return { valid: false, error: "Extracted Markdown content is empty." };
  }

  // Check unclosed code fences (``` count must be even)
  const codeBlockMatches = text.match(/```/g);
  if (codeBlockMatches && codeBlockMatches.length % 2 !== 0) {
    return {
      valid: false,
      error: "SyntaxError in Markdown: Unclosed code block (odd number of ``` fences).",
    };
  }

  // Check broken table delimiters
  const lines = text.split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (/^\|?\s*[:\s]*-[-\s:]*\|/.test(line)) {
      const prevLine = i > 0 ? lines[i - 1].trim() : "";
      if (!prevLine || !prevLine.includes("|")) {
        return {
          valid: false,
          error: `SyntaxError in Markdown: Table delimiter at line ${i + 1} without a valid preceding header row.`,
        };
      }
    }
  }

  return { valid: true };
}

/**
 * Master Syntax Fixer: Checks code using Node command validator.
 * If syntax error is detected:
 * Feeds the previous code + Node.js error message back to AI.
 * Strict restriction: AI must ONLY fix the specific syntax error.
 * Does NOT generate new code, does not alter content or structure.
 */
async function fixCodeSyntaxWithAI({
  code,
  type,
  validatorFn,
  maxRetries = 2,
  onProgress = null,
}) {
  let currentCode = code;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    const syntaxResult = validatorFn(currentCode);
    if (syntaxResult.valid) {
      if (attempt > 0) {
        console.log(`✅ [Syntax Self-Heal] Successfully fixed ${type.toUpperCase()} syntax error after ${attempt} attempt(s).`);
      }
      return currentCode;
    }

    console.warn(`⚠️ [Syntax Check] Detected syntax error in ${type.toUpperCase()} (attempt ${attempt + 1}/${maxRetries + 1}):\n${syntaxResult.error}`);

    if (attempt === maxRetries) {
      console.error(`❌ [Syntax Check] Max retries (${maxRetries}) reached for ${type.toUpperCase()} syntax fix.`);
      return currentCode;
    }

    if (typeof onProgress === "function") {
      onProgress(`Fixing ${type.toUpperCase()} syntax error (attempt ${attempt + 1})...`);
    }

    let typeSpecificInstruction = "";
    if (type === "pptx") {
      typeSpecificInstruction = "The code is raw JavaScript for pptxgenjs. Return ONLY raw executable JavaScript code with NO markdown fences, NO backticks, and NO explanations.";
    } else if (type === "docx") {
      typeSpecificInstruction = "The code is a JSON document plan for docx. Return ONLY valid, parseable raw JSON with NO markdown fences, NO backticks, and NO explanations.";
    } else if (type === "html") {
      typeSpecificInstruction = "The code is HTML / extracted data markup. Return ONLY the raw corrected HTML markup with NO markdown fences, NO backticks, and NO explanations.";
    } else if (type === "markdown") {
      typeSpecificInstruction = "The content is structured Markdown. Return ONLY valid, complete Markdown with properly closed code fences and balanced tables. Do NOT wrap the entire response in outer code blocks, and provide NO explanations.";
    }

    const systemPrompt = `You are an expert code syntax repair assistant.
A syntax error was detected by the Node.js compiler/runtime in the previously generated code.

CRITICAL RESTRICTIONS AND RULES:
1. Fix ONLY the specific syntax error indicated by Node.js.
2. DO NOT generate anything new.
3. DO NOT change any existing text, wording, facts, dates, values, numbers, or section content.
4. DO NOT add or remove slides, sections, table rows, or columns.
5. DO NOT change styling, color palettes, or layout choices.
6. ONLY fix the exact syntax mistake (such as an unclosed tag, unmatched bracket or parenthesis, missing comma, illegal token, unescaped quote, or malformed syntax).
7. ${typeSpecificInstruction}`;

    const userPrompt = `Here is the Node.js syntax error report:
---------------- NODE.JS ERROR LOG ----------------
${syntaxResult.error}
----------------------------------------------------

Here is the previous generated code with the syntax error:
---------------- PREVIOUS CODE ----------------
${currentCode}
-----------------------------------------------

CRITICAL REQUIREMENT:
Fix ONLY the specific syntax error identified above. Do NOT generate anything new. Do NOT alter any content or structure. Output ONLY the corrected code:`;

    try {
      let repaired = await callAI(systemPrompt, userPrompt);
      if (repaired && repaired.trim()) {
        if (type === "pptx") {
          repaired = repaired
            .replace(/^```(?:javascript|js)?\s*/gim, "")
            .replace(/```\s*$/gim, "")
            .trim();
        } else if (type === "docx") {
          repaired = repaired
            .replace(/^```(?:json|javascript|js)?\s*/gim, "")
            .replace(/```\s*$/gim, "")
            .trim();
        } else if (type === "html") {
          repaired = repaired
            .replace(/^```(?:html|xml)?\s*/gim, "")
            .replace(/```\s*$/gim, "")
            .trim();
        } else if (type === "markdown") {
          repaired = repaired
            .replace(/^```(?:markdown)?\s*\r?\n/i, "")
            .replace(/\r?\n```\s*$/i, "")
            .trim();
        }
        currentCode = repaired;
      }
    } catch (aiErr) {
      console.error(`❌ [Syntax Self-Heal] AI call failed during syntax repair attempt: ${aiErr.message}`);
      break;
    }
  }

  return currentCode;
}

// ----------------------------------------------------------------------------
// Extract and combine the functions from skill.md with the AI code
// ----------------------------------------------------------------------------
function prepareExecutionCode(aiCode, skillContent) {
  // We parse skill.md the same way as generate_test.js to get all the layout functions
  const blocks = [];
  const lines = skillContent.split("\n");
  let inBlock = false;
  let currentBlock = [];

  for (const line of lines) {
    if (line.startsWith("\`\`\`javascript")) {
      inBlock = true;
      currentBlock = [];
    } else if (line.startsWith("\`\`\`") && inBlock) {
      inBlock = false;
      blocks.push(currentBlock.join("\n"));
    } else if (inBlock) {
      currentBlock.push(line);
    }
  }

  let jsCode = blocks.join("\n\n");

  // Remove the boilerplate from the skill functions because AI provides its own
  jsCode = jsCode.replace(/const pptxgen = require\('pptxgenjs'\);/g, "");
  jsCode = jsCode.replace(/let pres = new pptxgen\(\);/g, "");
  jsCode = jsCode.replace(/pres\.layout = 'LAYOUT_WIDE';/g, "");
  jsCode = jsCode.replace(/pres\.author = 'AI Presentation Generator';/g, "");
  jsCode = jsCode.replace(/const C = PALETTES\..*?;/g, "");
  jsCode = jsCode.replace(/pres\.writeFile[\s\S]*?\.catch[\s\S]*?\);/g, "");

  // Fix the template literals so they don't break in evaluation
  jsCode = jsCode.replace(
    /`\$\{col\.emoji\}  \$\{col\.title\}`/g,
    "col.emoji + '  ' + col.title",
  );
  jsCode = jsCode.replace(
    /`\$\{card\.emoji\}  \$\{card\.title\}`/g,
    "card.emoji + '  ' + card.title",
  );
  jsCode = jsCode.replace(
    /`\$\{emoji\}  \$\{titleText\}`/g,
    "emoji + '  ' + titleText",
  );

  // Layouts with overflow support can return either one Slide or an array of
  // Slides. Normalize AI-generated direct speaker-note calls to the helper so
  // notes work for both shapes of return value.
  // Some model responses previously redefined addSpeakerNotes with a recursive
  // implementation. Rename any generated declaration so the tested wrapper
  // below remains the only callable note helper.
  aiCode = aiCode.replace(
    /\bfunction\s+addSpeakerNotes\s*\(/g,
    "function __aiGeneratedSpeakerNotesIgnored(",
  );
  aiCode = aiCode.replace(
    /\b([A-Za-z_$][\w$]*)\.addNotes\s*\(/g,
    "addSpeakerNotes($1, ",
  );

  const speakerNotesHelper = `
function addSpeakerNotes(slideOrSlides, notes) {
  const slides = Array.isArray(slideOrSlides) ? slideOrSlides : [slideOrSlides];
  slides.filter(Boolean).forEach((slide) => {
    if (typeof slide.addNotes === 'function') slide.addNotes(notes);
  });
  return slideOrSlides;
}
`;

  return (
    jsCode +
    speakerNotesHelper +
    "\n// ==================== AI GENERATED CODE ====================\n\n" +
    aiCode
  );
}

function safeFileName(value) {
  return (
    String(value)
      .replace(/[<>:"/\\|?*\x00-\x1F]/g, "")
      .trim()
      .slice(0, 80) || "AI_Document"
  );
}

function parseDocumentPlan(raw, fallbackContent, imageCount) {
  const cleaned = String(raw || "")
    .replace(/```json|```/gi, "")
    .trim();

  const tryParseJson = (value) => {
    try {
      return JSON.parse(value);
    } catch (_) {
      return null;
    }
  };

  const extractJsonObject = (text) => {
    let depth = 0;
    let inString = false;
    let escape = false;
    let start = -1;

    for (let i = 0; i < text.length; i++) {
      const char = text[i];
      if (escape) {
        escape = false;
        continue;
      }
      if (char === "\\") {
        escape = true;
        continue;
      }
      if (char === '"') {
        inString = !inString;
        continue;
      }
      if (inString) continue;
      if (char === "{") {
        if (depth === 0) start = i;
        depth += 1;
      }
      if (char === "}") {
        depth -= 1;
        if (depth === 0 && start !== -1) {
          return text.slice(start, i + 1);
        }
      }
    }
    return null;
  };

  let plan = tryParseJson(cleaned);
  if (!plan) {
    const snippet =
      extractJsonObject(cleaned) || cleaned.match(/\{[\s\S]*\}/)?.[0];
    plan = snippet ? tryParseJson(snippet) : null;
  }

  if (!plan || typeof plan !== "object") {
    plan = {
      title: "Document",
      subtitle: "",
      documentType: "Business document",
      sections: [
        {
          heading: "AI Plan Parse Error",
          paragraphs: [
            "The AI response could not be parsed into a valid document plan. Please retry with a shorter or more specific request.",
          ],
        },
      ],
    };
  }

  plan.title = String(plan.title || "Untitled Document");
  plan.subtitle = String(plan.subtitle || "");
  const VALID_PRESETS = [
    "standard_business",
    "formal_administrative_register",
    "formal_letter",
    "professional_cv",
    "business_proposal",
  ];
  plan.designPreset = VALID_PRESETS.includes(plan.designPreset)
    ? plan.designPreset
    : "standard_business";
  plan.metadata = Array.isArray(plan.metadata)
    ? plan.metadata
        .map((value) => String(value))
        .filter(Boolean)
        .slice(0, 6)
    : [];

  // Preserve letterMeta for formal_letter preset
  if (
    plan.designPreset === "formal_letter" &&
    plan.letterMeta &&
    typeof plan.letterMeta === "object"
  ) {
    plan.letterMeta = {
      date: String(plan.letterMeta.date || ""),
      recipientLines: Array.isArray(plan.letterMeta.recipientLines)
        ? plan.letterMeta.recipientLines.map((l) => String(l))
        : [],
      subject: String(plan.letterMeta.subject || ""),
      salutation: String(plan.letterMeta.salutation || ""),
      closing: String(plan.letterMeta.closing || ""),
      signatory: String(plan.letterMeta.signatory || ""),
      signatoryTitle: String(plan.letterMeta.signatoryTitle || ""),
      signatoryOrg: String(plan.letterMeta.signatoryOrg || ""),
    };
  } else if (plan.designPreset === "formal_letter") {
    // AI chose formal_letter but forgot letterMeta тАФ provide empty defaults
    plan.letterMeta = {
      date: "",
      recipientLines: [],
      subject: "",
      salutation: "",
      closing: "",
      signatory: "",
      signatoryTitle: "",
      signatoryOrg: "",
    };
  }

  // Preserve cvMeta for professional_cv preset
  if (
    plan.designPreset === "professional_cv" &&
    plan.cvMeta &&
    typeof plan.cvMeta === "object"
  ) {
    plan.cvMeta = {
      name: String(plan.cvMeta.name || ""),
      title: String(plan.cvMeta.title || ""),
      contact: Array.isArray(plan.cvMeta.contact)
        ? plan.cvMeta.contact.map((c) => ({
            icon: String(c.icon || ""),
            value: String(c.value || ""),
          }))
        : [],
      education: Array.isArray(plan.cvMeta.education)
        ? plan.cvMeta.education.map((e) => ({
            degree: String(e.degree || ""),
            institution: String(e.institution || ""),
            year: String(e.year || ""),
          }))
        : [],
      skills: Array.isArray(plan.cvMeta.skills)
        ? plan.cvMeta.skills.map(String)
        : [],
      languages: Array.isArray(plan.cvMeta.languages)
        ? plan.cvMeta.languages.map(String)
        : [],
      profile: String(plan.cvMeta.profile || ""),
      experience: Array.isArray(plan.cvMeta.experience)
        ? plan.cvMeta.experience.map((exp) => ({
            title: String(exp.title || ""),
            company: String(exp.company || ""),
            date: String(exp.date || ""),
            bullets: Array.isArray(exp.bullets) ? exp.bullets.map(String) : [],
          }))
        : [],
      projects: Array.isArray(plan.cvMeta.projects)
        ? plan.cvMeta.projects.map((p) => ({
            name: String(p.name || ""),
            description: String(p.description || ""),
            link: String(p.link || ""),
          }))
        : [],
    };
  } else if (plan.designPreset === "professional_cv") {
    plan.cvMeta = {
      name: "",
      title: "",
      contact: [],
      education: [],
      skills: [],
      languages: [],
      profile: "",
      experience: [],
      projects: [],
    };
  }

  if (
    plan.designPreset === "business_proposal" &&
    plan.proposalMeta &&
    typeof plan.proposalMeta === "object"
  ) {
    plan.proposalMeta = {
      title: String(plan.proposalMeta.title || ""),
      subtitle: String(plan.proposalMeta.subtitle || ""),
      preparedFor: String(plan.proposalMeta.preparedFor || ""),
      preparedBy: String(plan.proposalMeta.preparedBy || ""),
      date: String(plan.proposalMeta.date || ""),
    };
  } else if (plan.designPreset === "business_proposal") {
    plan.proposalMeta = {
      title: "",
      subtitle: "",
      preparedFor: "",
      preparedBy: "",
      date: "",
    };
  }

  plan.sections =
    Array.isArray(plan.sections) && plan.sections.length
      ? plan.sections
      : [
          {
            heading: "AI Plan Parse Error",
            paragraphs: [
              "The AI response could not be parsed into a valid document plan. Please retry with a shorter or more specific request.",
            ],
          },
        ];
  plan.imagePlacements = Array.isArray(plan.imagePlacements)
    ? plan.imagePlacements.filter(
        (p) => Number(p.imageIndex) >= 1 && Number(p.imageIndex) <= imageCount,
      )
    : [];
  return plan;
}

function createDocumentDebugScript(outputId, plan, images) {
  const planJson = JSON.stringify(plan, null, 2);
  const imagesJson = JSON.stringify(images, null, 2);

  return `/**
 * Reproduction script for generated DOCX output.
 * Run from the output directory:
 *   node Document_${outputId}.debug.js
 */
const fs = require("fs");
const path = require("path");
const { buildDocxBuffer } = require("../server");

const plan = ${planJson};
const images = ${imagesJson};

(async function reproduce() {
  try {
    const buffer = await buildDocxBuffer(plan, images);
    const outPath = path.join(__dirname, "Document_${outputId}.repro.docx");
    fs.writeFileSync(outPath, buffer);
    console.log("Reproduced DOCX saved to:", outPath);
  } catch (error) {
    console.error("Reproduction failed:", error);
    process.exit(1);
  }
})();
`;
}

const TOKENS = {
  font: "Aptos",
  color: {
    ink: "111827",
    muted: "4B5563",
    border: "9FB3D1",
    white: "FFFFFF",
    accent: "244D89",
    accentDark: "152F5B",
    zebraA: "E8EDF5",
    zebraB: "FFFFFF",
    header: "244D89",
    sectionRule: "B5893A",
    labelFill: "2D5594",
    labelText: "FFFFFF",
  },
  size: {
    title: 46,
    subtitle: 23,
    body: 21,
    caption: 18,
    section: 26,
  },
};

const AUDIT_THEME = {
  color: {
    header: "244D89",
    summaryHeading: "244D89",
    labelFill: "2D5594",
    labelText: "FFFFFF",
    zebra: "E8EDF5",
    grid: "9FB3D1",
    heading: "1F3864",
    sectionRule: "B5893A",
  },
};

const BENGALI_DIGITS = [
  "০",
  "১",
  "২",
  "৩",
  "৪",
  "৫",
  "৬",
  "৭",
  "৮",
  "৯",
];
const toBengaliNumeral = (n) =>
  String(n)
    .split("")
    .map((d) => BENGALI_DIGITS[Number(d)] || d)
    .join("");
const bengaliParenPrefix = (n) => `${toBengaliNumeral(n)})`;
const bengaliDaariPrefix = (n) => `${toBengaliNumeral(n)}।`;
const subDecimalPrefix = (parent, child, lang = "bn") =>
  lang === "bn"
    ? `${toBengaliNumeral(parent)}.${toBengaliNumeral(child)}`
    : `${parent}.${child}`;

function parseTextWithSubSuper(text, baseOptions = {}) {
  const str = String(text || "");
  const runs = [];
  const regex = /(<sup[^>]*>.*?<\/sup>|<sub[^>]*>.*?<\/sub>)/gi;
  const parts = str.split(regex);
  
  for (const part of parts) {
    if (!part) continue;
    if (part.toLowerCase().startsWith("<sup")) {
      const match = part.match(/<sup[^>]*>(.*?)<\/sup>/i);
      const innerText = match ? match[1] : "";
      runs.push(new TextRun({
        text: innerText,
        superScript: true,
        ...baseOptions
      }));
    } else if (part.toLowerCase().startsWith("<sub")) {
      const match = part.match(/<sub[^>]*>(.*?)<\/sub>/i);
      const innerText = match ? match[1] : "";
      runs.push(new TextRun({
        text: innerText,
        subScript: true,
        ...baseOptions
      }));
    } else {
      runs.push(new TextRun({
        text: part,
        ...baseOptions
      }));
    }
  }
  
  if (runs.length === 0) {
    runs.push(new TextRun({ text: "", ...baseOptions }));
  }
  
  return runs;
}

function docText(text, options = {}) {
  const fontConfig = {
    ascii: TOKENS.font,
    cs: "Kalpurush",
    eastAsia: "Kalpurush",
    hAnsi: TOKENS.font,
    hint: "default"
  };

  let mergedFont = fontConfig;
  if (options.font) {
    if (typeof options.font === "string") {
      mergedFont = {
        ascii: options.font,
        cs: "Kalpurush",
        eastAsia: "Kalpurush",
        hAnsi: options.font,
        hint: "default"
      };
    } else {
      mergedFont = { ...fontConfig, ...options.font };
    }
  }

  const mergedOptions = {
    font: mergedFont,
    size: TOKENS.size.body,
    color: TOKENS.color.ink,
    ...options,
  };
  
  return parseTextWithSubSuper(text, mergedOptions);
}

const NO_BORDERS = {
  top: { style: BorderStyle.NONE, size: 0, color: TOKENS.color.border },
  bottom: { style: BorderStyle.NONE, size: 0, color: TOKENS.color.border },
  left: { style: BorderStyle.NONE, size: 0, color: TOKENS.color.border },
  right: { style: BorderStyle.NONE, size: 0, color: TOKENS.color.border },
};

const FORMAL_CELL_MARGINS = { top: 55, bottom: 55, left: 80, right: 80 };

const gridBorders = (color, size) => ({
  top: { style: BorderStyle.SINGLE, size, color },
  bottom: { style: BorderStyle.SINGLE, size, color },
  left: { style: BorderStyle.SINGLE, size, color },
  right: { style: BorderStyle.SINGLE, size, color },
  insideHorizontal: { style: BorderStyle.SINGLE, size, color },
  insideVertical: { style: BorderStyle.SINGLE, size, color },
});

const createDualHeader = (leftText, rightText, opts = {}) =>
  new Header({
    children: [
      new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        spacing: { after: 120 },
        border: {
          bottom: {
            color: opts.ruleColor || TOKENS.color.accent,
            style: BorderStyle.SINGLE,
            size: 6,
            space: 4,
          },
        },
        children: [
          new TextRun({
            text: leftText,
            bold: opts.boldLeft !== false,
            color: opts.textColor || TOKENS.color.accentDark,
            font: TOKENS.font,
            size: TOKENS.size.caption,
          }),
          new TextRun({
            text: "\t",
            font: TOKENS.font,
            size: TOKENS.size.caption,
          }),
          new TextRun({
            text: rightText,
            color: opts.textColor || TOKENS.color.accentDark,
            font: TOKENS.font,
            size: TOKENS.size.caption,
          }),
        ],
      }),
    ],
  });

const createFooter = (opts = {}) => {
  const { locale = "en", leftText } = opts;
  const t = (text) =>
    new TextRun({
      text,
      color: TOKENS.color.muted,
      font: TOKENS.font,
      size: TOKENS.size.caption,
    });
  return new Footer({
    children: [
      new Paragraph({
        alignment: leftText ? AlignmentType.LEFT : AlignmentType.CENTER,
        border: {
          top: {
            style: BorderStyle.SINGLE,
            size: 4,
            color: TOKENS.color.border,
            space: 1,
          },
        },
        children: [
          ...(leftText ? [t(leftText), t(" ")] : []),
          ...(locale === "bn"
            ? [
                t("পৃষ্ঠা "),
                new TextRun({
                  children: [PageNumber.CURRENT],
                  color: TOKENS.color.muted,
                  font: TOKENS.font,
                  size: TOKENS.size.caption,
                }),
                t(" এর "),
                new TextRun({
                  children: [PageNumber.TOTAL_PAGES],
                  color: TOKENS.color.muted,
                  font: TOKENS.font,
                  size: TOKENS.size.caption,
                }),
              ]
            : [
                t("Page "),
                new TextRun({
                  children: [PageNumber.CURRENT],
                  color: TOKENS.color.muted,
                  font: TOKENS.font,
                  size: TOKENS.size.caption,
                }),
                t(" of "),
                new TextRun({
                  children: [PageNumber.TOTAL_PAGES],
                  color: TOKENS.color.muted,
                  font: TOKENS.font,
                  size: TOKENS.size.caption,
                }),
              ]),
        ],
      }),
    ],
  });
};

const createMetadataTable = (rows, variant = "navyLabel", labelWidthPct = 35) =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders:
      variant === "navyLabel"
        ? gridBorders(AUDIT_THEME.color.grid, 4)
        : NO_BORDERS,
    rows: rows.map((row, index) => {
      const labelFill =
        variant === "navyLabel"
          ? AUDIT_THEME.color.labelFill
          : TOKENS.color.zebraA;
      const labelColor =
        variant === "navyLabel"
          ? AUDIT_THEME.color.labelText
          : TOKENS.color.ink;
      const valueFill =
        variant === "navyLabel"
          ? index % 2 === 0
            ? AUDIT_THEME.color.zebra
            : TOKENS.color.white
          : TOKENS.color.white;
      return new TableRow({
        children: [
          new TableCell({
            width: { size: labelWidthPct, type: WidthType.PERCENTAGE },
            shading: { fill: labelFill, type: ShadingType.CLEAR },
            margins: FORMAL_CELL_MARGINS,
            children: [
              new Paragraph({
                spacing: { after: 40 },
                children: [
                  new TextRun({
                    text: row.label,
                    bold: true,
                    font: TOKENS.font,
                    size: TOKENS.size.body,
                    color: labelColor,
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: { size: 100 - labelWidthPct, type: WidthType.PERCENTAGE },
            shading: { fill: valueFill, type: ShadingType.CLEAR },
            margins: FORMAL_CELL_MARGINS,
            children: [
              new Paragraph({
                spacing: { after: 40 },
                children: [
                  new TextRun({
                    text: row.value,
                    font: TOKENS.font,
                    size: TOKENS.size.body,
                    color: TOKENS.color.ink,
                  }),
                ],
              }),
            ],
          }),
        ],
      });
    }),
  });

const createNumberedSectionHeading = (prefix, title, opts = {}) =>
  new Paragraph({
    spacing: { before: 220, after: 120 },
    border: {
      bottom: {
        color: opts.ruleColor || AUDIT_THEME.color.sectionRule,
        style: BorderStyle.SINGLE,
        size: 6,
      },
    },
    children: [
      new TextRun({
        text: `${prefix} ${title}`,
        bold: true,
        font: opts.font || TOKENS.font,
        size: TOKENS.size.section,
        color: opts.color || AUDIT_THEME.color.heading,
      }),
    ],
  });

const createCenteredHeading = (text, opts = {}) =>
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 180 },
    children: [
      new TextRun({
        text,
        bold: true,
        font: TOKENS.font,
        size: TOKENS.size.section,
        color: opts.color || AUDIT_THEME.color.summaryHeading,
      }),
    ],
  });

async function loadDocumentImage(image) {
  try {
    // Local upload (serverPath)
    if (
      image &&
      image.type === "upload" &&
      image.serverPath &&
      fs.existsSync(image.serverPath)
    ) {
      return fs.readFileSync(image.serverPath);
    }

    // Data URI (base64 embedded)
    if (image && image.url && /^data:image\//i.test(image.url)) {
      const m = image.url.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/i);
      if (m && m[2]) return Buffer.from(m[2], "base64");
    }

    // Remote http(s) URL
    if (image && image.url && /^https?:\/\//i.test(image.url)) {
      const response = await fetch(image.url);
      if (response && response.ok)
        return Buffer.from(await response.arrayBuffer());
    }

    // file:// local path support
    if (image && image.url && image.url.startsWith("file://")) {
      const localPath = image.url.replace(/^file:\/\//i, "");
      if (fs.existsSync(localPath)) return fs.readFileSync(localPath);
    }
  } catch (error) {
    console.error(
      "Failed to load document image:",
      error && error.message ? error.message : error,
    );
  }
  return null;
}

// ----------------------------------------------------------------------------
// LAYOUT: Formal Letter / Application (দরখাস্ত)
// ----------------------------------------------------------------------------
async function buildLetterDocx(plan, images) {
  const lm = plan.letterMeta || {};
  const hasBengali = /[\u0980-\u09FF]/.test(
    `${lm.date || ""}${lm.subject || ""}${lm.signatory || ""}${(lm.recipientLines || []).join(" ")}`,
  );
  const locale = hasBengali ? "bn" : "en";

  // Resolve font
  if (hasBengali) TOKENS.font = "Nikosh";

  const footer = createFooter({ locale });

  const children = [];

  // тФАтФА Date тФАтФА
  if (lm.date) {
    children.push(
      new Paragraph({
        spacing: { after: 300 },
        children: [docText(lm.date, { bold: true, size: 22 })],
      }),
    );
  }

  // тФАтФА Recipient Block тФАтФА
  if (Array.isArray(lm.recipientLines) && lm.recipientLines.length) {
    lm.recipientLines.forEach((line) => {
      children.push(
        new Paragraph({
          spacing: { after: 40 },
          children: [docText(line, { size: 21 })],
        }),
      );
    });
    children.push(new Paragraph({ spacing: { after: 200 } }));
  }

  // тФАтФА Subject Line тФАтФА
  if (lm.subject) {
    const subjectLabel = hasBengali ? "বিষয়:" : "Subject:";
    children.push(
      new Paragraph({
        spacing: { before: 100, after: 200 },
        children: [
          docText(`${subjectLabel} `, {
            bold: true,
            size: 28,
            color: "1A5276",
          }),
          docText(lm.subject, {
            italics: true,
            size: 28,
            color: "1A5276",
          }),
        ],
      }),
    );
  }

  // тФАтФА Salutation тФАтФА
  if (lm.salutation) {
    children.push(
      new Paragraph({
        spacing: { after: 200 },
        children: [docText(lm.salutation, { size: 21 })],
      }),
    );
  }

  // тФАтФА Body Paragraphs тФАтФА
  const bodyParagraphs =
    plan.sections &&
    plan.sections[0] &&
    Array.isArray(plan.sections[0].paragraphs)
      ? plan.sections[0].paragraphs
      : [];
  bodyParagraphs.forEach((p) => {
    children.push(
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200, line: 300 },
        children: [docText(p, { size: 21 })],
      }),
    );
  });

  // тФАтФА Closing тФАтФА
  if (lm.closing) {
    children.push(new Paragraph({ spacing: { after: 80 } }));
    children.push(
      new Paragraph({
        spacing: { after: 400 },
        children: [docText(lm.closing, { size: 21 })],
      }),
    );
  }

  // тФАтФА Signature Name (large, stylized) тФАтФА
  if (lm.signatory) {
    children.push(
      new Paragraph({
        spacing: { after: 60 },
        children: [
          docText(lm.signatory, {
            bold: true,
            size: 40,
            color: "2C3E50",
            font: hasBengali ? "Nikosh" : "Georgia",
          }),
        ],
      }),
    );
  }

  // тФАтФА Signatory Title тФАтФА
  if (lm.signatoryTitle) {
    children.push(
      new Paragraph({
        spacing: { after: 30 },
        children: [
          docText(lm.signatoryTitle, { size: 19, color: TOKENS.color.subink }),
        ],
      }),
    );
  }

  // тФАтФА Signatory Organization тФАтФА
  if (lm.signatoryOrg) {
    children.push(
      new Paragraph({
        spacing: { after: 30 },
        children: [
          docText(lm.signatoryOrg, { size: 19, color: TOKENS.color.subink }),
        ],
      }),
    );
  }

  const document = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: { top: 1200, right: 1200, bottom: 1000, left: 1200 },
          },
        },
        footers: {
          default: footer,
        },
        children,
      },
    ],
  });
  return Packer.toBuffer(document);
}

// ----------------------------------------------------------------------------
// LAYOUT: Professional CV
// ----------------------------------------------------------------------------
async function buildCVDocx(plan, images) {
  const cv = plan.cvMeta || {};
  const imageData = await Promise.all(images.map(loadDocumentImage));
  const photoBuffer = imageData.length > 0 ? imageData[0] : null;

  const leftColChildren = [];
  const rightColChildren = [];

  // Helper for left column headers
  const leftHeader = (text) => {
    return new Paragraph({
      spacing: { before: 200, after: 100 },
      borders: {
        bottom: {
          color: "FFFFFF",
          space: 1,
          style: BorderStyle.SINGLE,
          size: 6,
        },
      },
      children: [
        docText(text.toUpperCase(), { bold: true, color: "FFFFFF", size: 24 }),
      ],
    });
  };

  // Helper for right column headers
  const rightHeader = (text) => {
    return new Paragraph({
      spacing: { before: 300, after: 100 },
      borders: {
        bottom: {
          color: "1B365D",
          space: 1,
          style: BorderStyle.SINGLE,
          size: 12,
        },
      },
      children: [
        docText(text.toUpperCase(), { bold: true, color: "1B365D", size: 28 }),
      ],
    });
  };

  // Left Column: Photo
  if (photoBuffer) {
    leftColChildren.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 400 },
        children: [
          new ImageRun({
            data: photoBuffer,
            transformation: { width: 150, height: 150 },
          }),
        ],
      }),
    );
  } else {
    leftColChildren.push(new Paragraph({ spacing: { before: 800 } }));
  }

  // Left Column: Contact
  if (Array.isArray(cv.contact) && cv.contact.length > 0) {
    leftColChildren.push(leftHeader("CONTACT"));
    cv.contact.forEach((c) => {
      leftColChildren.push(
        new Paragraph({
          spacing: { after: 100 },
          children: [
            docText(`${c.icon ? c.icon + " " : ""}${c.value}`, {
              color: "FFFFFF",
              size: 18,
            }),
          ],
        }),
      );
    });
  }

  // Left Column: Education
  if (Array.isArray(cv.education) && cv.education.length > 0) {
    leftColChildren.push(leftHeader("EDUCATION"));
    cv.education.forEach((e) => {
      leftColChildren.push(
        new Paragraph({
          spacing: { before: 100 },
          children: [
            docText(e.year, { color: "FFFFFF", size: 18, bold: true }),
          ],
        }),
        new Paragraph({
          spacing: { before: 50 },
          children: [
            docText(e.institution, { color: "FFFFFF", size: 18, bold: true }),
          ],
        }),
        new Paragraph({
          spacing: { before: 50, after: 100 },
          children: [docText(`тАв ${e.degree}`, { color: "FFFFFF", size: 18 })],
        }),
      );
    });
  }

  // Left Column: Skills
  if (Array.isArray(cv.skills) && cv.skills.length > 0) {
    leftColChildren.push(leftHeader("SKILLS"));
    cv.skills.forEach((s) => {
      leftColChildren.push(
        new Paragraph({
          spacing: { after: 100 },
          children: [docText(`тАв ${s}`, { color: "FFFFFF", size: 18 })],
        }),
      );
    });
  }

  // Left Column: Languages
  if (Array.isArray(cv.languages) && cv.languages.length > 0) {
    leftColChildren.push(leftHeader("LANGUAGES"));
    cv.languages.forEach((l) => {
      leftColChildren.push(
        new Paragraph({
          spacing: { after: 100 },
          children: [docText(`тАв ${l}`, { color: "FFFFFF", size: 18 })],
        }),
      );
    });
  }

  // Right Column: Header (Name & Title)
  rightColChildren.push(
    new Paragraph({
      spacing: { before: 400, after: 100 },
      children: [
        docText(cv.name || "NAME", { bold: true, size: 56, color: "2C3E50" }),
      ],
    }),
    new Paragraph({
      spacing: { after: 400 },
      children: [docText(cv.title || "TITLE", { size: 28, color: "7F8C8D" })],
    }),
  );

  // Right Column: Profile
  if (cv.profile) {
    rightColChildren.push(rightHeader("PROFILE"));
    rightColChildren.push(
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200, line: 300 },
        children: [docText(cv.profile, { size: 20 })],
      }),
    );
  }

  // Right Column: Experience
  if (Array.isArray(cv.experience) && cv.experience.length > 0) {
    rightColChildren.push(rightHeader("WORK EXPERIENCE"));
    cv.experience.forEach((exp) => {
      rightColChildren.push(
        new Paragraph({
          spacing: { before: 200, after: 50 },
          tabStops: [{ type: TabStopType.RIGHT, position: 9000 }],
          children: [
            docText(exp.title, { bold: true, size: 22 }),
            new TextRun({ text: "\t" }), // Tab
            docText(exp.date, { size: 18, color: "7F8C8D" }),
          ],
        }),
        new Paragraph({
          spacing: { after: 100 },
          children: [
            docText(exp.company, { size: 20, color: "163A5F", italics: true }),
          ],
        }),
      );
      if (Array.isArray(exp.bullets)) {
        exp.bullets.forEach((b) => {
          rightColChildren.push(
            new Paragraph({
              spacing: { after: 100 },
              children: [docText(`тАв ${b}`, { size: 18 })],
            }),
          );
        });
      }
    });
  }

  // Right Column: Projects
  if (Array.isArray(cv.projects) && cv.projects.length > 0) {
    rightColChildren.push(rightHeader("PROJECTS"));
    cv.projects.forEach((p) => {
      rightColChildren.push(
        new Paragraph({
          spacing: { after: 100 },
          children: [
            docText(`тАв ${p.name}`, { bold: true, size: 20 }),
            docText(` (${p.description})`, { size: 18, color: "7F8C8D" }),
            docText(p.link ? ` -> ${p.link}` : "", {
              size: 18,
              italics: true,
              color: "163A5F",
            }),
          ],
        }),
      );
    });
  }

  // Build the main table
  const mainTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 35, type: WidthType.PERCENTAGE },
            shading: { fill: "1B365D", type: ShadingType.CLEAR },
            margins: { top: 400, bottom: 400, left: 400, right: 400 },
            children: leftColChildren,
            borders: {
              top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            },
          }),
          new TableCell({
            width: { size: 65, type: WidthType.PERCENTAGE },
            margins: { top: 400, bottom: 400, left: 600, right: 600 },
            children: rightColChildren,
            borders: {
              top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            },
          }),
        ],
      }),
    ],
  });

  const document = new Document({
    sections: [
      {
        properties: {
          page: { margin: { top: 0, right: 0, bottom: 0, left: 0 } },
        },
        children: [mainTable],
      },
    ],
  });
  return Packer.toBuffer(document);
}

async function buildCVDocx(plan, images) {
  const cv = plan.cvMeta || {};
  const hasBengali = /[\u0980-\u09FF]/.test(JSON.stringify(cv));
  if (hasBengali) TOKENS.font = "Nikosh";

  const children = [];
  const leftColChildren = [];

  // Left Column Styling Helpers
  const leftHeading = (text) =>
    new Paragraph({
      spacing: { before: 240, after: 120 },
      border: {
        bottom: {
          color: "FFFFFF",
          space: 5,
          style: BorderStyle.SINGLE,
          size: 6,
        },
      },
      children: [docText(text, { bold: true, size: 24, color: "FFFFFF" })],
    });

  const leftItem = (text, isBold = false) =>
    new Paragraph({
      spacing: { after: 60 },
      children: [docText(text, { size: 20, color: "FFFFFF", bold: isBold })],
    });

  const leftBullet = (text) =>
    new Paragraph({
      spacing: { after: 60 },
      indent: { left: 200, hanging: 100 },
      children: [docText(`тАв ${text}`, { size: 18, color: "FFFFFF" })],
    });

  // Photo
  if (images && images.length > 0) {
    const imageData = await loadDocumentImage(images[0]);
    if (imageData && Buffer.isBuffer(imageData) && imageData.length > 0) {
      leftColChildren.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 300 },
          children: [
            new ImageRun({
              data: imageData,
              transformation: { width: 150, height: 150 },
              type: "jpeg",
            }),
          ],
        }),
      );
    }
  }

  // Contact
  if (cv.contact && cv.contact.length) {
    leftColChildren.push(
      leftHeading(hasBengali ? "যোগাযোগ" : "CONTACT"),
    );
    cv.contact.forEach((c) => {
      leftColChildren.push(leftItem(c.value));
    });
  }

  // Education (Moved to Left)
  if (cv.education && cv.education.length) {
    leftColChildren.push(
      leftHeading(hasBengali ? "শিক্ষা" : "EDUCATION"),
    );
    cv.education.forEach((edu) => {
      leftColChildren.push(leftItem(edu.year || "", true));
      leftColChildren.push(leftItem(edu.institution, true));
      leftColChildren.push(leftBullet(edu.degree));
      leftColChildren.push(new Paragraph({ spacing: { after: 100 } }));
    });
  }

  // Skills
  if (cv.skills && cv.skills.length) {
    leftColChildren.push(
      leftHeading(hasBengali ? "দক্ষতা" : "SKILLS"),
    );
    cv.skills.forEach((s) => {
      leftColChildren.push(leftBullet(s));
    });
  }

  // Languages
  if (cv.languages && cv.languages.length) {
    leftColChildren.push(
      leftHeading(hasBengali ? "ভাষা" : "LANGUAGES"),
    );
    cv.languages.forEach((l) => {
      leftColChildren.push(leftBullet(l));
    });
  }

  // Right Column Content
  const rightColChildren = [];

  const rightHeading = (text) =>
    new Paragraph({
      spacing: { before: 240, after: 120 },
      border: {
        bottom: {
          color: "1E3A5F",
          space: 5,
          style: BorderStyle.SINGLE,
          size: 12,
        },
      },
      children: [docText(text, { bold: true, size: 28, color: "1E3A5F" })],
    });

  // Name & Title
  if (cv.name) {
    rightColChildren.push(
      new Paragraph({
        spacing: { after: 50 },
        children: [
          docText(cv.name.toUpperCase(), {
            bold: true,
            size: 48,
            color: "2C3E50",
          }),
        ],
      }),
    );
  }
  if (cv.title) {
    rightColChildren.push(
      new Paragraph({
        spacing: { after: 300 },
        border: {
          bottom: {
            color: "2C3E50",
            space: 5,
            style: BorderStyle.SINGLE,
            size: 6,
          },
        },
        children: [
          docText(cv.title.toUpperCase(), { size: 24, color: "2C3E50" }),
        ],
      }),
    );
  }

  // Profile
  if (cv.profile) {
    rightColChildren.push(
      rightHeading(hasBengali ? "প্রোফাইল" : "PROFILE"),
    );
    rightColChildren.push(
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [docText(cv.profile, { size: 20 })],
      }),
    );
  }

  // Experience
  if (cv.experience && cv.experience.length) {
    rightColChildren.push(
      rightHeading(hasBengali ? "অভিজ্ঞতা" : "WORK EXPERIENCE"),
    );
    cv.experience.forEach((exp) => {
      rightColChildren.push(
        new Paragraph({
          spacing: { after: 50 },
          tabStops: [{ type: TabStopType.RIGHT, position: 7500 }],
          children: [
            docText(exp.title, { bold: true, size: 24, color: "2C3E50" }),
            new TextRun({
              text: `\t${exp.date || ""}`,
              size: 20,
              color: "2C3E50",
            }),
          ],
        }),
      );
      if (exp.company) {
        rightColChildren.push(
          new Paragraph({
            spacing: { after: 80 },
            children: [
              docText(exp.company, {
                size: 22,
                color: "2980B9",
                underline: {},
              }),
            ],
          }),
        );
      }
      if (exp.bullets && exp.bullets.length) {
        exp.bullets.forEach((b) => {
          rightColChildren.push(
            new Paragraph({
              spacing: { after: 60 },
              indent: { left: 200, hanging: 100 },
              alignment: AlignmentType.JUSTIFIED,
              children: [docText(`тАв ${b}`, { size: 20 })],
            }),
          );
        });
      }
      rightColChildren.push(new Paragraph({ spacing: { after: 200 } }));
    });
  }

  // Projects
  if (cv.projects && cv.projects.length) {
    rightColChildren.push(
      rightHeading(
        hasBengali ? "প্রজেক্টসমূহ" : "PROJECTS",
      ),
    );
    cv.projects.forEach((proj) => {
      rightColChildren.push(
        new Paragraph({
          spacing: { after: 50 },
          tabStops: [{ type: TabStopType.RIGHT, position: 7500 }],
          children: [
            docText(proj.name, { bold: true, size: 22, color: "2C3E50" }),
            proj.link
              ? new TextRun({
                  text: `\t-> ${proj.link}`,
                  size: 18,
                  color: "2980B9",
                  underline: {},
                })
              : new TextRun({ text: "" }),
          ],
        }),
      );
      if (proj.description) {
        rightColChildren.push(
          new Paragraph({
            spacing: { after: 100 },
            indent: { left: 200 },
            children: [docText(`тАв ${proj.description}`, { size: 20 })],
          }),
        );
      }
      rightColChildren.push(new Paragraph({ spacing: { after: 100 } }));
    });
  }

  if (!leftColChildren.length) leftColChildren.push(new Paragraph(""));
  if (!rightColChildren.length) rightColChildren.push(new Paragraph(""));

  const layoutTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: NO_BORDERS,
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: { fill: "1E3A5F", type: ShadingType.CLEAR },
            margins: { top: 400, bottom: 400, left: 300, right: 300 },
            children: leftColChildren,
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            margins: { top: 400, bottom: 400, left: 400, right: 300 },
            children: rightColChildren,
          }),
        ],
      }),
    ],
  });

  children.push(layoutTable);

  const document = new Document({
    sections: [
      {
        properties: {
          page: { margin: { top: 0, right: 0, bottom: 0, left: 0 } },
        },
        children,
      },
    ],
  });
  return Packer.toBuffer(document);
}

const PROPOSAL_THEME = {
  font: "Arial",
  color: {
    header: "5F6B7A",
    heading: "3F4F4F",
    accent: "C2B5A3",
    ink: "333333",
    subink: "666666",
    rule: "E5E7EB",
  },
};

const createProposalHeader = (leftText, rightText) =>
  new docx.Header({
    children: [
      new docx.Paragraph({
        tabStops: [
          { type: docx.TabStopType.RIGHT, position: docx.TabStopPosition.MAX },
        ],
        spacing: { after: 400 },
        children: [
          docText(leftText, {
            color: PROPOSAL_THEME.color.header,
            font: PROPOSAL_THEME.font,
          }),
          docText("\t", { font: PROPOSAL_THEME.font }),
          docText(rightText, {
            color: PROPOSAL_THEME.color.header,
            font: PROPOSAL_THEME.font,
          }),
        ],
      }),
    ],
  });

const createProposalCover = (pm) => [
  new docx.Paragraph({ spacing: { before: 1600 }, children: [docText("")] }),
  new docx.Paragraph({
    spacing: { after: 100 },
    children: [
      docText(pm.title || "Project Proposal тЖЧ", {
        bold: true,
        size: 120,
        color: PROPOSAL_THEME.color.heading,
        font: PROPOSAL_THEME.font,
      }),
    ],
  }),
  new docx.Paragraph({
    spacing: { after: 800 },
    children: [
      docText(pm.subtitle || "Digital Marketing Transformation Initiative", {
        size: 52,
        color: PROPOSAL_THEME.color.accent,
        font: PROPOSAL_THEME.font,
      }),
    ],
  }),
  new docx.Table({
    width: { size: 100, type: docx.WidthType.PERCENTAGE },
    borders: {
      top: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      insideVertical: {
        style: docx.BorderStyle.NONE,
        size: 0,
        color: "FFFFFF",
      },
    },
    rows: [
      new docx.TableRow({
        children: [
          new docx.TableCell({
            width: { size: 50, type: docx.WidthType.PERCENTAGE },
            shading: {
              fill: PROPOSAL_THEME.color.heading,
              type: docx.ShadingType.CLEAR,
            },
            margins: { top: 600, bottom: 600, left: 600, right: 600 },
            children: [
              new docx.Paragraph({
                spacing: { after: 160 },
                children: [
                  docText("Prepared by :", {
                    color: "FFFFFF",
                    size: 28,
                    font: PROPOSAL_THEME.font,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  docText(pm.preparedBy || "Avery Davis", {
                    bold: true,
                    color: "FFFFFF",
                    size: 40,
                    font: PROPOSAL_THEME.font,
                  }),
                ],
              }),
            ],
          }),
          new docx.TableCell({
            width: { size: 50, type: docx.WidthType.PERCENTAGE },
            shading: {
              fill: PROPOSAL_THEME.color.heading,
              type: docx.ShadingType.CLEAR,
            },
            margins: { top: 600, bottom: 600, left: 600, right: 600 },
            children: [
              new docx.Paragraph({
                spacing: { after: 160 },
                children: [
                  docText("Prepared for :", {
                    color: "FFFFFF",
                    size: 28,
                    font: PROPOSAL_THEME.font,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  docText(pm.preparedFor || "Wardiere. Inc", {
                    bold: true,
                    color: "FFFFFF",
                    size: 40,
                    font: PROPOSAL_THEME.font,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  }),
];

const createProposalTOC = (sections) => {
  const children = [];
  children.push(
    new docx.Paragraph({
      spacing: { before: 800, after: 1200 },
      children: [
        docText("Table of Contents", {
          bold: true,
          size: 80,
          color: "FFFFFF",
          font: PROPOSAL_THEME.font,
        }),
      ],
    }),
  );

  const headingSections = sections.filter((sec) => sec.heading);
  headingSections.forEach((sec, idx) => {
    const numStr = String(idx + 1).padStart(2, "0");
    const isLast = idx === headingSections.length - 1;
    children.push(
      new docx.Paragraph({
        spacing: { before: 300, after: 300 },
        tabStops: [{ type: docx.TabStopType.LEFT, position: 1000 }],
        ...(isLast
          ? {}
          : {
              border: {
                bottom: {
                  color: "FFFFFF",
                  style: docx.BorderStyle.SINGLE,
                  size: 6,
                },
              },
            }),
        children: [
          docText(numStr, {
            size: 40,
            color: PROPOSAL_THEME.color.accent,
            font: PROPOSAL_THEME.font,
          }),
          docText("\t", { font: PROPOSAL_THEME.font }),
          docText(sec.heading, {
            size: 40,
            color: "FFFFFF",
            font: PROPOSAL_THEME.font,
          }),
        ],
      }),
    );
  });

  return new docx.Table({
    width: { size: 100, type: docx.WidthType.PERCENTAGE },
    borders: {
      top: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
    },
    rows: [
      new docx.TableRow({
        children: [
          new docx.TableCell({
            shading: {
              fill: PROPOSAL_THEME.color.heading,
              type: docx.ShadingType.CLEAR,
            },
            margins: { top: 1200, bottom: 4000, left: 600, right: 600 },
            children,
          }),
        ],
      }),
    ],
  });
};

const createBigNumberList = (items) =>
  items.flatMap((item) => [
    new docx.Paragraph({
      spacing: { before: 400, after: 80 },
      children: [
        docText(item.number || "", {
          size: 90,
          color: PROPOSAL_THEME.color.accent,
          font: PROPOSAL_THEME.font,
        }),
      ],
    }),
    new docx.Paragraph({
      spacing: { after: 160 },
      children: [
        docText(item.title || "", {
          bold: true,
          size: 44,
          color: PROPOSAL_THEME.color.heading,
          font: PROPOSAL_THEME.font,
        }),
      ],
    }),
    new docx.Paragraph({
      spacing: { after: 200 },
      children: [
        docText(item.text || "", {
          size: 28,
          color: PROPOSAL_THEME.color.subink,
          font: PROPOSAL_THEME.font,
        }),
      ],
    }),
  ]);

const createDividedList = (items) =>
  items.flatMap((item) => [
    new docx.Paragraph({
      spacing: { before: 300, after: 120 },
      children: [
        docText("тЖТ  ", {
          bold: true,
          size: 36,
          color: PROPOSAL_THEME.color.heading,
          font: PROPOSAL_THEME.font,
        }),
        docText(item.title || "", {
          bold: true,
          size: 36,
          color: PROPOSAL_THEME.color.heading,
          font: PROPOSAL_THEME.font,
        }),
      ],
    }),
    new docx.Paragraph({
      spacing: { after: 200 },
      indent: { left: 600 },
      children: [
        docText(item.subtitle || "", {
          size: 28,
          color: PROPOSAL_THEME.color.subink,
          font: PROPOSAL_THEME.font,
        }),
      ],
    }),
    new docx.Paragraph({
      border: {
        bottom: {
          color: PROPOSAL_THEME.color.rule,
          style: docx.BorderStyle.SINGLE,
          size: 6,
        },
      },
      children: [docText("")],
    }),
  ]);

async function buildProposalDocx(plan, images) {
  const pm = plan.proposalMeta || {};

  const imageData = await Promise.all(images.map(loadDocumentImage));
  const placedAfter = new Map();
  (plan.imagePlacements || []).forEach((p) => {
    const key = Number(p.afterSection) || 0;
    if (!placedAfter.has(key)) placedAfter.set(key, []);
    placedAfter.get(key).push(p);
  });

  const makeImageParagraphs = (placement) => {
    const idx = Number(placement.imageIndex) - 1;
    const data = imageData[idx];
    if (!data || !Buffer.isBuffer(data) || data.length === 0) return [];
    const mime =
      data.length > 12 && data[0] === 0xff && data[1] === 0xd8
        ? "image/jpeg"
        : "image/png";
    return [
      new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        spacing: { before: 200, after: 200 },
        children: [
          new docx.ImageRun({
            data,
            transformation: { width: 600, height: 300 },
            type: mime,
          }),
        ],
      }),
    ];
  };

  const children = [];

  children.push(...createProposalCover(pm));
  children.push(new docx.Paragraph({ children: [new docx.PageBreak()] }));

  children.push(createProposalTOC(plan.sections || []));
  children.push(new docx.Paragraph({ children: [new docx.PageBreak()] }));

  if (placedAfter.has(0)) {
    placedAfter.get(0).forEach((pl) => {
      children.push(...makeImageParagraphs(pl));
    });
  }

  (plan.sections || []).forEach((sec, idx) => {
    const secId = idx + 1;

    if (sec.heading) {
      children.push(
        new docx.Paragraph({
          spacing: { before: 600, after: 300 },
          children: [
            docText(sec.heading, {
              bold: true,
              size: 72,
              color: PROPOSAL_THEME.color.heading,
              font: PROPOSAL_THEME.font,
            }),
          ],
        }),
      );
    }

    if (sec.paragraphs) {
      sec.paragraphs.forEach((p) => {
        children.push(
          new docx.Paragraph({
            spacing: { after: 200 },
            children: [
              docText(p, {
                size: 26,
                color: "000000",
                font: PROPOSAL_THEME.font,
              }),
            ],
          }),
        );
      });
    }

    if (sec.bigNumberList) {
      children.push(...createBigNumberList(sec.bigNumberList));
    }

    if (sec.dividedList) {
      children.push(...createDividedList(sec.dividedList));
    }

    const table = sec.table;
    if (
      table &&
      Array.isArray(table.headers) &&
      table.headers.length &&
      Array.isArray(table.rows)
    ) {
      children.push(
        new docx.Table({
          width: { size: 100, type: docx.WidthType.PERCENTAGE },
          layout: docx.TableLayoutType.FIXED,
          borders: {
            top: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
            bottom: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
            left: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
            right: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
            insideVertical: {
              style: docx.BorderStyle.NONE,
              size: 0,
              color: "FFFFFF",
            },
            insideHorizontal: {
              style: docx.BorderStyle.SINGLE,
              size: 6,
              color: PROPOSAL_THEME.color.rule,
            },
          },
          rows: [
            new docx.TableRow({
              tableHeader: true,
              children: table.headers.map(
                (h) =>
                  new docx.TableCell({
                    shading: { fill: "FFFFFF", type: docx.ShadingType.CLEAR },
                    margins: { top: 200, bottom: 200, left: 120, right: 120 },
                    borders: {
                      bottom: {
                        style: docx.BorderStyle.SINGLE,
                        size: 12,
                        color: PROPOSAL_THEME.color.heading,
                      },
                    },
                    children: [
                      new docx.Paragraph({
                        children: [
                          docText(h, {
                            bold: true,
                            color: PROPOSAL_THEME.color.accent,
                            size: 28,
                            font: PROPOSAL_THEME.font,
                          }),
                        ],
                      }),
                    ],
                  }),
              ),
            }),
            ...table.rows.map(
              (row, r) =>
                new docx.TableRow({
                  children: table.headers.map(
                    (_, c) =>
                      new docx.TableCell({
                        shading: {
                          fill: "FFFFFF",
                          type: docx.ShadingType.CLEAR,
                        },
                        margins: {
                          top: 180,
                          bottom: 180,
                          left: 120,
                          right: 120,
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              docText(Array.isArray(row) ? row[c] || "" : "", {
                                size: 24,
                                color: "000000",
                                font: PROPOSAL_THEME.font,
                              }),
                            ],
                          }),
                        ],
                      }),
                  ),
                }),
            ),
          ],
        }),
      );
    }

    if (placedAfter.has(secId)) {
      placedAfter.get(secId).forEach((pl) => {
        children.push(...makeImageParagraphs(pl));
      });
    }
  });

  const document = new docx.Document({
    creator: "Antigravity AI",
    sections: [
      {
        properties: {},
        headers: {
          default: createProposalHeader(
            pm.preparedBy || "Liceria Corporate",
            pm.date || "June 2030",
          ),
        },
        children,
      },
    ],
  });

  return docx.Packer.toBuffer(document);
}

async function buildDocxBuffer(plan, images) {
  const imageData = await Promise.all(images.map(loadDocumentImage));
  const placedAfter = new Map();
  plan.imagePlacements.forEach((p) => {
    const key = Number(p.afterSection) || 0;
    if (!placedAfter.has(key)) placedAfter.set(key, []);
    placedAfter.get(key).push(p);
  });

  const makeImageParagraphs = (placement) => {
    const idx = Number(placement.imageIndex) - 1;
    const data = imageData[idx];
    if (!data || !Buffer.isBuffer(data) || data.length === 0) {
      console.warn(
        `[docx] Skipping image #${placement.imageIndex} (index ${idx}): no data or invalid buffer`,
      );
      return [];
    }

    const inferMime = (buf) => {
      if (!buf || buf.length < 12) return null;
      const b0 = buf[0],
        b1 = buf[1],
        b2 = buf[2],
        b3 = buf[3];
      if (b0 === 0xff && b1 === 0xd8) return "image/jpeg";
      if (b0 === 0x89 && b1 === 0x50 && b2 === 0x4e && b3 === 0x47)
        return "image/png";
      if (b0 === 0x47 && b1 === 0x49 && b2 === 0x46) return "image/gif";
      if (b0 === 0x42 && b1 === 0x4d) return "image/bmp";
      if (b0 === 0x52 && b1 === 0x49 && b2 === 0x46 && b3 === 0x46) {
        try {
          const tag = buf.slice(8, 12).toString("ascii");
          if (tag === "WEBP") return "image/webp";
        } catch (e) {
          /* ignore */
        }
      }
      try {
        const start = buf.slice(0, 512).toString("utf8");
        if (start.indexOf("<svg") !== -1) return "image/svg+xml";
      } catch (e) {
        /* ignore decode errors */
      }
      return null;
    };

    const mime = inferMime(data);
    if (!mime) {
      console.warn(
        `[docx] Skipping image #${placement.imageIndex}: unknown mime. firstBytes=${data.slice(0, 12).toString("hex")}`,
      );
      return [];
    }
    const dataUri = `data:${mime};base64,${data.toString("base64")}`;

    const captionText =
      placement.caption ||
      images[Number(placement.imageIndex) - 1]?.description ||
      "";
    const captionParagraph = captionText
      ? [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 180 },
            children: [
              docText(captionText, {
                size: 17,
                italics: true,
                color: "64748B",
              }),
            ],
          }),
        ]
      : [];

    switch (placement.placement) {
      case "banner":
        return [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 180, after: 80 },
            children: [
              new ImageRun({
                data: dataUri,
                transformation: { width: 600, height: 240 },
              }),
            ],
          }),
          ...captionParagraph,
        ];
      case "left":
      case "right":
        return [
          new Paragraph({
            spacing: { before: 180, after: 120, line: 300 },
            children: [
              new ImageRun({
                data: dataUri,
                transformation: { width: 240, height: 180 },
                floating: {
                  horizontalPosition: {
                    relative: "column",
                    align: placement.placement === "left" ? "left" : "right",
                  },
                  verticalPosition: { relative: "paragraph", offset: 0 },
                  wrap: {
                    type: "square",
                    side: placement.placement === "left" ? "right" : "left",
                  },
                  margins: {
                    left: 100000,
                    right: 100000,
                    top: 0,
                    bottom: 100000,
                  },
                },
              }),
            ],
          }),
          ...captionParagraph,
        ];
      case "watermark":
        return [
          new Paragraph({
            children: [
              new ImageRun({
                data: dataUri,
                transformation: { width: 420, height: 420 },
                floating: {
                  horizontalPosition: { relative: "page", align: "center" },
                  verticalPosition: { relative: "page", align: "center" },
                  wrap: { type: "none" },
                  behindDocument: true,
                },
              }),
            ],
          }),
        ];
      case "header":
        return [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 120, after: 80 },
            children: [
              new ImageRun({
                data: dataUri,
                transformation: { width: 360, height: 120 },
              }),
            ],
          }),
          ...captionParagraph,
        ];
      case "grid":
        return [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 180, after: 80 },
            children: [
              new ImageRun({
                data: dataUri,
                transformation: { width: 480, height: 270 },
              }),
            ],
          }),
          ...captionParagraph,
        ];
      case "wide":
      default:
        return [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 180, after: 60 },
            children: [
              new ImageRun({
                data: dataUri,
                transformation: { width: 480, height: 270 },
              }),
            ],
          }),
          ...captionParagraph,
        ];
    }
  };

  const addImages = (items) =>
    (items || []).flatMap((placement) => makeImageParagraphs(placement));

  const parseMetadataRows = () => {
    const rows = [];
    (Array.isArray(plan.metadata) ? plan.metadata : []).forEach((entry) => {
      if (!entry || typeof entry !== "string") return;
      const parts = entry.split(/[:я╝Ъ]/).map((part) => part.trim());
      if (parts.length >= 2) {
        rows.push({ label: parts[0], value: parts.slice(1).join(": ") });
      }
    });
    return rows;
  };

  const metadataRows = parseMetadataRows();
  const dateRow = metadataRows.find((row) =>
    /তারিখ|সময়কাল|শুরু|শেষ|date|period|duration/i.test(
      row.label,
    ),
  );
  const hasBengaliText = /[\u0980-\u09FF]/.test(
    `${plan.title || ""}${plan.subtitle || ""}${plan.documentType || ""}${metadataRows
      .map((row) => `${row.label}${row.value}`)
      .join(" ")}`,
  );

  const header = dateRow
    ? createDualHeader(
        plan.title || plan.documentType || "AI Document",
        dateRow.value,
        {
          textColor: AUDIT_THEME.color.header,
          ruleColor: AUDIT_THEME.color.header,
        },
      )
    : new Header({
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [
              docText(plan.documentType || plan.title || "AI Document", {
                size: TOKENS.size.caption,
                color: TOKENS.color.muted,
              }),
            ],
          }),
        ],
      });

  const footer = createFooter({ locale: hasBengaliText ? "bn" : "en" });

  const children = [];
  if (metadataRows.length) {
    children.push(
      createCenteredHeading(
        "প্রতিবেদনের সারসংক্ষেপ",
        {
          color: AUDIT_THEME.color.summaryHeading,
        },
      ),
    );
    children.push(createMetadataTable(metadataRows, "navyLabel"));
    children.push(new Paragraph({ spacing: { after: 220 } }));
  } else {
    children.push(
      new Paragraph({
        spacing: { before: 120, after: plan.subtitle ? 80 : 300 },
        children: [
          docText(plan.title, { bold: true, size: 46, color: "163A5F" }),
        ],
      }),
    );
    if (plan.subtitle) {
      children.push(
        new Paragraph({
          spacing: { after: 300 },
          children: [docText(plan.subtitle, { size: 23, color: "64748B" })],
        }),
      );
    }
    children.push(...addImages(placedAfter.get(0)));
  }

  plan.sections.forEach((section, index) => {
    if (section.heading) {
      children.push(
        createNumberedSectionHeading(section.heading, "", {
          color: AUDIT_THEME.color.heading,
          ruleColor: AUDIT_THEME.color.sectionRule,
          font: TOKENS.font,
        }),
      );
    }

    (Array.isArray(section.paragraphs) ? section.paragraphs : []).forEach((p) =>
      children.push(
        new Paragraph({
          spacing: { after: 120, line: 300 },
          children: [docText(p)],
        }),
      ),
    );

    if (section.callout) {
      children.push(
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  shading: { fill: "EAF4FB", type: ShadingType.CLEAR },
                  margins: { top: 140, bottom: 140, left: 180, right: 180 },
                  children: [
                    new Paragraph({
                      children: [
                        docText(section.callout, {
                          bold: true,
                          color: "174C70",
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      );
    }

    (Array.isArray(section.bullets) ? section.bullets : []).forEach((item) =>
      children.push(
        new Paragraph({
          bullet: { level: 0 },
          spacing: { after: 65 },
          children: [docText(item)],
        }),
      ),
    );

    (Array.isArray(section.numbered) ? section.numbered : []).forEach((item) =>
      children.push(
        new Paragraph({
          numbering: { reference: "document-numbering", level: 0 },
          spacing: { after: 65 },
          children: [docText(item)],
        }),
      ),
    );

    if (Array.isArray(section.signatures) && section.signatures.length > 0) {
      children.push(
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          borders: {
            top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            insideHorizontal: {
              style: BorderStyle.NONE,
              size: 0,
              color: "FFFFFF",
            },
            insideVertical: {
              style: BorderStyle.NONE,
              size: 0,
              color: "FFFFFF",
            },
          },
          rows: section.signatures.map(
            (sig) =>
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 20, type: WidthType.PERCENTAGE },
                    margins: { top: 120, bottom: 120 },
                    borders: {
                      top: {
                        style: BorderStyle.NONE,
                        size: 0,
                        color: "FFFFFF",
                      },
                      bottom: {
                        style: BorderStyle.NONE,
                        size: 0,
                        color: "FFFFFF",
                      },
                      left: {
                        style: BorderStyle.NONE,
                        size: 0,
                        color: "FFFFFF",
                      },
                      right: {
                        style: BorderStyle.NONE,
                        size: 0,
                        color: "FFFFFF",
                      },
                    },
                    children: [
                      new Paragraph({
                        children: [
                          docText(sig.role || "", {
                            size: 21,
                            color: "000000",
                          }),
                        ],
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 45, type: WidthType.PERCENTAGE },
                    margins: { top: 120, bottom: 120 },
                    borders: {
                      top: {
                        style: BorderStyle.NONE,
                        size: 0,
                        color: "FFFFFF",
                      },
                      bottom: {
                        style: BorderStyle.NONE,
                        size: 0,
                        color: "FFFFFF",
                      },
                      left: {
                        style: BorderStyle.NONE,
                        size: 0,
                        color: "FFFFFF",
                      },
                      right: {
                        style: BorderStyle.NONE,
                        size: 0,
                        color: "FFFFFF",
                      },
                    },
                    children: [
                      new Paragraph({
                        children: [
                          docText(sig.name || "", {
                            size: 21,
                            color: "000000",
                          }),
                        ],
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 35, type: WidthType.PERCENTAGE },
                    margins: { top: 120, bottom: 120 },
                    borders: {
                      top: {
                        style: BorderStyle.NONE,
                        size: 0,
                        color: "FFFFFF",
                      },
                      bottom: {
                        style: BorderStyle.NONE,
                        size: 0,
                        color: "FFFFFF",
                      },
                      left: {
                        style: BorderStyle.NONE,
                        size: 0,
                        color: "FFFFFF",
                      },
                      right: {
                        style: BorderStyle.NONE,
                        size: 0,
                        color: "FFFFFF",
                      },
                    },
                    children: [
                      new Paragraph({
                        children: [
                          docText("______________________", {
                            size: 21,
                            color: "000000",
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
          ),
        }),
      );
    }

    const table = section.table;
    if (
      table &&
      Array.isArray(table.headers) &&
      table.headers.length &&
      Array.isArray(table.rows)
    ) {
      children.push(
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          rows: [
            new TableRow({
              tableHeader: true,
              children: table.headers.map(
                (h) =>
                  new TableCell({
                    shading: { fill: "163A5F", type: ShadingType.CLEAR },
                    margins: { top: 100, bottom: 100, left: 120, right: 120 },
                    children: [
                      new Paragraph({
                        children: [
                          docText(h, { bold: true, color: "FFFFFF", size: 19 }),
                        ],
                      }),
                    ],
                  }),
              ),
            }),
            ...table.rows.map(
              (row, r) =>
                new TableRow({
                  children: table.headers.map(
                    (_, c) =>
                      new TableCell({
                        shading: {
                          fill: r % 2 ? "FFFFFF" : "F5F9FC",
                          type: ShadingType.CLEAR,
                        },
                        margins: { top: 90, bottom: 90, left: 120, right: 120 },
                        children: [
                          new Paragraph({
                            children: [
                              docText(Array.isArray(row) ? row[c] || "" : "", {
                                size: 19,
                              }),
                            ],
                          }),
                        ],
                      }),
                  ),
                }),
            ),
          ],
        }),
      );
    }

    children.push(...addImages(placedAfter.get(index + 1)));
  });

  const document = new Document({
    numbering: {
      config: [
        {
          reference: "document-numbering",
          levels: [
            {
              level: 0,
              format: "decimal",
              text: "%1.",
              alignment: AlignmentType.START,
              style: { paragraph: { indent: { left: 720, hanging: 360 } } },
            },
          ],
        },
      ],
    },
    sections: [
      {
        properties: {
          page: { margin: { top: 1000, right: 1000, bottom: 900, left: 1000 } },
        },
        headers: {
          default: header,
        },
        footers: {
          default: footer,
        },
        children,
      },
    ],
  });
  return Packer.toBuffer(document);
}

// Catch unhandled errors so the process doesn't silently crash on shared hosting
process.on("uncaughtException", (err) => {
  console.error("Uncaught exception:", err);
});
process.on("unhandledRejection", (reason) => {
  console.error("Unhandled rejection:", reason);
});

// Start Server тАФ only call app.listen() when run directly via CLI.
// On cPanel / Phusion Passenger the host manages the port; Passenger
// imports this module and expects module.exports = app.
if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`ЁЯЪА SlideForge server running at http://localhost:${PORT}`);
    console.log(
      `ЁЯдЦ Configured to use AI API: ${process.env.OPENAI_BASE_URL} (Model: ${process.env.OPENAI_MODEL})`,
    );
  });
}

// Passenger requires `module.exports = app` (the raw Express instance),
// NOT an object like { app, ... }. Exporting an object causes a 503.
module.exports = app;
