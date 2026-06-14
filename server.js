// ===== ADCDA Build Service =====
// خدمة بناء العروض: تستقبل JSON المحتوى وترجع ملف PPTX بهوية ADCDA.
// + نقطة تأصيل: GET /topic/:code تُرجع نص الموضوع المعتمد ومراجعه من قاعدة المعرفة.
const express = require("express");
const { execFileSync } = require("child_process");
const fs = require("fs");
const path = require("path");
const os = require("os");

const app = express();
app.use(express.json({ limit: "5mb" }));

const PORT = process.env.PORT || 3000;
const BUILD = path.join(__dirname, "build.js");
const FIXRTL = path.join(__dirname, "fix_rtl.py");

// ===== قاعدة المعرفة (التأصيل) =====
// corpus.json: خريطة { "A.1": { topic_code, title, package, num_slides, full_text, refs:[...] }, ... }
// المصدر الوحيد المعتمد للمحتوى — النموذج لا يضيف من عنده، فقط يحوّل هذا النص لشرائح.
let CORPUS = {};
try { CORPUS = JSON.parse(fs.readFileSync(path.join(__dirname, "corpus.json"), "utf8")); }
catch (e) { console.warn("corpus.json not found — /topic will 404"); }

app.get("/", (_req, res) => res.send("ADCDA Build Service — POST /build | GET /topic/:code | GET /health"));

app.get("/health", (_req, res) => res.json({ ok: true, topics: Object.keys(CORPUS).length }));

app.get("/topic/:code", (req, res) => {
  const code = String(req.params.code || "").trim().toUpperCase();
  const t = CORPUS[code];
  if (!t) return res.status(404).json({ error: "topic not found", code });
  res.json(t);
});

app.get("/files/:code", (req, res) => {
  const code = String(req.params.code || "").replace(/[^A-Za-z0-9.]/g, "");
  const fp = path.join(__dirname, code + ".pptx");
  if (!fs.existsSync(fp)) return res.status(404).json({ error: "file not found", code });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
  res.setHeader("Content-Disposition", 'inline; filename="' + code + '.pptx"');
  res.sendFile(fp);
});

app.post("/build", (req, res) => {
  const content = req.body || {};
  const id = "job_" + Date.now() + "_" + Math.random().toString(36).slice(2, 7);
  const work = path.join(os.tmpdir(), id);
  fs.mkdirSync(work, { recursive: true });
  const jsonPath = path.join(work, "content.json");
  const outName = ((content.code || "F.X").replace(/\./g, "")) + "_ADCDA.pptx";
  const outPath = path.join(work, outName);
  content.output = outPath;
  try {
    fs.writeFileSync(jsonPath, JSON.stringify(content), "utf8");
    execFileSync("node", [BUILD, jsonPath], { stdio: "pipe" });
    try { execFileSync("python3", [FIXRTL, outPath], { stdio: "pipe" }); } catch (e) {}
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", 'attachment; filename="' + outName + '"');
    res.send(fs.readFileSync(outPath));
  } catch (err) {
    res.status(500).json({ error: String(err && err.message || err) });
  } finally {
    try { fs.rmSync(work, { recursive: true, force: true }); } catch (_) {}
  }
});

app.listen(PORT, () => console.log("ADCDA Build Service on port " + PORT));
