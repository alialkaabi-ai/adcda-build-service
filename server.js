// ===== ADCDA Build Service v2.3 =====
// خدمة بناء العروض: POST /build | GET /topic/:code | GET /search | GET /files/:code | GET /pdf/:code | GET /dash | GET /health
// v2.1: /build يقبل upload_url اختياريًا — يبني ويرفع الملف مباشرة إلى OneDrive عبر جلسة رفع مفوّضة
// ويرجع JSON خفيفًا { uploaded, name, size, id, webUrl } — الملف لا يمر عبر n8n إطلاقًا.
// v2.3: استضافة مركز القيادة (الداش بورد) على GET /dash
const express = require("express");
const { execFileSync } = require("child_process");
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");

const app = express();
app.use(express.json({ limit: "5mb" }));

const PORT = process.env.PORT || 3000;
const BUILD = path.join(__dirname, "build.js");
const FIXRTL = path.join(__dirname, "fix_rtl.py");

// ===== قاعدة المعرفة (التأصيل) =====
let CORPUS = {};
try { CORPUS = JSON.parse(fs.readFileSync(path.join(__dirname, "corpus.json"), "utf8")); }
catch (e) { console.warn("corpus.json not found — /topic will 404"); }

app.get("/", (_req, res) => res.send("ADCDA Build Service v2.3 — POST /build | GET /topic/:code | GET /dash | GET /health"));

app.get("/dash", (_req, res) => res.sendFile(path.join(__dirname, "dash.html")));

app.get("/health", (_req, res) => res.json({ ok: true, topics: Object.keys(CORPUS).length, version: "2.3" }));

app.get("/topic/:code", (req, res) => {
  const code = String(req.params.code || "").trim().toUpperCase();
  const t = CORPUS[code];
  if (!t) return res.status(404).json({ error: "topic not found", code });
  res.json(t);
});

app.get("/files/:code", (req, res) => {
  const code = String(req.params.code || "").replace(/[^A-Za-z0-9.]/g, "");
  const suf = req.query.lang === "en" ? "_en" : "";
  const fp = path.join(__dirname, code + suf + ".pptx");
  if (!fs.existsSync(fp)) return res.status(404).json({ error: "file not found", code });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
  res.setHeader("Content-Disposition", 'inline; filename="' + code + '.pptx"');
  res.sendFile(fp);
});

app.get("/pdf/:code", (req, res) => {
  const code = String(req.params.code || "").replace(/[^A-Za-z0-9.]/g, "");
  const suf = req.query.lang === "en" ? "_en" : "";
  const fp = path.join(__dirname, code + suf + ".pdf");
  if (!fs.existsSync(fp)) return res.status(404).json({ error: "pdf not found", code });
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", 'inline; filename="' + code + '.pdf"');
  res.sendFile(fp);
});

app.get("/search", (req, res) => {
  const q = String(req.query.q || "").trim();
  if (!q) return res.json({ query: q, matches: [] });
  const stop = new Set(["في","من","على","عند","الى","إلى","عن","مع","the","of","and","a","an"]);
  const toks = q.replace(/[^ء-يa-zA-Z0-9 ]/g, " ").split(/\s+/).filter(t => t.length >= 3 && !stop.has(t));
  const scored = Object.values(CORPUS).map(t => {
    const hay = (t.title + " " + (t.full_text || "")).toLowerCase();
    let score = 0;
    for (const tk of toks) { if (hay.includes(tk.toLowerCase())) score++; }
    return { topic_code: t.topic_code, title: t.title, package: t.package, score, coverage: toks.length ? score / toks.length : 0 };
  }).filter(x => x.score > 0).sort((a, b) => b.score - a.score).slice(0, 5);
  res.json({ query: q, tokens: toks, matches: scored });
});

// ===== المدقق الرسمي v1.1 — validator.py يعمل كما هو (مصدر الحقيقة الواحد) =====
const VALIDATOR = path.join(__dirname, "validator.py");

app.post("/validate", (req, res) => {
  const doc = req.body || {};
  // حقن حقول التوافق فقط إن غابت — بلا تغيير أي فحص في validator.py
  if (doc.topic_profile == null) doc.topic_profile = "";
  if (!doc.standards) {
    const s15 = (doc.slides || []).find(s => s && s.type === "sources") || {};
    doc.standards = {
      conditional: { iso45001: false, iso22320_22322: false },
      compliance_line_ar: (Array.isArray(s15.iso26000) && s15.iso26000.length)
        ? "نسترشد بموجهات ISO 26000 — البنود " + s15.iso26000.join("، ")
        : ""
    };
  }
  const tmp = path.join(os.tmpdir(), "val_" + Date.now() + "_" + Math.random().toString(36).slice(2, 7) + ".json");
  let out = "", ok = false;
  try {
    fs.writeFileSync(tmp, JSON.stringify(doc), "utf8");
    try {
      out = execFileSync("python3", [VALIDATOR, tmp], { stdio: "pipe", encoding: "utf8" });
      ok = true;
    } catch (e) {
      out = String((e.stdout || "") + (e.stderr || ""));
      ok = false;
    }
    const checks = out.split("\n").filter(l => /^[✅❌⏳]/.test(l.trim())).map(l => l.trim());
    const fails = checks.filter(l => l.startsWith("❌"));
    res.json({ ok, source: "validator.py v1.1 (رسمي)", pass: checks.filter(l => l.startsWith("✅")).length,
      fail: fails.length, pending: checks.filter(l => l.startsWith("⏳")).length, fails, checks });
  } catch (err) {
    res.status(500).json({ error: String(err && err.message || err) });
  } finally {
    try { fs.unlinkSync(tmp); } catch (_) {}
  }
});

// ===== الرفع المباشر إلى OneDrive عبر جلسة رفع مفوّضة =====
function putToUploadUrl(uploadUrl, buf) {
  return new Promise((resolve, reject) => {
    const u = new URL(uploadUrl);
    const req = https.request({
      hostname: u.hostname,
      path: u.pathname + u.search,
      method: "PUT",
      headers: {
        "Content-Length": buf.length,
        "Content-Range": "bytes 0-" + (buf.length - 1) + "/" + buf.length
      }
    }, r => {
      let d = "";
      r.on("data", c => d += c);
      r.on("end", () => {
        if (r.statusCode >= 200 && r.statusCode < 300) {
          try { resolve(JSON.parse(d)); } catch (e) { resolve({ ok: true }); }
        } else reject(new Error("upload failed " + r.statusCode + ": " + String(d).slice(0, 300)));
      });
    });
    req.on("error", reject);
    req.end(buf);
  });
}

app.post("/build", async (req, res) => {
  const content = req.body || {};
  const uploadUrl = content.upload_url;
  delete content.upload_url;
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
    const buf = fs.readFileSync(outPath);
    if (uploadUrl) {
      const item = await putToUploadUrl(uploadUrl, buf);
      res.json({
        uploaded: true,
        name: item.name || outName,
        size: buf.length,
        id: item.id || null,
        webUrl: item.webUrl || null
      });
    } else {
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
      res.setHeader("Content-Disposition", 'attachment; filename="' + outName + '"');
      res.send(buf);
    }
  } catch (err) {
    res.status(500).json({ error: String(err && err.message || err) });
  } finally {
    try { fs.rmSync(work, { recursive: true, force: true }); } catch (_) {}
  }
});

app.listen(PORT, () => console.log("ADCDA Build Service v2.3 on port " + PORT));
