// ===== ADCDA Build Service v2.4 =====
// POST /build | GET /topic/:code | GET /search | GET /files/:code | GET /pdf/:code | GET /dash | GET /portal | GET /health
// v2.4: استضافة بوابة المحتوى على GET /portal (بجانب مركز القيادة /dash)
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

let CORPUS = {};
try { CORPUS = JSON.parse(fs.readFileSync(path.join(__dirname, "corpus.json"), "utf8")); }
catch (e) { console.warn("corpus.json not found — /topic will 404"); }

app.get("/", (_req, res) => res.send("ADCDA Build Service v2.4 — POST /build | GET /topic/:code | GET /dash | GET /portal | GET /health"));

app.get("/dash", (_req, res) => res.sendFile(path.join(__dirname, "dash.html")));
app.get("/dash2", (_req, res) => res.sendFile(path.join(__dirname, "dash2.html")));

app.get("/portal", (_req, res) => res.sendFile(path.join(__dirname, "portal.html")));

app.get("/health", (_req, res) => res.json({ ok: true, topics: Object.keys(CORPUS).length, version: "2.4" }));

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

const VALIDATOR = path.join(__dirname, "validator.py");

app.post("/validate", (req, res) => {
  const doc = req.body || {};
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
  try { const qz = extractQuiz(content); if (qz && qz.code) fs.writeFileSync(path.join(QUIZ_DIR, qz.code.replace(/[^\w.\-]/g, "") + ".json"), JSON.stringify(qz), "utf8"); } catch (e) {}
  try {
    fs.writeFileSync(jsonPath, JSON.stringify(content), "utf8");
    execFileSync("node", [BUILD, jsonPath], { stdio: "pipe" });
    try { execFileSync("python3", [FIXRTL, outPath], { stdio: "pipe" }); } catch (e) {}
    const buf = fs.readFileSync(outPath);
    if (uploadUrl) {
      const item = await putToUploadUrl(uploadUrl, buf);
      res.json({ uploaded: true, name: item.name || outName, size: buf.length, id: item.id || null, webUrl: item.webUrl || null });
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

// ─── نظام قياس الأثر: اختبار تفاعلي لكل عرض (QR) ───
const QUIZ_DIR = path.join(__dirname, "quizzes");
try { fs.mkdirSync(QUIZ_DIR, { recursive: true }); } catch (e) {}

function extractQuiz(content) {
  const slides = content.slides || [];
  const qz = slides.find(s => s && s.type === "quiz" && Array.isArray(s.qa) && s.qa.length);
  if (!qz) return null;
  const key = slides.find(s => s && s.type === "list" && /رسائل ذهبية/.test(s.title || ""));
  const pt = p => Array.isArray(p) ? p.map(r => (r && r.text) || "").join("") : (p && typeof p === "object" ? ((p.lead || "") + " " + (p.text || "")) : String(p || ""));
  return {
    code: content.code || "",
    title: content.title_ar || content.title || "",
    main_message: content.main_message || "",
    qa: qz.qa.slice(0, 6).map(x => ({ q: String(x.q || ""), a: String(x.a || "") })),
    messages: key ? (key.points || []).map(pt).slice(0, 5) : []
  };
}

app.get("/quiz-data/:code", (req, res) => {
  const f = path.join(QUIZ_DIR, String(req.params.code).replace(/[^\w.\-]/g, "") + ".json");
  if (!fs.existsSync(f)) return res.status(404).json({ error: "لا يوجد اختبار لهذا العرض بعد" });
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.send(fs.readFileSync(f, "utf8"));
});

app.get("/quiz/:code", (_req, res) => res.sendFile(path.join(__dirname, "quiz.html")));


app.listen(PORT, () => console.log("ADCDA Build Service v2.4 on port " + PORT));
