// ===== ADCDA Build Service v2.3 =====
// خدمة بناء العروض: POST /build | GET /topic/:code | GET /search | GET /files/:code | GET /pdf/:code | GET /health
// v2.3: /build يقبل upload_url اختياريًا — يبني ويرفع الملف مباشرة إلى OneDrive عبر جلسة رفع مفوّضة
// ويرجع JSON خفيفًا { uploaded, name, size, id, webUrl } — الملف لا يمر عبر n8n إطلاقًا.
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

app.get("/", (_req, res) => res.send("ADCDA Build Service v2.3 — POST /build | GET /topic/:code | GET /health"));

app.get("/dash", (_req, res) => res.sendFile(path.join(__dirname, "dash.html")));

app.get("/health", (_req, res) => res.json({ ok: true, topics: Object.keys(CORPUS).length, version: "2.1" }));

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

// ===== المدقق الرسمي v1.1 — v