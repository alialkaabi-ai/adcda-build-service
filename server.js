// ===== ADCDA Build Service v2 =====
// POST /build  body = JSON المحتوى
//   بدون upload_url: يرجع ملف PPTX (binary) — السلوك القديم كما هو.
//   مع upload_url (جلسة رفع OneDrive من n8n): يبني ويرفع الملف مباشرة إلى OneDrive
//   ويرجع JSON خفيفًا { uploaded, name, size, id, webUrl } — الملف لا يمر عبر n8n إطلاقًا.
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

app.get("/", (_req, res) => res.send("ADCDA Build Service v2 — POST /build with content JSON (optional upload_url)"));

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
    try { execFileSync("python3", [FIXRTL, outPath], { stdio: "pipe" }); } catch (e) { /* اختياري */ }
    const buf = fs.readFileSync(outPath);
    if (uploadUrl) {
      // رفع مباشر إلى OneDrive عبر جلسة الرفع المفوّضة — n8n لا يستقبل الملف
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

app.listen(PORT, () => console.log("✅ ADCDA Build Service v2 on port " + PORT + "  (POST /build)"));
