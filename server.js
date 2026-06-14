// ===== ADCDA Build Service =====
// خدمة بناء العروض: تستقبل JSON المحتوى وترجع ملف PPTX بهوية ADCDA.
// التشغيل:  node server.js   (المنفذ الافتراضي 3000)
// الاستدعاء من n8n (عقدة Design PPT): POST http://<host>:3000/build  body = JSON المحتوى
//   النتيجة: ملف .pptx (binary) — اربطه بعقدة Save Files.
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

app.get("/", (_req, res) => res.send("ADCDA Build Service — POST /build with content JSON"));

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
    // 1) build PPTX with the fixed ADCDA template
    execFileSync("node", [BUILD, jsonPath], { stdio: "pipe" });
    // 2) RTL fix for Arabic textboxes
    try { execFileSync("python3", [FIXRTL, outPath], { stdio: "pipe" }); } catch (e) { /* fix_rtl optional */ }
    // 3) return the file
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", 'attachment; filename="' + outName + '"');
    res.send(fs.readFileSync(outPath));
  } catch (err) {
    res.status(500).json({ error: String(err && err.message || err) });
  } finally {
    try { fs.rmSync(work, { recursive: true, force: true }); } catch (_) {}
  }
});

app.listen(PORT, () => console.log("✅ ADCDA Build Service on port " + PORT + "  (POST /build)"));
