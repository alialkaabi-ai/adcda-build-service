"""يصلح خصائص RTL فقط للنصوص العربية (تحتوي على أحرف عربية)"""
import zipfile, re, shutil, os, sys

src = sys.argv[1]
work = "/tmp/pptx_fix_" + str(os.getpid())
os.makedirs(work, exist_ok=True)

# فك ضغط
with zipfile.ZipFile(src, 'r') as z:
    z.extractall(work)

ARABIC_RE = re.compile(r'[\u0600-\u06FF]')

def fix_textbox(match):
    """يأخذ نص <p:txBody>...</p:txBody> ويرجعه بعد إصلاح pPr فقط لو فيه عربي"""
    txbody = match.group(0)
    # هل التكست بوكس فيه نص عربي؟
    has_arabic = bool(ARABIC_RE.search(txbody))
    if not has_arabic:
        return txbody
    # أضف rtl="1" فقط للـ pPr اللي ما فيها rtl
    txbody = re.sub(
        r'<a:pPr algn="(r|ctr)"',
        r'<a:pPr rtl="1" algn="\1"',
        txbody
    )
    # أضف rtlCol="1" للـ bodyPr داخل التكست بوكس العربي فقط
    txbody = re.sub(
        r'rtlCol="0"',
        r'rtlCol="1"',
        txbody
    )
    return txbody

slides_dir = os.path.join(work, "ppt", "slides")
fixes = 0
for fname in sorted(os.listdir(slides_dir)):
    if not fname.endswith(".xml"):
        continue
    fpath = os.path.join(slides_dir, fname)
    with open(fpath, "r", encoding="utf-8") as f:
        content = f.read()
    original = content
    # طبق على كل txBody
    content = re.sub(r'<p:txBody>.*?</p:txBody>', fix_textbox, content, flags=re.DOTALL)
    if content != original:
        with open(fpath, "w", encoding="utf-8") as f:
            f.write(content)
        fixes += 1

# إعادة الضغط
out = src
with zipfile.ZipFile(out + ".tmp", 'w', zipfile.ZIP_DEFLATED) as zout:
    for root, _, files in os.walk(work):
        for file in files:
            full = os.path.join(root, file)
            arc = os.path.relpath(full, work)
            zout.write(full, arc)

shutil.move(out + ".tmp", out)
shutil.rmtree(work)
print(f"✅ Fixed RTL in {fixes} slides (Arabic-only textboxes)")
