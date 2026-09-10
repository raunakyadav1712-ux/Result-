"""
Result Analysis Web Tool — public, brandable.
Any coaching pastes its test results, sets its institute name/logo, and downloads
a ready branded analysis PDF. Wraps the same generator engine used to build the decks.

Run locally:   uvicorn app:app --host 0.0.0.0 --port 8000
Deploy:        see README.md (Docker / Render one-click).
"""
import os, tempfile, hashlib, io
from fastapi import FastAPI, Form, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from PIL import Image, ImageDraw, ImageFont

import generate as G
from parse_input import parse_students, is_neet

app = FastAPI(title="Result Analysis Tool")
WORK = tempfile.mkdtemp(prefix="ratool_")

INDEX = """<!doctype html><html lang=en><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1">
<title>Result Analysis Tool</title>
<style>
 :root{--bg:#161324;--card:#221c38;--bd:#3a3358;--ink:#f2f0fa;--mut:#a8a2c2;
   --gold:#f5b90a;--acc:#c0abf2;--grn:#36d392}
 *{box-sizing:border-box}body{margin:0;font-family:system-ui,Segoe UI,Roboto,Arial;
   background:var(--bg);color:var(--ink);line-height:1.5}
 .wrap{max-width:820px;margin:0 auto;padding:26px 18px 60px}
 h1{font-size:26px;margin:.2em 0}.sub{color:var(--mut);margin:0 0 22px}
 label{display:block;font-weight:600;margin:14px 0 5px;font-size:14px}
 input,select,textarea{width:100%;padding:11px 12px;border-radius:10px;border:1px solid var(--bd);
   background:#1c1830;color:var(--ink);font-size:15px;font-family:inherit}
 textarea{min-height:230px;font-family:ui-monospace,Menlo,Consolas,monospace;font-size:13px}
 .row{display:flex;gap:12px;flex-wrap:wrap}.row>div{flex:1;min-width:150px}
 .card{background:var(--card);border:1px solid var(--bd);border-radius:16px;padding:18px 18px 22px;margin-top:16px}
 button{margin-top:20px;width:100%;padding:14px;border:0;border-radius:12px;cursor:pointer;
   background:linear-gradient(90deg,#6b47ae,#452a78);color:#fff;font-size:16px;font-weight:700}
 .hint{color:var(--mut);font-size:12.5px;margin-top:6px}
 code{background:#1c1830;padding:1px 5px;border-radius:5px;color:var(--acc)}
 .gold{color:var(--gold)}.acc{color:var(--acc)}
 pre{white-space:pre-wrap;background:#1c1830;border:1px solid var(--bd);border-radius:10px;
   padding:12px;color:var(--mut);font-size:12.5px;overflow:auto}
</style></head><body><div class=wrap>
<h1>Result <span class=acc>Analysis</span> Tool</h1>
<p class=sub>Paste a test's results → download a branded, video-ready analysis PDF. Works for JEE & NEET.</p>
<form method=post action=/generate enctype=multipart/form-data class=card>
 <div class=row>
   <div><label>Institute name</label><input name=institute placeholder="Your Coaching Name" required></div>
   <div><label>Logo (optional)</label><input type=file name=logo accept="image/*"></div>
 </div>
 <div class=row>
   <div><label>Batch / Class</label><input name=batch placeholder="12th JEE" required></div>
   <div><label>Date</label><input name=date placeholder="10 Sep 2026"></div>
   <div><label>Exam</label><select name=exam>
     <option value=auto>Auto-detect</option><option value=jee>JEE</option><option value=neet>NEET</option></select></div>
 </div>
 <div class=row>
   <div><label>Correct mark</label><input name=cmark value=4 type=number step=0.25></div>
   <div><label>Wrong mark</label><input name=wmark value=-1 type=number step=0.25></div>
 </div>
 <label>Student results</label>
 <textarea name=data required placeholder="Abhay Singh
Physics   C14 W6 U5
Chemistry C12 W7 U6
Maths     C14 W5 U6

Vaishnavi Prajapati
Physics   C11 W10 U4
Chemistry C8 W13 U4
Maths     C16 W8 U1"></textarea>
 <div class=hint>One name line per student, then a line per subject.
 <code>C</code>=correct, <code>W</code>=wrong, <code>U</code>=unattempted. Add <code>absent</code>
 on a name line for absentees, <code>prev 120</code> for last-test improvement. NEET → use Biology.</div>
 <button type=submit>Generate PDF</button>
</form>
<div class=card><b class=gold>Tips</b>
<pre>• Accuracy, negatives, rank, percentile, class average, leaderboard — all auto-computed.
• Marks are recomputed from correct/wrong, so manual add-up errors get fixed.
• Leave "Logo" empty and a clean initials badge is used from your institute name.</pre></div>
</div></body></html>"""


def _fallback_logo(institute):
    """Circular initials badge when no logo uploaded."""
    parts = [p for p in institute.replace(".", " ").split() if p]
    initials = "".join(p[0] for p in parts[:2]).upper() or "R"
    D = 512
    im = Image.new("RGBA", (D, D), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.ellipse((6, 6, D - 6, D - 6), fill=(74, 42, 123, 255), outline=(245, 185, 10, 255), width=16)
    try:
        f = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 210)
    except Exception:
        f = ImageFont.load_default()
    tb = d.textbbox((0, 0), initials, font=f)
    d.text(((D - (tb[2] - tb[0])) / 2 - tb[0], (D - (tb[3] - tb[1])) / 2 - tb[1]),
           initials, font=f, fill=(255, 255, 255, 255))
    p = os.path.join(WORK, "logo_" + hashlib.md5(institute.encode()).hexdigest() + ".png")
    im.save(p)
    return p


@app.get("/", response_class=HTMLResponse)
def index():
    return INDEX


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/generate")
async def generate(institute: str = Form(...), batch: str = Form(...),
                   date: str = Form(""), exam: str = Form("auto"),
                   cmark: float = Form(4), wmark: float = Form(-1),
                   data: str = Form(...), logo: UploadFile = File(None)):
    students = parse_students(data)
    if not students:
        return JSONResponse({"error": "Could not read any students. Check the format."}, status_code=400)

    neet = is_neet(students) if exam == "auto" else (exam == "neet")
    # logo
    logo_path = None
    if logo is not None and logo.filename:
        raw = await logo.read()
        try:
            im = Image.open(io.BytesIO(raw)).convert("RGBA")
            im = im.crop(im.getbbox() or (0, 0, im.width, im.height))
            side = max(im.size)
            canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
            canvas.paste(im, ((side - im.width) // 2, (side - im.height) // 2), im)
            logo_path = os.path.join(WORK, "up_" + hashlib.md5(raw).hexdigest() + ".png")
            canvas.save(logo_path)
        except Exception:
            logo_path = None
    if logo_path is None:
        logo_path = _fallback_logo(institute)

    meta = {"institute": institute.upper(), "test_title": "Result Analysis",
            "batch": batch, "date": date,
            "marking": {"correct": cmark, "incorrect": wmark},
            "default_questions": 45 if neet else 25,
            "class_total": len(students),
            "present": len([s for s in students if not s.get("absent")]),
            "logo": logo_path}

    stem = f"{batch} - Result Analysis".strip().replace("/", "-")
    out = os.path.join(WORK, stem + ".pdf")
    try:
        G.build({"meta": meta, "students": students}, out)
    except Exception as e:
        return JSONResponse({"error": f"Generation failed: {e}"}, status_code=500)
    return FileResponse(out, filename=stem + ".pdf", media_type="application/pdf")
