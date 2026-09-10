# Result Analysis Tool (public web app)

A brandable web tool: any coaching sets its **institute name + logo**, pastes a test's
results, and downloads a ready **analysis PDF** (dark theme, per-student scorecards +
analytics, class insights, leaderboard). JEE & NEET, auto-detected.

It wraps the same engine (`generate.py`) that produces the decks — so the output is
identical quality.

## What's inside
- `app.py` — the web app (FastAPI): a form + a `/generate` endpoint that returns the PDF.
- `parse_input.py` — turns pasted results into structured data (handles one-line and
  multi-line "CORRECT / INCORRECT / UNATTEMPTED" layouts).
- `generate.py` + `assets/` — the analysis engine and default logo.
- `Dockerfile`, `render.yaml`, `requirements.txt` — deployment.

## Run locally
```bash
pip install -r requirements.txt         # needs Python 3.11+
# For PDF export you also need LibreOffice installed (apt-get install libreoffice-impress)
uvicorn app:app --host 0.0.0.0 --port 8000
# open http://localhost:8000
```

## Deploy so anyone can use it (free, ~5 min)
The app needs **LibreOffice** on the server to export PDF — the included `Dockerfile`
installs it, so deploy via Docker anywhere.

**Render.com (easiest):**
1. Push this folder to a GitHub repo.
2. On render.com → New → Web Service → connect the repo.
3. Render reads `render.yaml` (Docker, free plan, health check `/health`). Click Deploy.
4. You get a public URL like `https://your-tool.onrender.com` — share it with any teacher.

**Railway / Fly.io / any VPS:** build the Docker image and run it:
```bash
docker build -t result-tool .
docker run -p 8000:8000 result-tool
```

> Free tiers sleep when idle; the first request after a nap takes ~30s to wake. For heavy
> use pick a paid small instance.

## Input format (shown on the page too)
```
Abhay Singh
Physics   C14 W6 U5
Chemistry C12 W7 U6
Maths     C14 W5 U6

Bhavana Dubey                 (NEET → use Biology)
Physics 69 marks
correct 21
incorrect 15
unattempted 9
...
```
`C`=correct, `W`=wrong, `U`=unattempted. `absent` on a name line = absentee;
`prev 120` = last-test score (adds an improvement metric). Marks are recomputed from
correct/wrong (so manual add-up errors are fixed automatically).

## Coming next (needs a sample)
Your OMR scanner app exports a **PDF report**. Send one real exported PDF and I'll add an
"upload OMR PDF" button that auto-reads every student's marks — so the flow becomes
**scan → upload → branded PDF** with zero typing.
