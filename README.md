# All In One Validation Formatter

**The GOAT. The one tool to replace them all.**

### Run Locally (Easiest)

**Option 1 (Recommended for Windows):**
- Double-click `run_app.bat` or `run_app.ps1`

**Option 2:**
```bash
streamlit run app.py
```

### Deploy to Streamlit Community Cloud (recommended)

1. Push this entire folder to a GitHub repo.
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Point it at `app.py`
4. Add `requirements.txt` (already present)

Primary tool: **AIO Formatter** (the first / main page).

It works for **any hall, any fabric type, any source** (Slack or LVV Portal). The cutsheet tells it everything it needs.

## How the AIO Works (Core Idea)

The **cutsheet is the source of truth**.

When you give the AIO a cutsheet + raw validation files, it inspects the cutsheet for:
- Hall / room (JPB19, JBP15, SYD20, future halls...)
- Fabric / pathway type (GFAB, CFAB, HOPS, LV-Portal T0-to-Host, etc.)
- Source system (Slack export vs LVV Portal)
- Any other signals we put in the cutsheet

It then automatically dispatches to the correct rich logic we have already built and battle-tested.

This is why one tool can handle every eventuality.

## Current Status

| Tool              | Status   | Recommended?          | Page                     | Notes |
|-------------------|----------|-----------------------|--------------------------|-------|
| **AIO Formatter** | ✅ Ready | **Yes — use this first** | `00_AIO_Formatter.py`   | Master cutsheet-based detector + dispatcher |
| GFAB              | ✅ Ready | Only when explicit    | `01_JPB19_GFAB.py`       | Full rich port |
| HOPS              | ✅ Ready | Only when explicit    | `02_JPB19_HOPS.py`       | Full rich port (GPU + combined cutsheet) |
| CFAB              | ✅ Ready | Only when explicit    | `03_JPB19_CFAB.py`       | Full rich port (T3-T2/T2-T1-T0 + PP) |
| IPR               | ✅ Ready | Only when explicit    | `04_JPB19_IPR.py`        | Multi-building combiner + enrichment |

## How to Run

```bash
cd "C:\Users\toddy\Desktop\TODD\All in One"
streamlit run Home.py
```

**Start on the AIO Formatter page** (it appears first in the sidebar).

The AIO is designed so validators almost never need to pick a specialized tool manually.

## Architecture & Extensibility

The app is deliberately designed to be easy to evolve as new halls and ideas come from testing.

**Recommended structure (already in place):**

- `logic/` — All heavy business logic lives here (easy to extend)
- `app.py` — Main experience + feedback form
- `pages/` — Specialized deep tools (secondary)

See the "🛠️ For Developers" expander inside the running app for the current recommended process to add new fabric support.

The goal is that validators almost never have to think "which script do I use?" — they just use the AIO.
- Easy testing and future improvements (mismatch detection, filtering, new columns, etc.)
- No more "it only works on my laptop" problems

## Adding a New Tool

1. Create `jpb19_newtool_logic.py` with at least:
   - `process_newtool_validation(input_path, cutsheet_path, ...) -> (bytes, suggested_filename)`
   - `process_multiple_newtool_files(paths, cutsheet_path) -> zip_bytes`

2. Create `pages/0N_NewTool.py` using the exact defensive form + tempdir pattern from the existing pages.

3. Add a short description to `Home.py`.

## Original Scripts (for reference)

All original Claude/Tkinter versions live in:
`LEE's Script\JBP19 ... Script\`

Do **not** edit them. The versions in this folder are the maintained ones.

---

**Goal:** Every JPB19 validation reporting job eventually lives here in clean, reliable, team-accessible form.
