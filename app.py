#!/usr/bin/env python3
"""
All In One Validation Formatter

The master / universal tool for formatting validation reports from any hall.

Run with:
    streamlit run app.py
"""

import streamlit as st
import sys
from pathlib import Path

# Robust import so it works both locally and on Streamlit Cloud
root_dir = Path(__file__).parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

st.set_page_config(
    page_title="All In One Validation Formatter",
    page_icon="🧠",
    layout="wide"
)

st.title("All In One Validation Formatter")
st.caption("The GOAT • Any Hall • Any Fabric • Any Source • One Tool to Rule Them All")

st.markdown("""
### Welcome to the Master Tool

This is the **definitive replacement** for every previous validation formatter.

- Upload your raw files + the cutsheet
- It reads the cutsheet to understand exactly what this is
- It runs the correct rich logic automatically
- **Before you download**, you get a live analytics dashboard showing error counts, breakdowns, and the Summary preview

No more guessing which script to use. No more opening files just to see how bad it is.
""")

# Hero recommendation
st.success("**👉 Start here: [00_AIO_Formatter](00_AIO_Formatter)** — The All-In-One tool. Upload anything + its cutsheet. It figures out the rest.")

st.markdown("""
### Available Tools

| Tool | When to use it directly |
|------|-------------------------|
| **AIO Formatter** (recommended) | Almost always. Detects hall, fabric type (GFAB/CFAB/HOPS/LV-Portal/etc.), source (Slack vs LVV Portal) from the cutsheet and runs the right logic automatically. |
| **GFAB** | You know it's a standard GFAB fabric and want the classic GFAB formatter. |
| **HOPS** | Explicit GPU/HOPS validation (the one with GPU sheets and combined DeviceA/DeviceB cutsheet). |
| **CFAB** | CFAB with T3-T2 vs T2-T1-T0 long/short path splitting. |
| **IPR** | Multi-building audit combining + cutsheet enrichment. |

---

### How the AIO Works (the important bit)

The **cutsheet is the source of truth**.

When you upload a cutsheet, the AIO inspects it for:
- Hall / room (JPB19, JBP15, SYD20, or any future hall)
- Fabric / pathway type (GFAB, CFAB, HOPS, LV-Portal T0-to-Host, etc.)
- Source system (Slack exports vs LVV Portal reports)
- Any other signals we put in the cutsheet

It then automatically uses the correct rich processing logic.

One tool. Every hall. Every type. Every source.

---

### Architecture

All tools follow the same clean pattern:
- Heavy domain logic lives in `*_logic.py` (pure, importable, testable)
- UI pages are thin (uploads + forms + progress + downloads)
- The AIO (`logic/aio_logic.py`) is the orchestrator on top

**Status — The Master Tool**

This is the definitive All-In-One replacement for every previous specialized formatter.

The specialized pages (GFAB, HOPS, CFAB, IPR) are still available in the sidebar for power users, but **99% of the time you should just use the AIO Formatter above**.
""")

st.info("**👉 Recommended:** Go to the **AIO Formatter** page in the sidebar (it appears first). That's the master tool.", icon="⭐")

st.divider()

# =============================================================================
# SUGGESTIONS + FEEDBACK (GitHub friendly)
# =============================================================================
st.subheader("💡 Suggest a New Feature or Improvement")

st.markdown("""
The team is actively testing and coming up with ideas. This section makes it **dead simple** to capture feedback that can go straight into GitHub.

**How it works on Streamlit Cloud:**
1. Fill out the form below.
2. Click **"Generate GitHub Issue Link"**.
3. It will open (or give you a link to) a pre-filled GitHub Issue with your text already populated.
4. Just submit the issue. No login required for the link generation.
""")

with st.form("suggestion_form", clear_on_submit=True):
    suggestion_title = st.text_input("Short title for your idea", placeholder="e.g. Add mismatch severity scoring")
    suggestion_body = st.text_area(
        "Describe the feature or improvement in as much detail as you want",
        height=150,
        placeholder="What problem does this solve? What would the ideal behavior look like? Any specific halls or report types?"
    )
    affected_area = st.multiselect(
        "Which area does this affect? (helps with triage)",
        ["Detection / Cutsheet Intelligence", "Analytics Dashboard", "GFAB logic", "HOPS logic", "CFAB logic", "LV Portal logic", "UI/UX", "Performance", "Other"],
        default=[]
    )
    submitted_suggestion = st.form_submit_button("🚀 Generate Pre-filled GitHub Issue Link", type="primary")

if submitted_suggestion:
    if not suggestion_title or not suggestion_body:
        st.error("Please provide both a title and a description.")
    else:
        area_text = ", ".join(affected_area) if affected_area else "General"
        
        body = f"""## Feature / Improvement Suggestion

**Affected areas:** {area_text}

### Description
{suggestion_body}

---
*Submitted via All In One Validation Formatter*
"""

        repo = "YOUR_USERNAME/YOUR_REPO"   # <--- CHANGE THIS after pushing to GitHub!

        if "YOUR_USERNAME" in repo or "YOUR_REPO" in repo:
            st.warning("⚠️ The GitHub repo placeholder hasn't been updated yet. Please edit `app.py` and set the correct repo name, then redeploy.")
            st.code(f"repo = \"YOUR_USERNAME/YOUR_REPO\"", language="python")
            st.info("Once updated, the suggestions form will generate real GitHub issue links.")
        else:
            import urllib.parse
            params = {
                "title": f"[Suggestion] {suggestion_title}",
                "body": body,
                "labels": "enhancement"
            }
            query = urllib.parse.urlencode(params, quote_via=urllib.parse.quote)
            issue_url = f"https://github.com/{repo}/issues/new?{query}"

            st.success("Link generated!")
            st.markdown(f"**Click here to open a pre-filled GitHub Issue:**")
            st.markdown(f"[➡️ Open GitHub Issue with your suggestion]({issue_url})", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("**Or copy the markdown below** and paste it into a new GitHub Issue manually:")
        st.code(f"**Title:** [Suggestion] {suggestion_title}\n\n{body}", language="markdown")

st.divider()

# =============================================================================
# EXTENSIBILITY NOTES (for the team)
# =============================================================================
with st.expander("🛠️ For Developers: How to Add Support for a New Hall or Fabric Type", expanded=False):
    st.markdown("""
    This app was built to be **easy to extend** as new halls and requirements come in.

    **Recommended process:**

    1. Take the original messy Tkinter/Claude script for the new hall/type.
    2. Port it into a clean `logic/new_hall_logic.py` following the same pattern as the existing ones (pure functions, `process_xxx_validation` + multi-file version).
    3. Add detection rules in `logic/aio_logic.py` inside `inspect_cutsheet()`.
    4. Register the new processor in the `PROCESSORS` dictionary.
    5. (Optional) Add a dedicated specialized page if needed.
    6. Update this suggestions section + README with the new capability.

    The goal is that the AIO keeps getting smarter over time with minimal friction.

    **Current processors live in the `logic/` folder.**
    """)

st.markdown("""
**Need help or found an issue?**  
Talk to the person maintaining these ports. The goal is one reliable, fast, shared place for validation reporting across any hall.
""")
