"""CREST Phase 1: Criteria Extraction & Generation Tool (Streamlit version).

The Anthropic API key is read on the server from Streamlit secrets (ANTHROPIC_API_KEY),
so visitors never see it and need no account, sign-in, or permission prompt.

Deployment (Streamlit Community Cloud):
  1. Add this file to the GitHub repository and add `anthropic` and `pypdf` to requirements.txt.
  2. Create a new app from the same repository with phase1_app.py as the main file.
  3. In the app's Settings > Secrets, add:
         ANTHROPIC_API_KEY = "sk-ant-..."
     Optional overrides (model IDs and the per-visitor request limit):
         MODEL_RECOMMENDED = "claude-sonnet-5-5"
         MODEL_FAST = "claude-haiku-4-5-20251001"
         MODEL_MOST_CAPABLE = "claude-opus-5-5"
         MAX_AI_CALLS_PER_SESSION = 40
"""

import io
import json
import os
import re
from datetime import datetime

import pandas as pd
import streamlit as st

# ================================================================
# PAGE CONFIGURATION AND STYLE
# ================================================================
st.set_page_config(page_title="CREST Phase 1: Criteria Extraction & Generation", page_icon="🧭",
                   layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    .stApp { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }
    .block-container { max-width: 1200px; padding-top: 2rem; }
    .crest-card { background: rgba(255,255,255,0.97); border-radius: 18px; padding: 26px 30px; margin-bottom: 18px;
                  box-shadow: 0 10px 30px rgba(0,0,0,0.18); color: #333; }
    .crest-title { text-align: center; font-size: 2.4rem !important; line-height: 1.2 !important; font-weight: 800; margin: 0;
                   background: linear-gradient(90deg,#667eea 0%,#764ba2 100%); -webkit-background-clip: text;
                   -webkit-text-fill-color: transparent; background-clip: text; }
    .crest-sub { text-align: center; font-weight: 600; color: #444; margin: 6px 0 14px 0; font-size: 1.15rem !important; }
    .crest-about { background: #f8f9fa; border-radius: 10px; padding: 16px 20px; font-size: 0.95rem; }
    .crest-about h4 { margin: 8px 0 4px 0; font-size: 1rem; }
    .info-blue { background: #e3f2fd; border-radius: 10px; padding: 14px 16px; margin: 8px 0 14px 0; }
    .info-green { background: #d4edda; border-radius: 10px; padding: 14px 16px; margin: 8px 0 14px 0; color: #155724; }
    .info-yellow { background: #fff3cd; border: 1px solid #ffeaa7; border-radius: 10px; padding: 14px 16px; margin: 8px 0 14px 0; }
    .crit { border: 1px solid #e3e6ea; border-radius: 10px; padding: 12px 14px; margin-bottom: 8px; background: #fff; }
    .crit-name { font-weight: 700; color: #495057; }
    .crit-desc { color: #6c757d; font-size: 0.9rem; margin-top: 3px; }
    .crit-meta { color: #666; font-size: 0.8rem; font-style: italic; margin-top: 3px; }
    .badge { display: inline-block; padding: 2px 9px; background: #e9ecef; border-radius: 12px; font-size: 0.75rem; color: #6c757d; }
    .stat { background: linear-gradient(135deg,#667eea 0%,#764ba2 100%); color: white; border-radius: 14px;
            padding: 16px; text-align: center; }
    .stat b { display: block; font-size: 2rem; }
    .st-key-panel { background: #ffffff; border-radius: 15px; padding: 26px 30px; box-shadow: 0 5px 15px rgba(0,0,0,0.12); }
    .st-key-tabbar { background: #ffffff; border-radius: 15px; padding: 8px; margin-bottom: 14px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); }
    .st-key-tabbar button { min-height: 58px; }
    .st-key-tabbar button p { white-space: normal !important; line-height: 1.25; font-weight: 600; }
    .stApp button[kind="primary"] { background: linear-gradient(90deg,#667eea 0%,#764ba2 100%); border: none; color: #fff; }
    .stApp button[kind="primary"]:hover { filter: brightness(1.07); color: #fff; }
</style>
""", unsafe_allow_html=True)

STEPS = [
    ("setup", "1. AI Setup (Required)"),
    ("domain", "2. Domain Setup (Required)"),
    ("manual", "3. Manual Addition (Optional)"),
    ("ai", "4. AI Criteria Generation (Optional)"),
    ("pdf", "5. PDF Analysis (Optional)"),
    ("results", "6. Consolidate & Export"),
]

OBJECTIVES_HELP = ("Objectives are the elements of interest to the stakeholders. They are connected to the stakeholders' "
                   "main values and refer to what they want to maximize, minimize, or optimize in the assessment of the "
                   "alternatives. Typically, an objective is either \"the higher the better\" or \"the lower the better\", "
                   "or its value is preferred within a given range. List one objective per line or separate them with semicolons.")


def secret(name, default=None):
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.environ.get(name, default)


MODEL_CHOICES = {
    "Recommended (balanced quality and speed)": secret("MODEL_RECOMMENDED", "claude-sonnet-5-5"),
    "Fast analysis": secret("MODEL_FAST", "claude-haiku-4-5-20251001"),
    "Most capable": secret("MODEL_MOST_CAPABLE", "claude-opus-5-5"),
}
MAX_CALLS = int(secret("MAX_AI_CALLS_PER_SESSION", 40))
PDF_CHARS = 6000          # characters of each PDF sent to the AI (same as the original page)

# ================================================================
# SESSION STATE
# ================================================================
DEFAULTS = {
    "step": "setup", "model_label": list(MODEL_CHOICES)[0], "personal_key": "",
    "domain": {"name": "", "context": "", "objective": ""},
    "manual": [], "ai": [], "pdf": [], "consolidated": None,
    "pdf_preview": "", "log": ["Ready for PDF analysis..."], "calls": 0, "next_id": 1,
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v if not isinstance(v, (dict, list)) else json.loads(json.dumps(v))
S = st.session_state


def new_id():
    S.next_id += 1
    return S.next_id


def log(msg):
    S.log.append(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


# ================================================================
# AI ACCESS (server side)
# ================================================================
def api_key():
    return S.personal_key.strip() or secret("ANTHROPIC_API_KEY", "")


def call_ai(prompt):
    """Send one prompt to the AI. The key never leaves the server."""
    import anthropic
    key = api_key()
    if not key:
        raise RuntimeError("The AI service is not configured yet. The site administrator must add ANTHROPIC_API_KEY to the app secrets.")
    if S.calls >= MAX_CALLS:
        raise RuntimeError(f"This session has reached its limit of {MAX_CALLS} AI requests. Reload the page to start a new session.")
    client = anthropic.Anthropic(api_key=key, max_retries=3, timeout=180)
    chosen = MODEL_CHOICES[S.model_label]
    fallbacks = [chosen] + [m for m in MODEL_CHOICES.values() if m != chosen]
    last_error = None
    for model in fallbacks:
        try:
            S.calls += 1
            resp = client.messages.create(model=model, max_tokens=4096,
                                          messages=[{"role": "user", "content": prompt}])
            text = "".join(getattr(b, "text", "") for b in resp.content)
            if not text:
                raise RuntimeError("The AI returned an empty response.")
            return text
        except anthropic.NotFoundError as e:          # model ID not available: try the next one
            last_error = e
            continue
        except anthropic.AuthenticationError:
            raise RuntimeError("The AI service rejected the API key. The site administrator must update ANTHROPIC_API_KEY.")
        except anthropic.RateLimitError:
            raise RuntimeError("The AI service is busy (rate limit reached). Please wait a minute and try again.")
        except anthropic.APIStatusError as e:
            if getattr(e, "status_code", None) == 529:
                raise RuntimeError("The AI service is temporarily overloaded. Please try again shortly.")
            raise RuntimeError(f"AI request failed ({getattr(e, 'status_code', 'error')}). Please try again.")
        except anthropic.APIConnectionError:
            raise RuntimeError("Could not reach the AI service. Please check the connection and try again.")
    raise RuntimeError(f"None of the configured AI models is available ({last_error}). The administrator should update the model IDs in the app secrets.")


def parse_criteria(text, source):
    m = re.search(r"\[[\s\S]*\]", text)
    if not m:
        return []
    try:
        items = json.loads(m.group(0))
    except json.JSONDecodeError:
        log("JSON parse failed")
        return []
    out = []
    for it in items:
        if not isinstance(it, dict):
            continue
        out.append({"id": new_id(), "name": str(it.get("name") or "Unnamed Criterion"),
                    "description": str(it.get("description") or ""), "category": str(it.get("category") or "General"),
                    "reasoning": str(it.get("reasoning") or ""), "frameworks": str(it.get("frameworks") or ""),
                    "source": source, "sourceFile": "", "timestamp": datetime.now().isoformat(timespec="seconds")})
    return out


def dedupe(items):
    seen, out = set(), []
    for c in items:
        k = c["name"].strip().lower()
        if k and k not in seen:
            seen.add(k)
            out.append(c)
    return out


SOURCE_LABEL = {"manual": "Manual Input", "ai": "AI", "pdf": "PDF Extract"}

# ================================================================
# LAYOUT HELPERS
# ================================================================
def header():
    st.markdown("""
    <div class="crest-card">
      <p class="crest-title">AI-Powered MCDM System</p>
      <p class="crest-sub">Phase 1: Criteria Extraction &amp; Generation Tool of CREST</p>
      <div class="crest-about">
        <h4>What This Tool Does:</h4>
        This system helps you create comprehensive criteria sets for multi-criteria decision making by:
        <ul style="margin:6px 0 6px 20px;">
          <li><b>Analyzing research papers:</b> Automatically extracts evaluation criteria from uploaded PDF documents using AI</li>
          <li><b>Generating domain-specific criteria:</b> Creates relevant criteria based on your problem description</li>
          <li><b>Manual input:</b> Allows you to add known criteria and expert knowledge</li>
          <li><b>Smart consolidation:</b> Merges and deduplicates criteria from all sources</li>
          <li><b>Export for validation:</b> Provides structured outputs for expert review and validation</li>
        </ul>
        <h4>What You'll Do:</h4>
        <b>Required steps:</b> 1 → 2 &nbsp;|&nbsp; <b>Optional steps:</b> 3, 4, 5 &nbsp;|&nbsp; <b>Final step:</b> 6<br>
        Follow the numbered tabs below in sequence. At least one optional step (3, 4, or 5) should be completed before generating results.
        <h4>Academic Transparency:</h4>
        All AI-generated criteria are traceable. The system is designed for academic research with full transparency about methodologies and sources.
      </div>
    </div>""", unsafe_allow_html=True)


def tab_bar():
    cols = st.container(key="tabbar").columns(len(STEPS))
    for col, (key, label) in zip(cols, STEPS):
        with col:
            if st.button(label, key=f"tab_{key}", width="stretch",
                         type="primary" if S.step == key else "secondary"):
                S.step = key
                st.rerun()


def go(step):
    S.step = step
    st.rerun()


def nav(buttons):
    cols = st.columns(len(buttons))
    for col, (label, target, primary) in zip(cols, buttons):
        with col:
            if st.button(label, key=f"nav_{S.step}_{target}_{label}", width="stretch",
                         type="primary" if primary else "secondary"):
                go(target)


def criteria_cards(items, kind, removable=True):
    if not items:
        return
    for c in list(items):
        cols = st.columns([12, 1])
        meta = f"Category: {c.get('category') or 'General'}" + (f" | Source: {c['sourceFile']}" if c.get("sourceFile") else "")
        extra = ""
        if c.get("reasoning"):
            extra += f"<div class='crit-desc'><b>Reasoning:</b> {c['reasoning']}</div>"
        if c.get("frameworks"):
            extra += f"<div class='crit-desc'><b>Frameworks:</b> {c['frameworks']}</div>"
        with cols[0]:
            st.markdown(f"<div class='crit'><span class='badge' style='float:right'>{SOURCE_LABEL[kind]}</span>"
                        f"<div class='crit-name'>{c['name']}</div><div class='crit-desc'>{c['description']}</div>"
                        f"<div class='crit-meta'>{meta}</div>{extra}</div>", unsafe_allow_html=True)
        with cols[1]:
            if removable and st.button("✕", key=f"rm_{kind}_{c['id']}", help="Remove this criterion"):
                S[kind] = [x for x in S[kind] if x["id"] != c["id"]]
                st.rerun()


# ================================================================
# STEPS
# ================================================================
def step_setup():
    st.subheader("1. Configure AI (Required)")
    st.markdown("<div class='info-blue'><b>What's Happening Behind the Scenes?</b><br>This system uses AI (by Anthropic) to "
                "analyze PDF documents, generate domain-specific criteria, and provide source transparency. Your documents are "
                "processed securely, and no data is stored permanently.</div>", unsafe_allow_html=True)
    with st.expander("Option 1: Use Provided Research API Key (Recommended, Free for Academic Use)", expanded=True):
        st.markdown("<div class='info-green'><b>For Academic Research: No Setup Required!</b> Completely free, instant access, "
                    "no registration or sign-in needed.</div>", unsafe_allow_html=True)
        st.selectbox("AI Model", list(MODEL_CHOICES), key="model_label")
        if st.button("Use Research API Key", type="primary"):
            S.personal_key = ""
            st.success(f"AI configured! Model: {S.model_label}")
    with st.expander("Option 2: Use Your Own Anthropic API Key (Advanced)"):
        st.markdown("1. Visit [console.anthropic.com](https://console.anthropic.com)  \n2. Create an account and open \"API Keys\"  \n"
                    "3. Generate a key and paste it below. It is used only for your session and is never stored.")
        k = st.text_input("Your Personal API Key", type="password", placeholder="sk-ant-...")
        if st.button("Setup Personal API"):
            if not k.startswith("sk-ant-"):
                st.error("Invalid API key format.")
            else:
                S.personal_key = k
                st.success(f"Personal AI API configured! Model: {S.model_label}")
    if secret("ANTHROPIC_API_KEY") or S.personal_key:
        st.markdown("<div class='info-green'>Research API key is pre-configured! You can proceed directly to step 2 (Domain Setup).</div>",
                    unsafe_allow_html=True)
    else:
        st.warning("The AI service has not been configured by the site administrator yet. Steps 3 and 6 still work; "
                   "AI generation and PDF analysis will be available once the key is added.")
    nav([("Next Phase: 2. Domain Setup →", "domain", True)])


def step_domain():
    st.subheader("2. Define Your Decision Domain (Required)")
    st.write("Describe your specific decision-making problem. This information helps the AI generate relevant criteria tailored to your context.")
    d = S.domain
    name = st.text_input("Domain Name", value=d["name"], placeholder="e.g., 'Healthcare Technology Selection', 'Supplier Evaluation'")
    context = st.text_area("Detailed Context", value=d["context"], height=110,
                           placeholder="Describe your situation, constraints, objectives, and requirements...")
    st.markdown("**Objectives**")
    st.caption(OBJECTIVES_HELP)
    objective = st.text_area("Objectives", value=d["objective"], height=90, label_visibility="collapsed",
                             placeholder="e.g., Minimize environmental impact; Maximize cost efficiency; Ensure system reliability")
    if st.button("Save Domain Setup", type="primary"):
        if not name.strip():
            st.error("Please enter a domain name.")
        else:
            S.domain = {"name": name.strip(), "context": context, "objective": objective}
            st.success(f"Domain saved!\n\nDomain: {name.strip()}\n\nObjectives: {objective or '-'}\n\nReady for criteria generation.")
    st.markdown("<div class='info-yellow'><b>Warning:</b> Please save your domain setup before proceeding to other steps.</div>",
                unsafe_allow_html=True)
    nav([("← Back: 1. AI Setup", "setup", False), ("Next: 3. Manual Addition →", "manual", False),
         ("Next: 4. AI Generation →", "ai", False), ("Next: 5. PDF Analysis →", "pdf", False)])


def step_manual():
    st.subheader("3. Manual Addition of Known Criteria (Optional)")
    st.write("Add criteria you already know are important, from regulations, policies, expert knowledge, or prior experience. Add them one by one.")
    with st.form("manual_form", clear_on_submit=True):
        name = st.text_input("Criterion Name", placeholder="e.g., 'Cost Effectiveness', 'Regulatory Compliance'")
        desc = st.text_area("Description", placeholder="Brief description of this criterion...", height=80)
        c1, c2, _ = st.columns([1, 1, 4])
        add = c1.form_submit_button("Add Criterion", type="primary")
        clear = c2.form_submit_button("Clear All")
    if add:
        if not name.strip():
            st.error("Please enter a criterion name.")
        else:
            S.manual.append({"id": new_id(), "name": name.strip(), "description": desc.strip() or "Manual input",
                             "category": "", "reasoning": "", "frameworks": "", "source": "manual", "sourceFile": "",
                             "timestamp": datetime.now().isoformat(timespec="seconds")})
    if clear:
        S.manual = []
    if S.manual:
        criteria_cards(S.manual, "manual")
    else:
        st.info("No criteria added yet. Add your first criterion above.")
    nav([("← Back: 2. Domain Setup", "domain", False), ("Next: 4. AI Generation →", "ai", False),
         ("Next: 5. PDF Analysis →", "pdf", False), ("Skip to 6. Results →", "results", True)])


def step_ai():
    st.subheader("4. Generate Criteria with AI (Optional)")
    st.write("Let the AI analyze your domain and generate intelligent criteria suggestions based on academic research and best practices.")
    st.markdown("<div class='info-blue'><b>How AI Generation Works:</b><br>The AI uses its training on academic literature and "
                "decision-making frameworks to suggest relevant criteria for your specific domain, considering your context and objectives.</div>",
                unsafe_allow_html=True)
    instructions = st.text_area("Additional Instructions for the AI", height=90,
                                help='Examples: "Must comply with GDPR" | "Budget under $50k" | "Focus on scalability"',
                                placeholder="Any specific requirements, constraints, or focus areas...")
    count = st.slider("Number of Criteria", 1, 50, 10)
    if st.button("Generate with AI", type="primary"):
        if not S.domain["name"]:
            st.error("Please set up and save your domain first (Step 2).")
        else:
            existing = [c["name"] for c in S.manual + S.pdf]
            prompt = (f"You are an expert in MCDM for the {S.domain['name']} domain.\nDOMAIN: {S.domain['name']}\n"
                      f"CONTEXT: {S.domain['context']}\nOBJECTIVES: {S.domain['objective'] or 'Not specified'}\n"
                      f"INSTRUCTIONS: {instructions}\n" + (f"AVOID: {', '.join(existing)}\n" if existing else "") +
                      f"\nGenerate exactly {count} distinct, measurable evaluation criteria. Return ONLY a JSON array:\n"
                      '[{"name":"…","description":"…","category":"…","reasoning":"…","frameworks":"…"}]')
            with st.spinner("The AI is generating criteria…"):
                try:
                    S.ai = parse_criteria(call_ai(prompt), "ai")[:count]
                    st.success(f"Generated {len(S.ai)} criteria using AI.")
                except RuntimeError as e:
                    st.error(str(e))
    if S.ai:
        criteria_cards(S.ai, "ai")
    else:
        st.info('Click "Generate with AI" to create intelligent criteria suggestions.')
    nav([("← Back: 3. Manual Addition", "manual", False), ("Next: 5. PDF Analysis →", "pdf", False),
         ("Skip to 6. Results →", "results", True)])


def extract_pdf_text(data):
    from pypdf import PdfReader
    reader = PdfReader(io.BytesIO(data))
    text = ""
    for i, page in enumerate(reader.pages, 1):
        t = re.sub(r"\s+", " ", page.extract_text() or "").strip()
        if len(t) > 10:
            text += f"\n--- PAGE {i} ---\n{t}\n"
    if len(text) < 100:
        raise RuntimeError("PDF appears to contain very little selectable text (may be scanned).")
    return text


def step_pdf():
    st.subheader("5. Extract Criteria from Research Papers with AI (Optional)")
    st.write("Upload research papers and let the AI extract relevant criteria using advanced reasoning and literature analysis.")
    st.markdown("<div class='info-yellow'><b>Academic Transparency &amp; Source Traceability:</b><ul style='margin:6px 0 4px 20px'>"
                "<li><b>Text extraction:</b> System extracts readable text from PDF documents</li>"
                "<li><b>AI analysis:</b> AI identifies evaluation frameworks, criteria lists, and decision factors in the literature</li>"
                "<li><b>Source attribution:</b> Each extracted criterion is linked to its source document</li></ul>"
                "<b>Note:</b> Works best with text-based PDFs. Scanned documents may require manual review.</div>", unsafe_allow_html=True)
    files = st.file_uploader("Upload Research Papers", type=["pdf"], accept_multiple_files=True)
    focus = st.text_area("Analysis Focus", height=70, placeholder="What should the AI focus on? e.g., 'Look for evaluation criteria for software selection'")
    limit = st.slider("Maximum Criteria per PDF", 1, 50, 10)
    if st.button("Analyze with AI", type="primary"):
        if not S.domain["name"]:
            st.error("Please define your domain first (Step 2).")
        elif not files:
            st.error("Please upload PDF files first.")
        else:
            S.log = ["Ready for PDF analysis..."]
            found, preview = [], ""
            progress = st.progress(0.0)
            for n, f in enumerate(files, 1):
                log(f"Processing {n}/{len(files)}: {f.name}")
                try:
                    text = extract_pdf_text(f.getvalue())
                    preview += f"\n\n=== {f.name} ===\n{text[:3000]}"
                    prompt = (f"You are an expert in MCDM analysis. Extract evaluation criteria from this document for the domain: "
                              f"\"{S.domain['name']}\".\nCONTEXT: {S.domain['context']}\n"
                              + (f"ANALYSIS FOCUS: {focus}\n" if focus.strip() else "") +
                              f"\nDOCUMENT TEXT:\n{text[:PDF_CHARS]}\n\nLook for criteria in tables, evaluation frameworks, decision models, "
                              f"and assessment categories.\nReturn at most {limit} of the most relevant criteria. Return ONLY a JSON array:\n"
                              '[{"name":"…","description":"…","category":"…"}]')
                    with st.spinner(f'Sending "{f.name}" to the AI…'):
                        items = parse_criteria(call_ai(prompt), "pdf")[:limit]
                    for it in items:
                        it["sourceFile"] = f.name
                    found += items
                    log(f"Extracted {len(items)} criteria from {f.name}")
                except RuntimeError as e:
                    log(f"Error: {e}")
                    if "limit" in str(e) or "configured" in str(e) or "rejected" in str(e):
                        st.error(str(e))
                        break
                except Exception as e:  # unreadable PDF and similar
                    log(f"Error: {e}")
                progress.progress(n / len(files))
            S.pdf = dedupe(found)
            S.pdf_preview = preview[:2000] + ("\n\n… [Full text analyzed by AI]" if len(preview) > 2000 else "")
            if S.pdf:
                st.success(f"SUCCESS! Extracted {len(S.pdf)} unique criteria from {len(files)} PDF(s).")
            else:
                st.error("No criteria extracted. Check that the PDFs contain selectable text.")
    if S.pdf:
        criteria_cards(S.pdf, "pdf")
    else:
        st.info('Upload PDF files and click "Analyze with AI" to extract criteria from research literature.')
    st.text_area("Extracted Text Preview", value=S.pdf_preview, height=150, disabled=True,
                 placeholder="Extracted text from PDFs will appear here for verification...")
    st.markdown("**Analysis Log**")
    st.code("\n".join(S.log), language=None)
    nav([("← Back: 4. AI Generation", "ai", False), ("Next: 6. Results & Export →", "results", True)])


def step_results():
    st.subheader("6. Consolidate & Export Final Results")
    st.markdown("Review all criteria, **uncheck any you want to exclude**, then export your final selection.")
    if st.button("Consolidate All Criteria", type="primary"):
        allc = S.manual + S.ai + S.pdf
        if not allc:
            st.warning("No criteria found. Please add criteria in steps 3 to 5 first.")
        else:
            rows = dedupe(allc)
            S.consolidated = pd.DataFrame([{"Include": True, "Name": c["name"], "Description": c["description"],
                                            "Category": c.get("category", ""), "Source": SOURCE_LABEL[c["source"]],
                                            "Source file": c.get("sourceFile", ""), "AI reasoning": c.get("reasoning", ""),
                                            "Related frameworks": c.get("frameworks", ""), "Timestamp": c.get("timestamp", "")}
                                           for c in rows])
    df = S.consolidated
    if df is None:
        st.info('Add criteria from previous tabs, then click "Consolidate All Criteria" to see your results here.')
        nav([("← Back: 5. PDF Analysis", "pdf", False)])
        return
    c1, c2, c3, c4 = st.columns(4)
    for col, num, label in [(c1, len(df), "Total Unique Criteria"), (c2, len(S.manual), "Manual Input"),
                            (c3, len(S.ai), "AI Generated"), (c4, len(S.pdf), "Literature Extracted")]:
        col.markdown(f"<div class='stat'><b>{num}</b>{label}</div>", unsafe_allow_html=True)
    st.write("")
    b1, b2, _ = st.columns([1, 1, 4])
    if b1.button("Select All"):
        S.consolidated["Include"] = True
        st.rerun()
    if b2.button("Deselect All"):
        S.consolidated["Include"] = False
        st.rerun()
    edited = st.data_editor(df, hide_index=True, width="stretch", key="editor",
                            column_config={"Include": st.column_config.CheckboxColumn("Include", width="small")},
                            disabled=[c for c in df.columns if c != "Include"])
    S.consolidated = edited
    sel = edited[edited["Include"]].drop(columns=["Include"]).reset_index(drop=True)
    sel.insert(0, "ID", range(1, len(sel) + 1))
    st.markdown(f"**{len(sel)} of {len(edited)} criteria selected for export**")
    e1, e2, e3 = st.columns(3)
    e1.download_button("Export CSV", sel.to_csv(index=False).encode("utf-8"), "mcdm_criteria.csv", "text/csv",
                       disabled=sel.empty, width="stretch")
    meta = {"exportDate": datetime.now().isoformat(timespec="seconds"), "domain": S.domain,
            "model": MODEL_CHOICES[S.model_label], "totalExported": len(sel)}
    e2.download_button("Export JSON", json.dumps({"metadata": meta, "criteria": sel.to_dict(orient="records")}, indent=2),
                       "mcdm_criteria.json", "application/json", disabled=sel.empty, width="stretch")
    if e3.button("Preview Criteria Pool", width="stretch"):
        lines = [f"Preview: {len(sel)} selected criteria", "", "ID | Name | Category | Source", "-" * 60]
        for _, r in sel.head(10).iterrows():
            lines.append(f"{r['ID']} | {r['Name'][:25]:25} | {str(r['Category'])[:12]:12} | {r['Source']}")
        if len(sel) > 10:
            lines.append(f"… and {len(sel) - 10} more criteria")
        st.code("\n".join(lines), language=None)
    nav([("← Back: 5. PDF Analysis", "pdf", False)])


# ================================================================
# MAIN
# ================================================================
def main():
    header()
    tab_bar()
    with st.container(key="panel"):
        {"setup": step_setup, "domain": step_domain, "manual": step_manual, "ai": step_ai,
         "pdf": step_pdf, "results": step_results}[S.step]()


main()
