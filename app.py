"""
TDS Automation - web app (Streamlit)
=====================================
Give it a BOQ (no model numbers needed), a cover-page template (no placeholders needed)
and your TDS repository. For every BOQ line it
  1. picks the best TDS from the repository (reads the datasheets; optional AI judge),
  2. reads the model no. / make off that datasheet,
  3. reads your cover page, works out which cell is BOQ ref / Description / Make / Model and fills it,
  4. puts the cover page in front of the TDS and highlights the BOQ description + model in it,
and returns ONE zip. You review/adjust every match before the package is built.

Run locally:   streamlit run app.py          (see ONLINE_SETUP.md)
"""

import io
import shutil
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

import cover_autofill
import llm_judge
import sharepoint_source as sp
import tds_automation as tool
import tds_matcher as matcher

st.set_page_config(page_title="TDS Automation", page_icon="📄", layout="wide")
NONE_LABEL = "— none (cover page only) —"


# --------------------------------------------------------------------------
# helpers (no Streamlit calls -> unit-testable)
# --------------------------------------------------------------------------

def secret(name, default=""):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


def save_upload(data, suffix):
    d = Path(tempfile.mkdtemp(prefix="tds_in_"))
    p = d / ("input" + suffix)
    p.write_bytes(data)
    return p


def zip_to_dir(zip_bytes):
    tmp = Path(tempfile.mkdtemp(prefix="tds_db_"))
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        z.extractall(tmp)
    return tmp


def plan_to_frame(plan, library):
    labels = matcher.option_labels(library.docs)
    rows = []
    for e in plan:
        sugg = e["match"]
        rows.append({
            "Use": bool(e["use"]), "S.No": e["sno"], "Description": e["description"],
            "Match": sugg or NONE_LABEL, "Confidence": e["confidence"], "Decided by": e["by"],
            "Model no.": e["model"], "Make": e["make"],
            "Why / other candidates": e["reason"] + ("  |  also: " + " · ".join(c[0] for c in e["candidates"][1:4]) if e["candidates"] else ""),
        })
    return pd.DataFrame(rows), sorted(labels.values())


def apply_edits(plan, edited, library):
    """Turn the user's edited review table back into a plan that run() can build from."""
    label_to_doc = {lab: d for d, lab in matcher.option_labels(library.docs).items()}
    by_sno = {r["S.No"]: r for r in edited.to_dict("records")}
    final = []
    for e in plan:
        r = by_sno[e["sno"]]
        e2 = dict(e)
        chosen = r["Match"] if r["Match"] != NONE_LABEL else ""
        doc = label_to_doc.get(chosen)
        e2["path"] = doc.path if doc is not None else None
        e2["use"] = bool(r["Use"]) and doc is not None
        e2["match"] = chosen
        user_model, user_make = str(r["Model no."] or "").strip(), str(r["Make"] or "").strip()
        if chosen != e["match"] and doc is not None:            # user switched the datasheet
            if user_model == e["model"]:
                user_model = matcher.infer_model(doc.name)
            if user_make == e["make"]:
                user_make = matcher.infer_make(doc.folder) or user_make
            e2["by"] = "you (manual)"
            e2["confidence"] = "Manual"
        e2["model"], e2["make"] = user_model, user_make
        final.append(e2)
    return final


def build_package(boq_path, cover_path, plan, highlight, progress_cb=None):
    work = Path(tempfile.mkdtemp(prefix="tds_job_"))
    out_dir = work / "Output"
    summary = tool.run(str(boq_path), str(cover_path), None, str(out_dir),
                       progress_cb=progress_cb, highlight=highlight, plan=plan)
    zip_path = shutil.make_archive(str(work / "TDS_Output"), "zip", root_dir=str(out_dir))
    return zip_path, summary, out_dir


def clear_repo_cache():
    for k in ("sp_prov", "library", "repo_key", "plan", "plan_sig", "result"):
        st.session_state.pop(k, None)


# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------

st.title("📄 TDS Automation")
st.caption("BOQ (description only) + your cover page + TDS repository → a cover page and the right, "
           "highlighted TDS for every line, zipped.")

left, right = st.columns(2)

# ---- 1. repository ----------------------------------------------------------------
with left:
    st.header("1 · TDS repository")
    source = st.radio("Source", ["Folder on this computer (OneDrive-synced)", "Upload a .zip of the repository",
                                 "SharePoint folder (needs Microsoft sign-in)"])
    if source.startswith("SharePoint"):
        client_id = st.text_input("App (client) ID", value=secret("CLIENT_ID"))
        tenant_id = st.text_input("Tenant ID or domain", value=secret("TENANT_ID", "organizations"))
        folder_url = st.text_input("SharePoint TDS folder link", value=secret("TDS_FOLDER_URL"))
        if "token" not in st.session_state:
            if st.button("Sign in with Microsoft", disabled=not client_id):
                try:
                    msal_app, flow = sp.start_device_login(client_id, tenant_id or "organizations")
                    st.info(f"1) Open **{flow['verification_uri']}**  \n2) Enter code **{flow['user_code']}**")
                    with st.spinner("Waiting for you to finish signing in..."):
                        st.session_state["token"] = sp.finish_device_login(msal_app, flow)
                    st.rerun()
                except Exception as e:
                    st.error(str(e))
        else:
            st.success("Signed in to Microsoft ✔")
    elif source.startswith("Folder"):
        local_dir = st.text_input("Path of the TDS folder on this computer",
                                  placeholder=r"C:\Users\you\OneDrive - Company\TDS Database")
        st.caption("Run the app on your own PC and sync the SharePoint folder with OneDrive (no admin rights needed).")
    else:
        db_zip = st.file_uploader("Repository (.zip of PDFs, subfolders OK)", type=["zip"])
        st.caption("In SharePoint: open the TDS folder → Download (arrives as a .zip).")

# ---- 2. inputs ------------------------------------------------------------------------
with right:
    st.header("2 · Your files")
    boq_file = st.file_uploader("BOQ (.xlsx) - Description (and Make if you have it); no Model no. needed", type=["xlsx"])
    cover_file = st.file_uploader("Cover page template (.docx) - no placeholders needed", type=["docx"])
    if cover_file is not None:
        info = cover_autofill.analyze_template(save_upload(cover_file.getvalue(), ".docx"))
        with st.expander("How I read your cover page", expanded=True):
            if info["recognised"]:
                for label, field in info["recognised"]:
                    st.write(f"✅ **{label}** → fills with the BOQ's **{field}**")
            else:
                st.warning("I couldn't find labels like 'Description', 'Make', 'Model' or 'BOQ Item No.' "
                           "- add them (or use {Description}, {Make}, {Model no.}, {S.No} placeholders).")
            if info["other"]:
                st.write("➖ Left exactly as you typed it: " + ", ".join(f"*{o}*" for o in info["other"]))
    highlight = st.checkbox("Highlight the BOQ description / model in the TDS", value=True)

    with st.expander("AI judge (optional) - makes matching much more reliable"):
        st.caption("Without it, the app ranks datasheets by their content and you confirm the doubtful lines. "
                   "With it, an AI model reads the BOQ line and the shortlisted datasheets (short excerpts only), "
                   "picks the right one - or says none fits - and reads the model/make off the datasheet. "
                   "Needs your own API key; usage is billed/limited by that provider.")
        prov = st.selectbox("Provider", ["None (free ranking only)", "Anthropic", "OpenAI-compatible (other services)"],
                            index={"anthropic": 1, "openai": 2}.get(secret("LLM_PROVIDER", ""), 0))
        api_key = st.text_input("API key", type="password", value=secret("LLM_API_KEY"))
        model_name = st.text_input("Model", value=secret("LLM_MODEL", "claude-haiku-5-5" if prov == "Anthropic" else ""))
        base_url = st.text_input("Base URL (OpenAI-compatible only)", value=secret("LLM_BASE_URL"))
    min_conf = st.select_slider("Pre-tick 'Use' for matches with confidence at least", ["High", "Medium", "Low"], value="Medium")

judge = None
if not prov.startswith("None") and api_key and model_name:
    judge = llm_judge.Judge("anthropic" if prov == "Anthropic" else "openai", api_key, model_name, base_url=base_url)

# ---- 3. analyze --------------------------------------------------------------------------
st.header("3 · Analyze")
ready = bool(boq_file and cover_file)
if source.startswith("SharePoint"):
    ready = ready and "token" in st.session_state and bool(folder_url)
elif source.startswith("Folder"):
    ready = ready and bool(local_dir) and Path(local_dir.strip().strip('"')).is_dir()
else:
    ready = ready and db_zip is not None

b1, b2 = st.columns([1, 5])
if b1.button("Re-read repository"):
    clear_repo_cache()
if b2.button("Analyze BOQ & pick TDS files", type="primary", disabled=not ready):
    try:
        status, bar = st.empty(), st.progress(0.0)
        # repository -> library of readable TDS files
        if source.startswith("SharePoint"):
            key = (folder_url, st.session_state["token"][:20])
            if st.session_state.get("repo_key") != key:
                status.write("Indexing the SharePoint folder...")
                spp = sp.SharePointTDSProvider(st.session_state["token"], folder_url, progress_cb=status.write)
                st.session_state["library"] = spp.build_library(status.write)
                st.session_state["repo_key"] = key
        elif source.startswith("Folder"):
            root = Path(local_dir.strip().strip('"'))
            key = ("dir", str(root))
            if st.session_state.get("repo_key") != key:
                status.write("Reading the TDS folder...")
                st.session_state["library"] = matcher.TDSLibrary.from_dir(root, status.write)
                st.session_state["repo_key"] = key
        else:
            key = ("zip", db_zip.name, db_zip.size)
            if st.session_state.get("repo_key") != key:
                status.write("Reading the uploaded repository...")
                st.session_state["library"] = matcher.TDSLibrary.from_dir(zip_to_dir(db_zip.getvalue()), status.write)
                st.session_state["repo_key"] = key
        library = st.session_state["library"]
        status.write(f"{library.count} TDS files read. Matching the BOQ...")

        boq_path = save_upload(boq_file.getvalue(), ".xlsx")
        cover_path = save_upload(cover_file.getvalue(), ".docx")
        _, items = tool.read_boq(boq_path)

        def on_progress(done, total, msg):
            bar.progress(min(done / max(total, 1), 1.0))
            status.write(msg)

        st.session_state["plan"] = library.plan(items, judge=judge, min_conf=min_conf, progress_cb=on_progress)
        st.session_state["paths"] = (boq_path, cover_path)
        st.session_state.pop("result", None)
        bar.progress(1.0)
        status.write("Done - review the matches below.")
    except Exception as e:
        st.error(f"Something went wrong: {e}")

# ---- 4. review + generate ----------------------------------------------------------------
plan = st.session_state.get("plan")
if plan:
    library = st.session_state["library"]
    st.header("4 · Review & generate")
    n_use = sum(1 for e in plan if e["use"])
    st.caption(f"{n_use} of {len(plan)} lines have a TDS selected. Check the **Match**, tick **Use** for lines you agree with, "
               "and fix **Model no.** / **Make** if needed. Lines without 'Use' get a cover page only.")
    df, options = plan_to_frame(plan, library)
    edited = st.data_editor(
        df, hide_index=True, width="stretch", key="review",
        column_config={
            "Use": st.column_config.CheckboxColumn(width="small"),
            "Match": st.column_config.SelectboxColumn(options=[NONE_LABEL] + options, width="large"),
            "Description": st.column_config.TextColumn(width="medium"),
        },
        disabled=["S.No", "Description", "Confidence", "Decided by", "Why / other candidates"])

    if st.button("Generate package", type="primary"):
        try:
            status, bar = st.empty(), st.progress(0.0)
            final = apply_edits(plan, edited, library)
            boq_path, cover_path = st.session_state["paths"]

            def on_progress(done, total, msg):
                bar.progress(min(done / max(total, 1), 1.0))
                status.write(msg)

            zip_path, summary, out_dir = build_package(boq_path, cover_path, final, highlight, on_progress)
            bar.progress(1.0)
            status.write("Done.")
            st.session_state["result"] = {"zip": zip_path, "summary": summary}
        except Exception as e:
            st.error(f"Something went wrong: {e}")

res = st.session_state.get("result")
if res:
    s = res["summary"]
    st.success(f"{s['items']} lines processed · {s['items'] - s['missing']} with a TDS · {s['missing']} cover page only")
    with open(res["zip"], "rb") as f:
        st.download_button("⬇️ Download all TDS (.zip)", f, file_name="TDS_Output.zip", mime="application/zip", type="primary")
