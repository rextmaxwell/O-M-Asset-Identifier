import os
import tempfile
import time
import pandas as pd
import streamlit as st

from matcher import match_files_to_assets, save_results_csv

st.set_page_config(page_title="O&M Matcher", layout="wide")
st.title("O&M Matcher — Link O&M Manuals to Assets")

st.markdown("""
Upload your **assets list** (CSV/Excel) and your **O&M files** (PDF/DOCX).  
The app extracts identifiers and proposes top matches with a confidence score.
""")

# --- Sidebar options ---
with st.sidebar:
    st.header("Options")
    compute_hash = st.checkbox("Compute file hashes (slower, more accurate)", value=False)
    auto_confirm_threshold = st.slider("Auto-confirm threshold", 60, 100, 80)
    st.caption("Candidates scoring ≥ threshold will be auto-selected.")

# --- Upload assets ---
assets_file = st.file_uploader("Upload assets list (CSV/Excel)", type=["csv", "xlsx"])
assets_df = None
if assets_file is not None:
    if assets_file.name.lower().endswith(".csv"):
        assets_df = pd.read_csv(assets_file)
    else:
        assets_df = pd.read_excel(assets_file)
    st.success(f"Loaded {len(assets_df)} assets.")
    st.dataframe(assets_df.head())

# --- Upload O&M files ---
om_files = st.file_uploader("Upload O&M files (PDF/DOCX)", type=["pdf", "docx"], accept_multiple_files=True)

if assets_df is not None and om_files:
    # Save uploads to temp files so extractors can read them
    temp_paths = []
    for f in om_files:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f.name)[1])
        tmp.write(f.read())
        tmp.flush(); tmp.close()
        temp_paths.append(tmp.name)

    with st.spinner("Matching O&M files to assets..."):
        results = match_files_to_assets(temp_paths, assets_df, compute_hash=compute_hash)

    # Store confirmations
    if "confirmed" not in st.session_state:
        st.session_state.confirmed = []

    # Display each file’s candidates
    for r in results:
        st.subheader(os.path.basename(r["file_path"]))
        col_left, col_right = st.columns([2, 1])

        with col_left:
            if "signals" in r:
                st.markdown("**Extracted identifiers**")
                st.write({
                    "asset_ids": r["signals"].get("asset_ids", []),
                    "serials": r["signals"].get("serials", []),
                    "models": r["signals"].get("models", []),
                    "manufacturers": r["signals"].get("manufacturers", []),
                })
            if r.get("error"):
                st.error(r["error"])

            candidates_df = pd.DataFrame(r.get("top_candidates", []))
            if not candidates_df.empty:
                st.dataframe(candidates_df)

        with col_right:
            options = ["None"] + [f"{c['asset_id']} — {c['name']} (score {c['score']})" for c in r.get("top_candidates", [])]
            default_index = 0
            if r.get("auto_choice"):
                # Pick the one with score ≥ threshold
                top = r["auto_choice"]
                if top["score"] >= auto_confirm_threshold:
                    # Find its index in options
                    for idx, o in enumerate(options):
                        if o.startswith(str(top["asset_id"]) + " —"):
                            default_index = idx
                            break

            choice = st.selectbox("Confirm match:", options, index=default_index, key=r["file_path"])
            note = st.text_input("Notes (optional)", key=r["file_path"] + "_note")

            if st.button("Save", key=r["file_path"] + "_save"):
                chosen_asset_id = None
                if choice != "None":
                    chosen_asset_id = choice.split(" — ")[0]
                st.session_state.confirmed.append({
                    "file_path": r["file_path"],
                    "chosen_asset_id": chosen_asset_id,
                    "note": note,
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "auto_chosen": bool(r.get("auto_choice") and choice != "None")
                })
                st.success("Saved!")

    # Export section
    st.divider()
    st.subheader("Export")
    if st.session_state.confirmed:
        export_df = pd.DataFrame(st.session_state.confirmed)
        st.dataframe(export_df)
        if st.button("Download CSV of confirmations"):
            out = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
            export_df.to_csv(out.name, index=False)
            st.download_button("Download", data=open(out.name, "rb").read(), file_name="om_matches.csv")
    else:
        st.info("No confirmations yet.")

else:
    st.info("Upload both assets list and O&M files to start matching.")
