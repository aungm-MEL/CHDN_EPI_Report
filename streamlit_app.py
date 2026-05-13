"""
Streamlit app for CHDN EPI data transformation pipeline.
Runs CHDN_EPI_clean.py followed by chdn_epi_master.py
"""

import subprocess
import sys
from pathlib import Path
import time
import streamlit as st

# Page config
st.set_page_config(page_title="CHDN EPI Pipeline", layout="wide", initial_sidebar_state="expanded")

# Paths
SCRIPT_DIR = Path(__file__).resolve().parent
CHDN_CLEAN_SCRIPT = SCRIPT_DIR / "CHDN_clean" / "CHDN_EPI_clean.py"
CHDN_MASTER_SCRIPT = SCRIPT_DIR / "chdn_epi_master.py"
DATA_FILE = SCRIPT_DIR / "EPI Database_CHDN.xlsx"
CLEAN_OUTPUT = SCRIPT_DIR / "CHDN_clean" / "CHDN_EPI_clean.xlsx"
MASTER_OUTPUT = SCRIPT_DIR / "CHDN dataset_long.xlsx"
VENV_PYTHON = Path.cwd().parent.parent / ".venv" / "Scripts" / "python.exe"

# Session state
if "step_cleaner_done" not in st.session_state:
    st.session_state.step_cleaner_done = False
if "step_master_done" not in st.session_state:
    st.session_state.step_master_done = False
if "cleaner_log" not in st.session_state:
    st.session_state.cleaner_log = ""
if "master_log" not in st.session_state:
    st.session_state.master_log = ""

st.title("🔄 CHDN EPI Data Transformation Pipeline")

# Sidebar
with st.sidebar:
    st.header("📂 Dataset Upload")
    
    uploaded_file = st.file_uploader(
        "Upload EPI Database file (xlsx/xlsm)",
        type=["xlsx", "xlsm"],
        help="Upload your EPI Database_CHDN.xlsx file"
    )
    
    if uploaded_file is not None:
        with st.spinner("Saving uploaded file..."):
            # Save uploaded file to the script directory
            with open(DATA_FILE, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.success(f"✓ File uploaded: {uploaded_file.name}")
            st.info(f"Saved to: {DATA_FILE}")
    
    st.divider()
    st.header("Pipeline Info")
    st.info(
        f"""
        **Input Dataset:** EPI Database_CHDN.xlsx
        
        **Processing Steps:**
        1. Run CHDN_EPI_clean.py (validate, clean, compute completions)
        2. Run chdn_epi_master.py (transform to long format, generate reports)
        
        **Output Files:**
        - CHDN_EPI_clean.xlsx (cleaned child & pregnancy sheets)
        - CHDN dataset_long.xlsx (all analysis sheets)
        """
    )
    st.divider()
    st.subheader("File Status")
    col1, col2 = st.columns(2)
    with col1:
        if DATA_FILE.exists():
            file_size = DATA_FILE.stat().st_size / (1024 * 1024)  # Convert to MB
            st.success(f"✓ Input dataset ({file_size:.1f} MB)")
        else:
            st.error(f"✗ Input dataset NOT found")
    with col2:
        if CLEAN_OUTPUT.exists():
            st.success(f"✓ Clean output exists")
        else:
            st.info(f"⊘ Clean output (will be created)")

# Main area
st.subheader("🚀 Run Data Pipeline")
st.write("Clean dataset → Transform to long format → Generate reports")

if st.button("▶ Run Full Pipeline", key="btn_pipeline", use_container_width=True, type="primary"):
    if not DATA_FILE.exists():
        st.error(f"❌ Input file not found: {DATA_FILE}\n\nPlease upload a file in the sidebar first.")
    elif not CHDN_CLEAN_SCRIPT.exists():
        st.error(f"❌ Script not found: {CHDN_CLEAN_SCRIPT}")
    elif not CHDN_MASTER_SCRIPT.exists():
        st.error(f"❌ Script not found: {CHDN_MASTER_SCRIPT}")
    else:
        progress_bar = st.progress(0, text="Initializing...")
        log_container = st.container()
        
        try:
            # Step 1: Run Cleaner
            with log_container:
                with st.spinner("🔄 Step 1/2: Running cleaner..."):
                    progress_bar.progress(25, text="Step 1/2: Cleaning dataset...")
                    
                    result = subprocess.run(
                        [str(VENV_PYTHON), str(CHDN_CLEAN_SCRIPT)],
                        cwd=SCRIPT_DIR,
                        capture_output=True,
                        text=True,
                        timeout=300
                    )
                    
                    st.session_state.cleaner_log = result.stdout + "\n" + result.stderr
                    
                    if result.returncode != 0:
                        st.error("❌ Cleaner failed!")
                        with st.expander("📋 View Error Output"):
                            st.code(st.session_state.cleaner_log, language="text")
                        st.stop()
                    
                    st.session_state.step_cleaner_done = True
                    st.success("✓ Step 1 Complete: Dataset cleaned successfully!")
            
            # Step 2: Run Master Transform
            with log_container:
                with st.spinner("🔄 Step 2/2: Running master transform..."):
                    progress_bar.progress(75, text="Step 2/2: Generating reports...")
                    
                    result = subprocess.run(
                        [str(VENV_PYTHON), str(CHDN_MASTER_SCRIPT)],
                        cwd=SCRIPT_DIR,
                        capture_output=True,
                        text=True,
                        timeout=600
                    )
                    
                    st.session_state.master_log = result.stdout + "\n" + result.stderr
                    
                    if result.returncode != 0:
                        st.error("❌ Master transform failed!")
                        with st.expander("📋 View Error Output"):
                            st.code(st.session_state.master_log, language="text")
                        st.stop()
                    
                    st.session_state.step_master_done = True
                    st.success("✓ Step 2 Complete: Reports generated successfully!")
            
            # Final success
            progress_bar.progress(100, text="✓ Pipeline Complete!")
            st.balloons()
            st.success("🎉 **Pipeline completed successfully!**")
            
            with st.expander("📋 View Full Output Logs"):
                st.subheader("Cleaner Output:")
                st.code(st.session_state.cleaner_log, language="text")
                st.divider()
                st.subheader("Master Transform Output:")
                st.code(st.session_state.master_log, language="text")
        
        except subprocess.TimeoutExpired:
            progress_bar.progress(0, text="❌ Timeout!")
            st.error("⏱️ Timeout: Script took too long (>10 minutes)")
        except Exception as e:
            progress_bar.progress(0, text="❌ Error!")
            st.error(f"❌ Error running pipeline: {str(e)}")

# Download section
st.divider()
st.subheader("📥 Download Output Files")

col1, col2 = st.columns(2)

with col1:
    if CLEAN_OUTPUT.exists():
        file_size = CLEAN_OUTPUT.stat().st_size / (1024 * 1024)
        st.write(f"**Cleaned Dataset** ({file_size:.1f} MB)")
        with open(CLEAN_OUTPUT, "rb") as f:
            st.download_button(
                label="⬇ CHDN_EPI_clean.xlsx",
                data=f.read(),
                file_name="CHDN_EPI_clean.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.info("🔲 Clean file will be created after running pipeline")

with col2:
    if MASTER_OUTPUT.exists():
        file_size = MASTER_OUTPUT.stat().st_size / (1024 * 1024)
        st.write(f"**Analysis Report** ({file_size:.1f} MB)")
        with open(MASTER_OUTPUT, "rb") as f:
            st.download_button(
                label="⬇ CHDN dataset_long.xlsx",
                data=f.read(),
                file_name="CHDN dataset_long.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.info("🔲 Report file will be created after running pipeline")

# Summary
st.divider()
st.subheader("📊 Pipeline Status")

if st.session_state.step_cleaner_done and st.session_state.step_master_done:
    st.success("✅ **Pipeline Complete!** Both output files are ready to download.")
elif st.session_state.step_cleaner_done:
    st.info("⏳ **In Progress:** Step 1 complete, waiting for Step 2...")
elif st.session_state.cleaner_log or st.session_state.master_log:
    st.warning("⚠️ **Pipeline Paused** — Check the output logs above for errors")
else:
    st.info("⊘ **Ready to run** — Upload a dataset and click 'Run Full Pipeline' to begin")

# Footer
st.divider()
st.caption(f"Data folder: {SCRIPT_DIR}")
