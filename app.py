import os
import streamlit as st
import pandas as pd
from openpyxl.styles import Font, PatternFill
from review_engine import (
    GRID_CODE_PRESETS,
    parse_nl_config,
    run_study_review,
    export_reviewed_excel,
    REQUIRED_COLUMNS,
)

def ensure_sample_workbook(filename="sample_results.xlsx"):
    """Self-contained synthetic workbook generator so app.py has zero external script dependencies."""
    data = [
        ["Base Case", "Bus 1 (Slack)", "Bus", 1.04, None, "Yes"],
        ["Base Case", "Bus 2 (PV)", "Bus", 1.02, None, "Yes"],
        ["Base Case", "Line L1-2", "Line", None, 72.4, "Yes"],
        ["Peak Load", "Bus 7 (PQ)", "Bus", 0.93, None, "Yes"],
        ["Peak Load", "Bus 14 (PQ)", "Bus", 0.89, None, "Yes"],
        ["Peak Load", "Transformer T1", "Transformer", None, 96.5, "Yes"],
        ["Line 2 Outage (N-1)", "Transformer T2", "Transformer", None, 112.0, "Yes"],
        ["Line 2 Outage (N-1)", "Line L4-5", "Line", None, 124.5, "Yes"],
        ["Line 2 Outage (N-1)", "Bus 9 (PQ)", "Bus", 0.94, None, "Yes"],
        ["Line 5 Outage (N-1)", "System", "System", None, None, "No"],
        ["Light Load", "Bus 4 (PQ)", "Bus", 1.06, None, "Yes"],
    ]
    columns = ["Case", "Element", "Type", "Voltage_pu", "Loading_percent", "Converged"]
    df = pd.DataFrame(data, columns=columns)

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Study_Results", startrow=2)
        ws = writer.sheets["Study_Results"]
        ws["A1"] = "NOTE: Synthetic data created exclusively for prototype testing. No proprietary or internal network data is used."
        ws["A1"].font = Font(bold=True, italic=True, color="9C0006")
        ws["A1"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

st.set_page_config(page_title="GridReview AI", layout="wide")

st.title("GridReview: Automated Power-System Study Results Reviewer")
st.caption("Hybrid Architecture: Natural-Language Intent Parsing + Deterministic N-0/N-1 Engineering Validation")

# Sidebar Configuration
with st.sidebar:
    st.header("1. Grid Code Preset")
    preset_name = st.selectbox("Select Criteria Profile", list(GRID_CODE_PRESETS.keys()))
    if "active_cfg" not in st.session_state or st.session_state.get("last_preset") != preset_name:
        st.session_state.active_cfg = GRID_CODE_PRESETS[preset_name].copy()
        st.session_state.last_preset = preset_name

    differentiate_n1 = st.checkbox("Apply separate N-1 Emergency Loading Limit", value=True)

    st.header("2. Editable Thresholds")
    cfg = st.session_state.active_cfg
    cfg["v_min_pu"] = st.number_input("Min Voltage Limit (pu)", 0.80, 1.00, float(cfg["v_min_pu"]), 0.01)
    cfg["v_max_pu"] = st.number_input("Max Voltage Limit (pu)", 1.00, 1.20, float(cfg["v_max_pu"]), 0.01)
    cfg["v_crit_min_pu"] = st.number_input("Critical Undervoltage (pu)", 0.70, 0.95, float(cfg["v_crit_min_pu"]), 0.01)
    cfg["load_warn_pct"] = st.number_input("Loading Warning (%)", 50.0, 100.0, float(cfg["load_warn_pct"]), 5.0)
    cfg["load_limit_n0_pct"] = st.number_input("N-0 Loading Violation (%)", 80.0, 120.0, float(cfg["load_limit_n0_pct"]), 5.0)
    if differentiate_n1:
        cfg["load_limit_n1_pct"] = st.number_input("N-1 Emergency Loading Limit (%)", 90.0, 130.0, float(cfg["load_limit_n1_pct"]), 5.0)
    cfg["load_crit_pct"] = st.number_input("Critical Loading (%)", 100.0, 150.0, float(cfg["load_crit_pct"]), 5.0)

# AI Natural Language Configuration Bar
with st.expander("AI Assistant: Natural-Language Threshold Configuration (Human-in-the-Loop)", expanded=True):
    nl_prompt = st.text_input(
        "Instruct the agent in plain English:",
        value="Review this study using voltage limits of 0.95 to 1.05 pu. Treat loading above 100% as a violation and loading above 90% as a warning."
    )
    if st.button("Parse Instruction & Update Thresholds"):
        new_cfg, trace_notes = parse_nl_config(nl_prompt, st.session_state.active_cfg)
        st.session_state.active_cfg = new_cfg
        st.session_state.nl_trace = trace_notes
        st.rerun()

    if "nl_trace" in st.session_state:
        st.json(st.session_state.active_cfg)

# File Upload or Built-in Synthetic Sample
col_up, col_demo = st.columns([2, 1])
with col_up:
    uploaded_file = st.file_uploader("Upload Study Results (.xlsx or .csv)", type=["xlsx", "csv"])
with col_demo:
    st.write("Quick Demo:")
    use_sample = st.button("Load Synthetic IEEE N-1 Sample Workbook")

df_input = None
if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df_input = pd.read_csv(uploaded_file)
    else:
        df_temp = pd.read_excel(uploaded_file)
        if not REQUIRED_COLUMNS.issubset(set(df_temp.columns)):
            df_input = pd.read_excel(uploaded_file, skiprows=2)
        else:
            df_input = df_temp
elif use_sample or st.session_state.get("sample_loaded"):
    st.session_state.sample_loaded = True
    if not os.path.exists("sample_results.xlsx"):
        ensure_sample_workbook("sample_results.xlsx")
    df_temp = pd.read_excel("sample_results.xlsx")
    if not REQUIRED_COLUMNS.issubset(set(df_temp.columns)):
        df_input = pd.read_excel("sample_results.xlsx", skiprows=2)
    else:
        df_input = df_temp

if df_input is not None:
    try:
        results = run_study_review(df_input, st.session_state.active_cfg, differentiate_n1)
        m = results["summary_metrics"]
        rev_df = results["reviewed_df"]

        st.divider()
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Cases Reviewed", m["total_cases"])
        c2.metric("Converged Cases", m["converged_cases"])
        c3.metric("Non-Convergent", m["non_converged_cases"])
        c4.metric("Voltage Violations", m["voltage_violations"])
        c5.metric("Equipment Overloads", m["equipment_overloads"])

        with st.expander("Agent Execution Trace (Deterministic Verification Log)", expanded=False):
            st.markdown(f"""
            1. **[Schema Guard]** Verified required columns `{list(REQUIRED_COLUMNS)}` across **{len(df_input)}** element rows.
            2. **[Contingency Classifier]** Tagged N-0 Normal vs. N-1 Contingency regimes (`differentiate_n1={differentiate_n1}`).
            3. **[Convergence Short-Circuit]** Isolated **{m['non_converged_cases']}** non-convergent case(s) prior to floating-point limit checks.
            4. **[Severity Auditor]** Ranked findings into `CRITICAL`, `VIOLATION`, `WARNING`, and `PASS`.
            """)

        tab1, tab2, tab3 = st.tabs(["Ranked Findings (All)", "Critical & Violations Only", "Automated Engineering Report"])
        with tab1:
            st.dataframe(rev_df, use_container_width=True)
        with tab2:
            st.dataframe(rev_df[rev_df["Severity"].isin(["CRITICAL", "VIOLATION", "WARNING"])], use_container_width=True)
        with tab3:
            st.code(results["text_report"], language="text")

        dl1, dl2 = st.columns(2)
        with dl1:
            excel_bytes = export_reviewed_excel(rev_df)
            st.download_button(
                "Download Color-Coded Reviewed Excel (.xlsx)",
                data=excel_bytes,
                file_name="GridReview_Audited_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with dl2:
            st.download_button(
                "Download Executive Summary (.txt)",
                data=results["text_report"],
                file_name="GridReview_Executive_Summary.txt",
                mime="text/plain"
            )
    except Exception as e:
        st.error(f"Validation Error: {e}")