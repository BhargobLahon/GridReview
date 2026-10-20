import re
import io
import pandas as pd
from openpyxl.styles import PatternFill, Font
REQUIRED_COLUMNS = {"Case", "Element", "Type", "Nominal_kV", "Voltage_pu", "Loading_percent", "Converged"}

GRID_CODE_PRESETS = {
    "Standard Utility (Default)": {
        "v_min_pu": 0.95,
        "v_max_pu": 1.05,
        "v_crit_min_pu": 0.90,
        "v_crit_max_pu": 1.10,
        "v_warn_margin_pu": 0.01,
        "load_warn_pct": 90.0,
        "load_limit_n0_pct": 100.0,
        "load_limit_n1_pct": 110.0,
        "load_crit_pct": 120.0,
    },
    "Strict Transmission (ENTSO-E / IEEE Style)": {
        "v_min_pu": 0.95,
        "v_max_pu": 1.05,
        "v_crit_min_pu": 0.92,
        "v_crit_max_pu": 1.08,
        "v_warn_margin_pu": 0.015,
        "load_warn_pct": 85.0,
        "load_limit_n0_pct": 100.0,
        "load_limit_n1_pct": 105.0,
        "load_crit_pct": 115.0,
    }
}

def parse_nl_config(prompt_text: str, current_config: dict) -> tuple[dict, list[str]]:
    """Deterministic NL-to-JSON intent parser (Agent Layer 1)."""
    updated = current_config.copy()
    trace = []
    text = prompt_text.lower()

    # Extract voltage range like "0.94 to 1.06" or "0.95 - 1.05"
    v_range = re.search(r"(0\.\d+)\s*(?:to|-|and)\s*(1\.\d+)", text)
    if v_range:
        updated["v_min_pu"] = float(v_range.group(1))
        updated["v_max_pu"] = float(v_range.group(2))
        trace.append(f"Extracted voltage window: [{updated['v_min_pu']}, {updated['v_max_pu']}] pu")

    # Extract warning loading like "above 85% as a warning" or "warn at 90%"
    warn_match = re.search(r"(?:warn\w*.*?(\d{2,3})\s*%)|(?:(\d{2,3})\s*%.*?warn)", text)
    if warn_match:
        val = warn_match.group(1) or warn_match.group(2)
        updated["load_warn_pct"] = float(val)
        trace.append(f"Updated loading warning threshold: {updated['load_warn_pct']}%")

    # Extract violation loading like "above 100% as a violation" or "limit of 105%"
    viol_match = re.search(r"(?:(?:violation|limit).*?(\d{2,3})\s*%)|(?:(\d{2,3})\s*%.*?(?:violation|limit))", text)
    if viol_match:
        val = viol_match.group(1) or viol_match.group(2)
        updated["load_limit_n0_pct"] = float(val)
        trace.append(f"Updated N-0 loading limit: {updated['load_limit_n0_pct']}%")

    if not trace:
        trace.append("No new numerical overrides detected in prompt; retained active thresholds.")

    return updated, trace

def is_contingency_case(case_name: str) -> bool:
    keywords = ["outage", "n-1", "contingency", "trip", "fault", "loss"]
    return any(k in str(case_name).lower() for k in keywords)

def evaluate_row(row, cfg: dict, differentiate_n1: bool = True):
    case = str(row["Case"])
    element = str(row["Element"])
    el_type = str(row["Type"]).strip().lower()
    conv = str(row["Converged"]).strip().lower()
    is_n1 = is_contingency_case(case)
    regime = "N-1 Contingency" if is_n1 else "N-0 Normal"

    # 1. Short-circuit non-convergent cases first
    if conv in ["no", "false", "0"]:
        return pd.Series([
            regime, "CRITICAL", 1, "Non-Convergent Simulation",
            "Power flow failed to converge; downstream bus/branch states invalid.",
            "Inspect reactive power (Q) limits, slack bus adequacy, or severe islanding under this outage."
        ])

    v = pd.to_numeric(row["Voltage_pu"], errors="coerce")
    loading = pd.to_numeric(row["Loading_percent"], errors="coerce")

    active_load_limit = cfg["load_limit_n1_pct"] if (differentiate_n1 and is_n1) else cfg["load_limit_n0_pct"]

    # 2. Check Bus Voltages
    if el_type == "bus":
        if pd.isna(v):
            return pd.Series([regime, "WARNING", 3, "Missing Voltage Value", "Bus row has NaN Voltage_pu.", "Verify parser export mapping."])
        if v < cfg["v_crit_min_pu"]:
            return pd.Series([regime, "CRITICAL", 1, "Severe Undervoltage", f"{v:.3f} pu < Critical {cfg['v_crit_min_pu']:.2f} pu", "Check local Q-support, shunt capacitor bank status, or nearby generator AVR clamping."])
        if v > cfg["v_crit_max_pu"]:
            return pd.Series([regime, "CRITICAL", 1, "Severe Overvoltage", f"{v:.3f} pu > Critical {cfg['v_crit_max_pu']:.2f} pu", "Check Ferranti effect on lightly loaded lines or switch in shunt reactors."])
        if v < cfg["v_min_pu"]:
            return pd.Series([regime, "VIOLATION", 2, "Undervoltage", f"{v:.3f} pu < Min {cfg['v_min_pu']:.2f} pu", "Candidate check: ULTC transformer tap adjustment or local VAr compensation."])
        if v > cfg["v_max_pu"]:
            return pd.Series([regime, "VIOLATION", 2, "Overvoltage", f"{v:.3f} pu > Max {cfg['v_max_pu']:.2f} pu", "Candidate check: Generator scheduled voltage setpoint or reactor tap."])
        if v < (cfg["v_min_pu"] + cfg["v_warn_margin_pu"]) or v > (cfg["v_max_pu"] - cfg["v_warn_margin_pu"]):
            return pd.Series([regime, "WARNING", 3, "Near Voltage Limit", f"{v:.3f} pu within {cfg['v_warn_margin_pu']} pu margin", "Monitor under slightly heavier load/generation sensitivity variations."])
        return pd.Series([regime, "PASS", 4, "Normal Voltage", f"{v:.3f} pu within [{cfg['v_min_pu']}, {cfg['v_max_pu']}] pu", "None required."])

    # 3. Check Equipment Loading (Lines / Transformers)
    if el_type in ["line", "transformer", "branch"]:
        if pd.isna(loading):
            return pd.Series([regime, "WARNING", 3, "Missing Loading Value", "Branch row has NaN Loading_percent.", "Verify thermal rating base in export file."])
        if loading > cfg["load_crit_pct"]:
            return pd.Series([regime, "CRITICAL", 1, "Critical Overload", f"{loading:.1f}% > Critical {cfg['load_crit_pct']:.0f}%", "Immediate thermal limit breach; evaluate generation redispatch or topology split."])
        if loading > active_load_limit:
            return pd.Series([regime, "VIOLATION", 2, "Overload", f"{loading:.1f}% > {regime} Limit ({active_load_limit:.0f}%)", "Verify short-term emergency thermal rating (STE) or phase-shifting / parallel path relief."])
        if loading >= cfg["load_warn_pct"]:
            return pd.Series([regime, "WARNING", 3, "High Loading Warning", f"{loading:.1f}% >= Warning {cfg['load_warn_pct']:.0f}%", "Approaching thermal ceiling; flag for peak-season sensitivity check."])
        return pd.Series([regime, "PASS", 4, "Normal Loading", f"{loading:.1f}% <= {cfg['load_warn_pct']:.0f}%", "None required."])

    return pd.Series([regime, "PASS", 4, "System Converged", "Converged", "None required."])

def run_study_review(df: pd.DataFrame, cfg: dict, differentiate_n1: bool = True) -> dict:
    # Handle files where row 1 is the compliance banner
    if not REQUIRED_COLUMNS.issubset(set(df.columns)):
        raise ValueError(f"Missing required columns: {REQUIRED_COLUMNS - set(df.columns)}")

    out = df.copy()
    eval_cols = ["Regime", "Severity", "Rank_Order", "Classification", "Finding_Detail", "Engineering_Triage_Note"]
    out[eval_cols] = out.apply(evaluate_row, cfg=cfg, differentiate_n1=differentiate_n1, axis=1)
    out = out.sort_values(by=["Rank_Order", "Case"]).reset_index(drop=True)

    total_cases = out["Case"].nunique()
    non_conv_cases = out[out["Classification"] == "Non-Convergent Simulation"]["Case"].nunique()
    conv_cases = total_cases - non_conv_cases
    v_viols = len(out[out["Classification"].isin(["Undervoltage", "Overvoltage", "Severe Undervoltage", "Severe Overvoltage"])])
    overloads = len(out[out["Classification"].isin(["Overload", "Critical Overload"])])

    flagged = out[out["Severity"].isin(["CRITICAL", "VIOLATION"])]

    # Build text report matching BizChat spec
    lines = [
        "POWER SYSTEM STUDY REVIEW",
        "=========================",
        f"Distinct cases analyzed : {total_cases}",
        f"Converged cases         : {conv_cases}",
        f"Non-converged cases     : {non_conv_cases}",
        f"Voltage violations      : {v_viols}",
        f"Equipment overloads     : {overloads}",
        "",
        "CRITICAL & VIOLATION FINDINGS:"
    ]
    if flagged.empty:
        lines.append("  None. All analyzed elements are within configured limits.")
    else:
        for idx, (_, r) in enumerate(flagged.iterrows(), 1):
            lines.append(f"{idx}. {r['Case']} [{r['Regime']}] - {r['Element']} ({r['Type']})")
            lines.append(f"   Severity       : {r['Severity']} ({r['Classification']})")
            lines.append(f"   Detail         : {r['Finding_Detail']}")
            lines.append(f"   Triage Note    : {r['Engineering_Triage_Note']}")

    lines.append("")
    lines.append("RECOMMENDED ACTION:")
    if len(flagged) > 0:
        lines.append(f"Review the {len(flagged)} flagged condition(s) above before accepting the study results.")
    else:
        lines.append("Study passes automated screening criteria.")

    return {
        "reviewed_df": out,
        "summary_metrics": {
            "total_cases": total_cases,
            "converged_cases": conv_cases,
            "non_converged_cases": non_conv_cases,
            "voltage_violations": v_viols,
            "equipment_overloads": overloads,
        },
        "text_report": "\n".join(lines)
    }

def export_reviewed_excel(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Reviewed_Results", startrow=2)
        ws = writer.sheets["Reviewed_Results"]
        ws["A1"] = "AUTOMATED STUDY REVIEW OUTPUT - SYNTHETIC PROTOTYPE DATA ONLY"
        ws["A1"].font = Font(bold=True, color="1F4E79")

        fills = {
            "CRITICAL": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
            "VIOLATION": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
            "WARNING": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
            "PASS": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
        }

        sev_col_idx = list(df.columns).index("Severity") + 1
        for row_idx in range(4, len(df) + 4):
            sev_val = ws.cell(row=row_idx, column=sev_col_idx).value
            if sev_val in fills:
                for col_idx in range(1, len(df.columns) + 1):
                    ws.cell(row=row_idx, column=col_idx).fill = fills[sev_val]

    return buffer.getvalue()