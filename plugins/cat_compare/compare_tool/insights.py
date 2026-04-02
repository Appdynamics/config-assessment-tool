"""
insights.py
-----------
This module provides functionality for generating insights from comparison data.

Purpose:
- Implements the `build_comparison_json` function to create JSON data for insights.
- Interacts with historical data stored in the `history` folder.

Key Features:
- Generates JSON data for use in the insights page.
"""

from __future__ import annotations

import os
import json
import datetime as dt
from typing import Optional, Dict, Any, Tuple, List

import pandas as pd
from openpyxl import load_workbook


def build_comparison_json(
    domain: str,
    comparison_result_path: str,
    current_file_path: str,
    previous_file_path: str,
    result_folder: str,
    meta: Optional[Dict[str, Any]] = None,
) -> Tuple[str, str, Dict[str, Any]]:
    """
    Build the JSON snapshot used by the Insights UI.
    """
    domain = domain.upper()
    result_folder = result_folder or "."

    # ---------- Base data from Analysis sheet ----------
    df_analysis = pd.read_excel(comparison_result_path, sheet_name="Analysis")

    AREA_MAP: Dict[str, List[str]] = {
        "APM": [
            "AppAgentsAPM",
            "MachineAgentsAPM",
            "BusinessTransactionsAPM",
            "BackendsAPM",
            "OverheadAPM",
            "ServiceEndpointsAPM",
            "ErrorConfigurationAPM",
            "HealthRulesAndAlertingAPM",
            "DataCollectorsAPM",
            "DashboardsAPM",
        ],
        "BRUM": ["NetworkRequestsBRUM", "HealthRulesAndAlertingBRUM"],
        "MRUM": ["NetworkRequestsMRUM", "HealthRulesAndAlertingMRUM"],
    }

    DETAIL_SHEETS: Dict[str, Dict[str, str]] = {
        "APM": {
            "AppAgentsAPM": "AppAgentsAPM",
            "MachineAgentsAPM": "MachineAgentsAPM",
            "BusinessTransactionsAPM": "BusinessTransactionsAPM",
            "BackendsAPM": "BackendsAPM",
            "OverheadAPM": "OverheadAPM",
            "ServiceEndpointsAPM": "ServiceEndpointsAPM",
            "ErrorConfigurationAPM": "ErrorConfigurationAPM",
            "HealthRulesAndAlertingAPM": "HealthRulesAndAlertingAPM",
            "DataCollectorsAPM": "DataCollectorsAPM",
            "DashboardsAPM": "DashboardsAPM",
        },
        "BRUM": {
            "NetworkRequestsBRUM": "NetworkRequestsBRUM",
            "HealthRulesAndAlertingBRUM": "HealthRulesAndAlertingBRUM",
        },
        "MRUM": {
            "NetworkRequestsMRUM": "NetworkRequestsMRUM",
            "HealthRulesAndAlertingMRUM": "HealthRulesAndAlertingMRUM",
        },
    }

    areas = AREA_MAP.get(domain, [])
    app_total = int(df_analysis["name"].dropna().astype(str).str.strip().ne("").sum())

    # ---------- Overall upgrade/degrade ----------
    def count_changes(df: pd.DataFrame, col: str):
        if col not in df.columns:
            return 0, 0, [], []
        s = df[col].astype(str)
        improved = s.str.contains("Upgraded", case=False, na=False)
        degraded = s.str.contains("Downgraded", case=False, na=False)
        imp_names = df.loc[improved, "name"].astype(str).str.strip().tolist()
        deg_names = df.loc[degraded, "name"].astype(str).str.strip().tolist()
        return improved.sum(), degraded.sum(), imp_names, deg_names

    overall_imp, overall_deg, _, _ = count_changes(df_analysis, "OverallAssessment")

    if overall_imp > overall_deg:
        overall_result = "Increase"
    elif overall_deg > overall_imp:
        overall_result = "Decrease"
    else:
        overall_result = "Even"

    overall_pct = (
        0 if overall_result == "Even"
        else round((overall_imp / max(1, overall_imp + overall_deg)) * 100)
    )

    # ---------- Per-area blocks ----------
    area_blocks = []
    for col in areas:
        imp, deg, imp_names, deg_names = count_changes(df_analysis, col)
        area_blocks.append(
            {
                "name": col,
                "improved": int(imp),
                "degraded": int(deg),
                "improvedApps": imp_names,
                "degradedApps": deg_names,
            }
        )

    # ---------- Tiers ----------
    tiers = {}
    try:
        sheet_name = f"OverallAssessment{domain}"
        xls = pd.ExcelFile(current_file_path) if current_file_path else None
        if xls is not None and sheet_name in xls.sheet_names:
            df_overall = pd.read_excel(current_file_path, sheet_name=sheet_name)

            def last_pct(col: str) -> Optional[str]:
                if col in df_overall.columns:
                    s = df_overall[col].astype(str).str.replace("%", "")
                    vals = pd.to_numeric(s, errors="coerce").dropna()
                    return f"{vals.iloc[-1]:.1f}%" if len(vals) else None
                return None

            tiers = {
                "platinum": last_pct("percentageTotalPlatinum"),
                "goldOrBetter": last_pct("percentageTotalGoldOrBetter"),
                "silverOrBetter": last_pct("percentageTotalSilverOrBetter"),
            }
            tiers = {k: v for k, v in tiers.items() if v is not None}
    except Exception:
        tiers = {}

    # ---------- Per-app index ----------
    appsIndex: Dict[str, Any] = {}
    app_names = df_analysis["name"].dropna().astype(str).str.strip().unique().tolist()

    # Load detail sheets
    detail_frames: Dict[str, Optional[pd.DataFrame]] = {}
    for area_col, sheet in DETAIL_SHEETS.get(domain, {}).items():
        try:
            detail_frames[area_col] = pd.read_excel(comparison_result_path, sheet_name=sheet)
        except Exception:
            detail_frames[area_col] = None

    def normalize_status(val: Any) -> str:
        s = str(val or "").lower()
        if "upgraded" in s:
            return "Upgraded"
        if "downgraded" in s:
            return "Downgraded"
        return "No Change"

    for app in app_names:
        row = df_analysis[df_analysis["name"].astype(str).str.strip() == app]
        per_app_areas = []
        per_app_detail = {}

        for area_col in areas:
            status = normalize_status(row[area_col].iloc[0] if not row.empty else None)
            per_app_areas.append({"name": area_col, "status": status})

            df_detail = detail_frames.get(area_col)
            if df_detail is not None and len(df_detail.columns) > 0:
                app_col = "application" if "application" in df_detail.columns else df_detail.columns[0]
                r = df_detail[df_detail[app_col].astype(str).str.strip() == app]
                if not r.empty:
                    vals = {str(c): ("" if pd.isna(r.iloc[0][c]) else str(r.iloc[0][c])) for c in df_detail.columns}
                    per_app_detail[area_col] = vals

        appsIndex[app] = {"areas": per_app_areas, "detail": per_app_detail}

    # ---------- Construct payload ----------
    payload = {
        "domain": domain,
        "generatedAt": dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z",
        "apps": {"total": int(app_total), "names": app_names},
        "overall": {
            "improved": int(overall_imp),
            "degraded": int(overall_deg),
            "result": overall_result,
            "percentage": int(overall_pct),
        },
        "tiers": tiers,
        "areas": area_blocks,
        "appsIndex": appsIndex,
    }

    # ---------- Build Meta ----------
    def _yyyymmdd(d): return d.strftime("%Y%m%d") if d else None

    def _guess_workbook_date(path):
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            d = wb.properties.created or wb.properties.modified
            wb.close()
            return _yyyymmdd(d)
        except:
            try:
                return dt.datetime.utcfromtimestamp(os.path.getmtime(path)).strftime("%Y%m%d")
            except:
                return None

    m = dict(meta or {})

    # controller inference
    if "controller" not in m or not m["controller"]:
        controller_val = None
        for df in detail_frames.values():
            if df is not None and "controller" in df.columns:
                s = df["controller"].dropna().astype(str).str.strip()
                if len(s):
                    controller_val = s.iloc[0]
                    break
        m["controller"] = controller_val or "Unknown"

    # --- NEW: ensure domain is also stored in meta for Trends / History ---
    m.setdefault("domain", domain)

    # dates
    m.setdefault("previousDate", _guess_workbook_date(previous_file_path) or "Unknown")
    m.setdefault("currentDate", _guess_workbook_date(current_file_path) or "Unknown")
    m.setdefault("compareDate", dt.datetime.utcnow().strftime("%Y%m%d_%H%M%S"))

    # ---------- ★ Add OVERALL + TIERS to meta (REQUIRED for Trends UI) ----------
    m.setdefault("improved", int(overall_imp))
    m.setdefault("degraded", int(overall_deg))
    m.setdefault("percentage", int(overall_pct))
    m.setdefault("tiers", tiers)

    payload["meta"] = m


    # ---------- Output ----------
    out_name = f"analysis_summary_{domain.lower()}_{m['compareDate']}.json"
    os.makedirs(result_folder, exist_ok=True)
    out_path = os.path.join(result_folder, out_name)

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    return out_path, out_name, payload


