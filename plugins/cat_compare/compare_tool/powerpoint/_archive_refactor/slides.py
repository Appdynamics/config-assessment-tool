# compare_tool/powerpoint/slides.py

from __future__ import annotations

import logging
from typing import Optional, Dict, Any

import pandas as pd

from pptx.util import Pt
from pptx.dml.color import RGBColor
from typing import Set

from ..base import (
    ChannelConfig,
    load_template_presentation,
    find_table_placeholder_by_name,
    insert_table_at_placeholder,
    slide_title_text,
    choose_slide_for_section,
    set_arrow_cell,
    set_shape_text,
    first_present_col,
    parse_grade_transition,
    norm_grade,
    PINK,
)

log = logging.getLogger(__name__)
log.info("slides.py imported")


def style_header_cell(cell) -> None:
    """
    Apply a simple header style: bold, centred, grey background.
    """
    cell.text_frame.paragraphs[0].font.bold = True
    cell.text_frame.paragraphs[0].font.size = Pt(11)

    # light grey fill
    fill = cell.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(230, 230, 230)


def autofit_table_columns(table) -> None:
    """
    Basic column width normalisation. python-pptx doesn't auto-fit,
    but we can at least make widths equal.
    """
    col_count = len(table.columns)
    if col_count == 0:
        return

    total_width = sum(col.width for col in table.columns)
    equal_width = total_width // col_count

    for col in table.columns:
        col.width = equal_width

def generate_powerpoint_for_channel(
    cfg: ChannelConfig,
    comparison_result_path: str,
    powerpoint_output_path: str,
    current_file_path: str,
    previous_file_path: str,
    config: Optional[Dict[str, Any]] = None,
) -> str:
    """
    Shared generator used by both BRUM and MRUM.

    BRUM/MRUM-specific differences are controlled by `cfg`.
    The internal steps map 1:1 to your existing functions:
      - key callouts
      - applications improved
      - summary slide
      - overall assessment
      - entity comparison
      - deep-dive slides (network requests + HRA)
      - textbox for #apps
    """
    log_prefix = f"[{cfg.name}]"

    try:
        prs = load_template_presentation(cfg, config=config)

        # -------- Load data -------- #
        df_current_analysis = pd.read_excel(current_file_path, sheet_name="Analysis")
        number_of_apps = (
            df_current_analysis["name"]
            .dropna()
            .astype(str)
            .str.strip()
            .ne("")
            .sum()
        )

        current_summary_df = pd.read_excel(current_file_path, sheet_name="Summary")
        previous_summary_df = pd.read_excel(previous_file_path, sheet_name="Summary")

        summary_df = pd.read_excel(comparison_result_path, sheet_name="Summary")
        df_analysis = pd.read_excel(comparison_result_path, sheet_name="Analysis")
        df_network_requests = pd.read_excel(comparison_result_path, sheet_name=cfg.sheet_network)
        df_health_rules = pd.read_excel(comparison_result_path, sheet_name=cfg.sheet_hra)

        try:
            curr_overall_df = pd.read_excel(comparison_result_path, sheet_name=cfg.sheet_overall)
        except Exception:
            curr_overall_df = pd.DataFrame()

        try:
            prev_overall_df = pd.read_excel(previous_file_path, sheet_name=cfg.sheet_overall)
        except Exception:
            prev_overall_df = pd.DataFrame()

        log.debug("%s Loaded comparison workbooks successfully.", log_prefix)

        used_slide_ids: set[int] = set()

        # -------- 1) Key Callouts / Maturity badge -------- #
        _build_key_callouts_and_maturity(
            prs,
            cfg,
            used_slide_ids,
            df_analysis,
            curr_overall_df,
            prev_overall_df,
            number_of_apps,
        )

        # -------- 2) Applications improved -------- #
        _build_improved_apps(prs, cfg, used_slide_ids, df_analysis)

        # -------- 3) Summary slide (prev/current/comparison) -------- #
        _build_summary_slide(
            prs,
            used_slide_ids,
            previous_summary_df,
            current_summary_df,
            summary_df,
        )

        # -------- 4) Overall assessment (single row) -------- #
        _build_overall_assessment(prs, cfg, used_slide_ids, df_analysis)

        # -------- 5) Entity comparison (NR + HRA) -------- #
        _build_entity_comparison(prs, cfg, used_slide_ids, df_analysis)

        # -------- 6) Slide 11 – Network Requests deep dive -------- #
        _build_network_requests_deep_dive(
            prs,
            cfg,
            used_slide_ids,
            df_analysis,
            df_network_requests,
        )

        # -------- 7) Slide 12 – Health Rules & Alerting deep dive -------- #
        _build_hra_deep_dive(
            prs,
            cfg,
            used_slide_ids,
            df_analysis,
            df_health_rules,
        )

        # -------- 8) TextBox 7 – number of applications -------- #
        _set_app_count_textbox(prs, number_of_apps)

        # -------- Save -------- #
        prs.save(powerpoint_output_path)
        log.debug("%s PowerPoint saved to: %s", log_prefix, powerpoint_output_path)
        return powerpoint_output_path

    except Exception:
        log.exception("%s Error generating PowerPoint", log_prefix)
        raise

# --------------------------------------------------------------------
# Stub builder helpers – currently no-ops so Pylance is happy.
# We will gradually move real logic from brum.py / mrum.py into these.
# --------------------------------------------------------------------

def _build_key_callouts_and_maturity(
    prs,
    cfg,
    used_slide_ids: Set[int],
    df_analysis,
    curr_overall_df,
    prev_overall_df,
    number_of_apps: int,
):
    """Placeholder: key callouts + maturity badge.
    MRUM/BRUM still use their legacy implementations.
    """
    return


def _build_improved_apps(
    prs,
    cfg,
    used_slide_ids: Set[int],
    df_analysis,
):
    """Placeholder: 'Applications Improved' slide."""
    return


def _build_summary_slide(
    prs,
    used_slide_ids: Set[int],
    previous_summary_df,
    current_summary_df,
    summary_df,
):
    """Placeholder: combined summary slide."""
    return


def _build_overall_assessment(
    prs,
    cfg,
    used_slide_ids: Set[int],
    df_analysis,
):
    """Placeholder: single-row overall assessment slide."""
    return


def _build_entity_comparison(
    prs,
    cfg,
    used_slide_ids: Set[int],
    df_analysis,
):
    """Placeholder: entity comparison slide."""
    return


def _build_network_requests_deep_dive(
    prs,
    cfg,
    used_slide_ids: Set[int],
    df_analysis,
    df_network_requests,
):
    """Placeholder: Network Requests deep dive (Slide 11)."""
    return


def _build_hra_deep_dive(
    prs,
    cfg,
    used_slide_ids: Set[int],
    df_analysis,
    df_health_rules,
):
    """Placeholder: Health Rules & Alerting deep dive (Slide 12)."""
    return


def _set_app_count_textbox(prs, number_of_apps: int):
    """Placeholder: populate 'TextBox 7' with number of apps."""
    return
