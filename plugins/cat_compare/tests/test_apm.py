import pytest
import shutil
import pandas as pd
from pathlib import Path
from compare_tool.powerpoint.apm import generate_powerpoint_from_apm

def test_generate_powerpoint_from_apm(tmp_path):
    # Setup
    # Path to the existing template
    valid_template_path = Path(__file__).parent.parent / "templates" / "template.pptx"

    # Copy the valid template to the temporary directory
    template_path = tmp_path / "template.pptx"
    shutil.copy(valid_template_path, template_path)

    # Path to the real uploaded files
    uploads_dir = Path(__file__).parent.parent / "uploads"

    # Copy the real uploaded files to the temporary directory
    current_apm_path = uploads_dir / "current_apm.xlsx"
    current_sum_path = uploads_dir / "current_sum.xlsx"
    previous_apm_path = uploads_dir / "previous_apm.xlsx"
    previous_sum_path = uploads_dir / "previous_sum.xlsx"

    tmp_current_apm_path = tmp_path / "current_apm.xlsx"
    tmp_current_sum_path = tmp_path / "current_sum.xlsx"
    tmp_previous_apm_path = tmp_path / "previous_apm.xlsx"
    tmp_previous_sum_path = tmp_path / "previous_sum.xlsx"

    shutil.copy(current_apm_path, tmp_current_apm_path)
    shutil.copy(current_sum_path, tmp_current_sum_path)
    shutil.copy(previous_apm_path, tmp_previous_apm_path)
    shutil.copy(previous_sum_path, tmp_previous_sum_path)

    # Create a valid Excel file for comparison_result_path
    comparison_result_path = tmp_path / "comparison.xlsx"
    with pd.ExcelWriter(comparison_result_path) as writer:
        df_summary = pd.DataFrame({"Metric": ["Metric1", "Metric2"], "Value": [10, 20]})
        df_summary.to_excel(writer, sheet_name="Summary", index=False)

    # Results directory
    results_dir = tmp_path / "results"
    results_dir.mkdir()

    # Act
    generate_powerpoint_from_apm(
        comparison_result_path=str(comparison_result_path),
        powerpoint_output_path=str(results_dir / "output.pptx"),
        current_file_path=str(tmp_current_apm_path),
        previous_file_path=str(tmp_previous_apm_path),
        template_path=str(template_path),
    )

    # Assert
    assert (results_dir / "output.pptx").exists(), "PowerPoint output file was not created."