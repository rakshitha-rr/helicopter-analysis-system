from core.data_manager import process_file, select_attributes
from core.validator import validate_data
from core.exporter import export_to_excel
from analysis.analyzer import analyze_data
from analysis.anomaly_detector import detect_all_spikes


def run_pipeline(file_path, selected_attributes, export_path=None):

    # 1. Process file
    file_result = process_file(file_path)

    df = file_result["dataframe"]

    # 2. Validate data
    validation_result = validate_data(df)

    if not validation_result["is_valid"]:
        return {
            "success": False,
            "file_info": file_result["file_info"],
            "file_type": file_result["file_type"],
            "validation": validation_result,
            "message": "Data validation failed."
        }

    # 3. Select attributes
    selected_df = select_attributes(
        df,
        selected_attributes
    )

    # 4. Statistical analysis
    analysis_result = analyze_data(
        selected_df
    )

    # 5. Detect spikes
    spike_result = detect_all_spikes(
        selected_df,
        selected_attributes
    )

    # 6. Export
    exported_file = None

    if export_path is not None:
        exported_file = export_to_excel(
            selected_df,
            export_path
        )

    # 7. Return complete result
    return {
        "success": True,
        "file_type": file_result["file_type"],
        "file_info": file_result["file_info"],
        "converted_file": file_result["converted_file"],
        "validation": validation_result,
        "selected_attributes": selected_attributes,
        "selected_data": selected_df,
        "analysis": analysis_result,
        "spikes": spike_result,
        "exported_file": exported_file
    }