from core.data_manager import load_input_file
from core.validator import validate_data
from analysis.analyzer import analyze_data


def process_file(file_path, selected_columns=None):
    """
    Complete backend processing pipeline.

    Steps:
        1. Load Excel or TXT
        2. Validate data
        3. Select attributes
        4. Perform analysis
        5. Return clean results
    """

    # ========================================================
    # 1. LOAD FILE
    # ========================================================

    file_result = load_input_file(file_path)

    df = file_result["dataframe"]

    # ========================================================
    # 2. VALIDATE DATA
    # ========================================================

    validation_result = validate_data(df)

    # ========================================================
    # 3. CHECK EMPTY FILE
    # ========================================================

    if validation_result["is_empty"]:
        raise ValueError(
            "The uploaded file contains no data."
        )

    # ========================================================
    # 4. DETERMINE SELECTED COLUMNS
    # ========================================================

    if selected_columns is None:

        selected_columns = (
            df.select_dtypes(include="number")
            .columns
            .tolist()
        )

    # ========================================================
    # 5. ANALYZE DATA
    # ========================================================

    analysis_result = analyze_data(
        df,
        selected_columns
    )

    # ========================================================
    # 6. CREATE CLEAN FILE INFORMATION
    # ========================================================

    file_information = {
        "file_type": file_result["file_type"],
        "rows": int(len(df)),
        "columns": list(df.columns),
        "converted_file": file_result["converted_file"]
    }

    # ========================================================
    # 7. CREATE CLEAN RESULT
    # ========================================================

    return {
        "file": file_information,

        "validation": validation_result,

        "analysis": {
            "selected_columns":
                analysis_result["selected_columns"],

            "numeric_columns":
                analysis_result["numeric_columns"],

            "statistics":
                analysis_result["statistics"],

            "correlation":
                analysis_result["correlation"],

            "anomalies":
                analysis_result["anomalies"],

            "sudden_changes":
                analysis_result["sudden_changes"]
        }
    }