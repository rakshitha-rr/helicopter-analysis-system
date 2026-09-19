import pandas as pd


def validate_data(df):
    """
    Validate a Pandas DataFrame and return
    a structured validation report.
    """

    report = {
        "is_empty": False,
        "empty_columns": [],
        "missing_values": {},
        "duplicate_rows": 0,
        "numeric_columns": [],
        "non_numeric_columns": [],
        "datetime_columns": [],
        "total_rows": len(df),
        "total_columns": len(df.columns),
        "is_valid": True
    }

    # Check empty DataFrame
    if df.empty:
        report["is_empty"] = True
        report["is_valid"] = False
        return report

    # Check empty column names
    for column in df.columns:

        if str(column).strip() == "":
            report["empty_columns"].append(column)

    if report["empty_columns"]:
        report["is_valid"] = False

    # Missing values
    missing = df.isnull().sum()

    for column, count in missing.items():

        if count > 0:
            report["missing_values"][column] = int(count)

    # Duplicate rows
    report["duplicate_rows"] = int(
        df.duplicated().sum()
    )

    # Numeric and non-numeric columns
    for column in df.columns:

        if pd.api.types.is_numeric_dtype(df[column]):
            report["numeric_columns"].append(column)

        else:
            report["non_numeric_columns"].append(column)

    # Date/time columns
    for column in df.columns:

        if pd.api.types.is_datetime64_any_dtype(
            df[column]
        ):
            report["datetime_columns"].append(column)

    return report