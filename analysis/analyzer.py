import pandas as pd


def analyze_data(df):
    """
    Perform basic statistical analysis
    on numeric columns.
    """

    if df.empty:
        raise ValueError(
            "Cannot analyze an empty DataFrame."
        )

    numeric_df = df.select_dtypes(
        include="number"
    )

    if numeric_df.empty:
        raise ValueError(
            "No numeric attributes available "
            "for analysis."
        )

    statistics = {}

    for column in numeric_df.columns:

        series = numeric_df[column]

        statistics[column] = {

            "count": int(
                series.count()
            ),

            "mean": float(
                series.mean()
            ),

            "median": float(
                series.median()
            ),

            "minimum": float(
                series.min()
            ),

            "maximum": float(
                series.max()
            ),

            "standard_deviation": float(
                series.std()
            ),

            "variance": float(
                series.var()
            ),

            "range": float(
                series.max() -
                series.min()
            ),

            "missing_values": int(
                series.isnull().sum()
            )
        }

    correlation = numeric_df.corr()

    duplicate_rows = int(
        df.duplicated().sum()
    )

    return {

        "statistics": statistics,

        "correlation": correlation,

        "duplicate_rows": duplicate_rows,

        "numeric_columns":
            list(numeric_df.columns)
    }