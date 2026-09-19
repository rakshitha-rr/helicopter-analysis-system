import os
import pandas as pd


def export_to_excel(df, output_path):
    """
    Export a DataFrame to an Excel file.
    """

    if df is None or df.empty:
        raise ValueError(
            "Cannot export an empty DataFrame."
        )

    output_directory = os.path.dirname(
        output_path
    )

    if output_directory:
        os.makedirs(
            output_directory,
            exist_ok=True
        )

    try:

        df.to_excel(
            output_path,
            index=False
        )

        return output_path

    except Exception as e:

        raise Exception(
            f"Error exporting Excel file: {e}"
        )