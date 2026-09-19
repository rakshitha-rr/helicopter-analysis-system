import os

from core.file_handler import load_excel, get_file_info
from core.txt_converter import txt_to_excel


def process_file(file_path):
    """
    Process an Excel or TXT file and return:
    - DataFrame
    - File information
    - File type
    - Converted Excel path
    """

    if not os.path.exists(file_path):
        raise FileNotFoundError("File not found.")

    extension = os.path.splitext(file_path)[1].lower()

    # Excel
    if extension in [".xlsx", ".xls"]:

        df = load_excel(file_path)

        return {
            "dataframe": df,
            "file_type": "Excel",
            "file_info": get_file_info(df),
            "converted_file": None
        }

    # TXT
    elif extension == ".txt":

        output_path = os.path.join(
            "output",
            os.path.splitext(
                os.path.basename(file_path)
            )[0] + ".xlsx"
        )

        df = txt_to_excel(
            file_path,
            output_path
        )

        return {
            "dataframe": df,
            "file_type": "TXT",
            "file_info": get_file_info(df),
            "converted_file": output_path
        }

    else:
        raise ValueError(
            "Unsupported file type. "
            "Please use .xlsx, .xls, or .txt."
        )


def get_attributes(df):
    """
    Return all available column names.
    """

    return list(df.columns)


def select_attributes(df, selected_attributes):
    """
    Return a DataFrame containing only
    the selected attributes.
    """

    if not selected_attributes:
        raise ValueError("No attributes selected.")

    invalid_attributes = [
        column
        for column in selected_attributes
        if column not in df.columns
    ]

    if invalid_attributes:
        raise ValueError(
            f"Invalid attributes selected: "
            f"{invalid_attributes}"
        )

    return df[selected_attributes].copy()