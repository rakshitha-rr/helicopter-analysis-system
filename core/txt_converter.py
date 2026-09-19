import pandas as pd
import os


def detect_delimiter(file_path):
    """
    Detect the delimiter used in a TXT file.
    """

    with open(file_path, "r", encoding="utf-8") as file:
        first_line = file.readline()

    delimiters = [",", "\t", ";", "|"]

    for delimiter in delimiters:
        if delimiter in first_line:
            return delimiter

    # Default to whitespace
    return r"\s+"


def save_as_excel(df, output_path):
    """
    Save a DataFrame as an Excel file.
    """

    try:
        output_directory = os.path.dirname(output_path)

        if output_directory:
            os.makedirs(output_directory, exist_ok=True)

        df.to_excel(output_path, index=False)

    except Exception as e:
        raise Exception(f"Error saving Excel file: {e}")


def txt_to_excel(file_path, output_path):
    """
    Read TXT, convert it to DataFrame,
    save it as Excel and return the DataFrame.
    """

    if not os.path.exists(file_path):
        raise FileNotFoundError("TXT file not found.")

    try:
        delimiter = detect_delimiter(file_path)

        df = pd.read_csv(
            file_path,
            sep=delimiter,
            engine="python"
        )

        save_as_excel(df, output_path)

        return df

    except Exception as e:
        raise Exception(f"Error converting TXT to Excel: {e}")