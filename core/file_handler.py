import os
import pandas as pd


# ============================================================
# LOAD EXCEL
# ============================================================

def load_excel(file_path):
    return pd.read_excel(file_path)


# ============================================================
# LOAD CSV
# ============================================================

def load_csv(file_path):
    return pd.read_csv(file_path)


# ============================================================
# LOAD TXT
# ============================================================

def load_txt(file_path):

    try:
        # Try whitespace-separated TXT
        data = pd.read_csv(
            file_path,
            sep=r"\s+",
            engine="python"
        )

        if len(data.columns) > 1:
            return data

    except Exception:
        pass

    # Fallback for comma-separated TXT
    return pd.read_csv(file_path)


# ============================================================
# GET FILE INFORMATION
# ============================================================

def get_file_info(data):

    return {
        "rows": len(data),
        "columns": len(data.columns),
        "column_names": list(data.columns)
    }


# ============================================================
# LOAD ANY SUPPORTED FILE
# ============================================================

def load_file(file_path):

    if not os.path.exists(file_path):

        raise FileNotFoundError(
            f"File not found: {file_path}"
        )


    extension = os.path.splitext(
        file_path
    )[1].lower()


    # --------------------------------------------------------
    # EXCEL
    # --------------------------------------------------------

    if extension in [".xlsx", ".xls"]:

        data = load_excel(file_path)

        file_type = "Excel"

        converted_file = None


    # --------------------------------------------------------
    # CSV
    # --------------------------------------------------------

    elif extension == ".csv":

        data = load_csv(file_path)

        file_type = "CSV"

        converted_file = None


    # --------------------------------------------------------
    # TXT
    # --------------------------------------------------------

    elif extension == ".txt":

        data = load_txt(file_path)

        file_type = "TXT"


        # ----------------------------------------------------
        # TXT → EXCEL
        # ----------------------------------------------------

        project_root = os.path.dirname(
            os.path.dirname(
                os.path.abspath(__file__)
            )
        )


        output_directory = os.path.join(
            project_root,
            "output"
        )


        os.makedirs(
            output_directory,
            exist_ok=True
        )


        base_name = os.path.splitext(
            os.path.basename(file_path)
        )[0]


        converted_file = os.path.join(
            output_directory,
            base_name + ".xlsx"
        )


        data.to_excel(
            converted_file,
            index=False
        )


    else:

        raise ValueError(
            f"Unsupported file type: {extension}"
        )


    # --------------------------------------------------------
    # FILE INFORMATION
    # --------------------------------------------------------

    return {
        "data": data,
        "file_type": file_type,
        "file_info": get_file_info(data),
        "converted_file": converted_file
    }


# ============================================================
# FILE HANDLER CLASS
# ============================================================

class FileHandler:

    def __init__(self):

        self.data = None

        self.file_type = None

        self.file_info = None

        self.converted_file = None


    # ========================================================
    # LOAD FILE
    # ========================================================

    def load_file(self, file_path):

        result = load_file(
            file_path
        )


        self.data = result["data"]

        self.file_type = result["file_type"]

        self.file_info = result["file_info"]

        self.converted_file = result["converted_file"]


        return self.data


    # ========================================================
    # GET ATTRIBUTES
    # ========================================================

    def get_attributes(self):

        if self.data is None:

            return []


        return list(
            self.data.columns
        )


    # ========================================================
    # GET DATA
    # ========================================================

    def get_data(self):

        return self.data