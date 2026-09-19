
# FILE: C:\Users\raksh\Helicopter_Analysis_System\.pytest_cache\README.md

```
# pytest cache directory #

This directory contains data from the pytest's cache plugin,
which provides the `--lf` and `--ff` options, as well as the `cache` fixture.

**Do not** commit this to version control.

See [the docs](https://docs.pytest.org/en/stable/how-to/cache.html) for more information.

```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\analysis\__init__.py

```
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\analysis\analyzer.py

```
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
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\analysis\anomaly_detector.py

```
import pandas as pd


def detect_spikes(df, column, threshold=2.0):
    """
    Detect unusually high or low values
    using standard deviation.
    """

    if column not in df.columns:
        raise ValueError(
            f"Column '{column}' not found."
        )

    if not pd.api.types.is_numeric_dtype(df[column]):
        raise ValueError(
            f"Column '{column}' must be numeric."
        )

    series = df[column]

    mean = series.mean()
    std = series.std()

    # No variation in the data
    if pd.isna(std) or std == 0:
        return {
            "column": column,
            "mean": float(mean),
            "standard_deviation": float(std)
            if not pd.isna(std) else 0.0,
            "threshold": threshold,
            "spike_indices": [],
            "spike_values": [],
            "spike_count": 0
        }

    upper_limit = mean + threshold * std
    lower_limit = mean - threshold * std

    spike_mask = (
        (series > upper_limit) |
        (series < lower_limit)
    )

    spike_indices = list(
        df.index[spike_mask]
    )

    spike_values = list(
        series[spike_mask]
    )

    return {
        "column": column,
        "mean": float(mean),
        "standard_deviation": float(std),
        "upper_limit": float(upper_limit),
        "lower_limit": float(lower_limit),
        "spike_indices": spike_indices,
        "spike_values": spike_values,
        "spike_count": len(spike_indices)
    }


def detect_all_spikes(
    df,
    selected_attributes,
    threshold=2.0
):
    """
    Detect spikes for all selected
    numeric attributes.
    """

    results = {}

    for column in selected_attributes:

        if column not in df.columns:
            raise ValueError(
                f"Column '{column}' not found."
            )

        # Skip non-numeric columns
        if not pd.api.types.is_numeric_dtype(
            df[column]
        ):
            continue

        results[column] = detect_spikes(
            df,
            column,
            threshold
        )

    return results
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\backend.py

```
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
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\core\__init__.py

```
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\core\backend_service.py

```
from core.data_manager import process_file, get_attributes
from core.validator import validate_data
from core.pipeline import run_pipeline


class BackendService:
    """
    Main interface between the backend and GUI.
    """

    def __init__(self):
        self.file_path = None
        self.dataframe = None
        self.file_info = None
        self.file_type = None
        self.attributes = []
        self.validation = None
        self.result = None

    # ----------------------------------
    # Load File
    # ----------------------------------

    def load_file(self, file_path):

        try:
            result = process_file(file_path)

            self.file_path = file_path
            self.dataframe = result["dataframe"]
            self.file_info = result["file_info"]
            self.file_type = result["file_type"]

            self.attributes = get_attributes(
                self.dataframe
            )

            return {
                "success": True,
                "file_type": self.file_type,
                "file_info": self.file_info,
                "attributes": self.attributes,
                "converted_file": result["converted_file"],
                "message": "File loaded successfully."
            }

        except FileNotFoundError as e:

            return {
                "success": False,
                "message": str(e)
            }

        except ValueError as e:

            return {
                "success": False,
                "message": str(e)
            }

        except Exception as e:

            return {
                "success": False,
                "message": f"Error loading file: {e}"
            }

    # ----------------------------------
    # Validate
    # ----------------------------------

    def validate(self):

        if self.dataframe is None:

            return {
                "success": False,
                "message": "No file has been loaded."
            }

        try:

            self.validation = validate_data(
                self.dataframe
            )

            return {
                "success": True,
                "validation": self.validation
            }

        except Exception as e:

            return {
                "success": False,
                "message": f"Validation error: {e}"
            }

    # ----------------------------------
    # Get Attributes
    # ----------------------------------

    def get_attributes(self):

        if self.dataframe is None:

            return {
                "success": False,
                "message": "No file has been loaded."
            }

        try:

            attributes = get_attributes(
                self.dataframe
            )

            return {
                "success": True,
                "attributes": attributes
            }

        except Exception as e:

            return {
                "success": False,
                "message": f"Error getting attributes: {e}"
            }

    # ----------------------------------
    # Analyze
    # ----------------------------------

    def analyze(
        self,
        selected_attributes,
        export_path=None
    ):

        if self.file_path is None:

            return {
                "success": False,
                "message": "No file has been loaded."
            }

        if not selected_attributes:

            return {
                "success": False,
                "message": "No attributes selected."
            }

        try:

            self.result = run_pipeline(
                self.file_path,
                selected_attributes,
                export_path
            )

            return self.result

        except ValueError as e:

            return {
                "success": False,
                "message": str(e)
            }

        except Exception as e:

            return {
                "success": False,
                "message": f"Analysis error: {e}"
            }

    # ----------------------------------
    # Get Latest Result
    # ----------------------------------

    def get_result(self):

        if self.result is None:

            return {
                "success": False,
                "message": "No analysis has been performed."
            }

        return self.result
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\core\data_manager.py

```
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
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\core\exporter.py

```
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
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\core\file_handler.py

```
import pandas as pd
import os


def load_excel(file_path):
    """
    Load an Excel file and return it as a Pandas DataFrame.
    Supports .xlsx and .xls files.
    """

    if not os.path.exists(file_path):
        raise FileNotFoundError("File not found.")

    extension = os.path.splitext(file_path)[1].lower()

    if extension not in [".xlsx", ".xls"]:
        raise ValueError("Only .xlsx and .xls files are supported.")

    try:
        df = pd.read_excel(file_path)
        return df

    except Exception as e:
        raise Exception(f"Error reading Excel file: {e}")


def get_file_info(df):
    """
    Return basic information about the loaded data.
    """

    return {
        "rows": len(df),
        "columns": len(df.columns),
        "column_names": list(df.columns)
    }
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\core\pipeline.py

```
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
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\core\txt_converter.py

```
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
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\core\validator.py

```
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
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\frontend\app.py

```
import sys
import os

# ============================================================
# PROJECT ROOT
# ============================================================

PROJECT_ROOT = os.path.dirname(
    os.path.dirname(
        os.path.abspath(__file__)
    )
)

if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)


# ============================================================
# IMPORTS
# ============================================================

import streamlit as st
import pandas as pd

from core.backend_service import BackendService


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Helicopter Analysis System",
    page_icon="🚁",
    layout="wide"
)


# ============================================================
# SESSION STATE
# ============================================================

if "service" not in st.session_state:
    st.session_state.service = BackendService()

if "file_loaded" not in st.session_state:
    st.session_state.file_loaded = False

if "analysis_result" not in st.session_state:
    st.session_state.analysis_result = None

if "selected_attributes" not in st.session_state:
    st.session_state.selected_attributes = []

if "graph_type" not in st.session_state:
    st.session_state.graph_type = "Line"

if "converted_file" not in st.session_state:
    st.session_state.converted_file = None


service = st.session_state.service


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.title("🚁 Helicopter")

    st.caption("Analysis System")

    st.divider()

    st.subheader("System Module")

    st.write("📊 Data Analysis")
    st.write("📈 Visualization")
    st.write("⚠️ Anomaly Detection")
    st.write("📥 Data Export")

    st.divider()

    st.subheader("System Status")

    st.success("Backend Online")

    st.caption("Module 01 • Data Analysis")


# ============================================================
# MAIN TITLE
# ============================================================

st.title("🚁 Helicopter Analysis System")

st.write(
    "Flight Data Analytics & Visualization Platform"
)

st.divider()


# ============================================================
# 01 — DATA INPUT
# ============================================================

st.header("01 — Data Input")

st.info(
    "Maximum file size: 200 MB | "
    "Supported formats: XLSX, XLS, TXT"
)

uploaded_file = st.file_uploader(
    "Upload helicopter flight data",
    type=["xlsx", "xls", "txt"]
)


# ============================================================
# PROCESS FILE
# ============================================================

if uploaded_file is not None:

    input_directory = os.path.join(
        PROJECT_ROOT,
        "input"
    )

    os.makedirs(
        input_directory,
        exist_ok=True
    )

    file_path = os.path.join(
        input_directory,
        uploaded_file.name
    )

    # Save uploaded file
    with open(
        file_path,
        "wb"
    ) as file:

        file.write(
            uploaded_file.getbuffer()
        )


    # Load using backend
    load_result = service.load_file(
        file_path
    )


    if load_result["success"]:

        st.session_state.file_loaded = True

        st.session_state.analysis_result = None

        st.session_state.selected_attributes = []


        st.success(
            "Flight data loaded successfully."
        )


        # ====================================================
        # TXT → EXCEL
        # ====================================================

        converted_file = load_result.get(
            "converted_file"
        )


        if converted_file:

            st.session_state.converted_file = (
                converted_file
            )


            st.success(
                "TXT file successfully converted to Excel."
            )


            st.write(
                "Converted Excel File:"
            )


            st.code(
                converted_file
            )


            try:

                with open(
                    converted_file,
                    "rb"
                ) as file:

                    converted_excel = file.read()


                st.download_button(
                    label="⬇️ Download Converted Excel",
                    data=converted_excel,
                    file_name="converted_data.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True
                )


            except FileNotFoundError:

                st.error(
                    "The converted Excel file could not be found."
                )


    else:

        st.session_state.file_loaded = False

        st.error(
            load_result.get(
                "message",
                "Unable to load file."
            )
        )


# ============================================================
# 02 — FILE OVERVIEW
# ============================================================

if st.session_state.file_loaded:

    st.header("02 — File Overview")


    info = service.file_info

    attributes = service.attributes


    col1, col2, col3 = st.columns(3)


    with col1:

        st.metric(
            "File Type",
            service.file_type
        )


    with col2:

        st.metric(
            "Data Points",
            info["rows"]
        )


    with col3:

        st.metric(
            "Parameters",
            info["columns"]
        )


    st.subheader(
        "Available Parameters"
    )


    st.write(
        attributes
    )


# ============================================================
# 03 — ATTRIBUTE SELECTION
# ============================================================

if st.session_state.file_loaded:

    st.header(
        "03 — Attribute Selection"
    )


    attributes = service.attributes


    # --------------------------------------------------------
    # TIME COLUMN
    # --------------------------------------------------------

    if "Time" in attributes:

        time_column = "Time"

    elif "Date" in attributes:

        time_column = "Date"

    else:

        time_column = attributes[0]


    st.info(
        f"Time axis: {time_column}"
    )


    # --------------------------------------------------------
    # SELECT ALL
    # --------------------------------------------------------

    select_all = st.checkbox(
        "Select All Attributes"
    )


    st.write(
        "Select the parameters you want to visualize:"
    )


    # --------------------------------------------------------
    # ATTRIBUTE CHECKBOXES
    # --------------------------------------------------------

    selectable_attributes = [
        attribute
        for attribute in attributes
        if attribute != time_column
    ]


    selected_attributes = []


    # --------------------------------------------------------
    # ALL ATTRIBUTES
    # --------------------------------------------------------

    if select_all:

        selected_attributes = (
            selectable_attributes.copy()
        )

        st.success(
            f"All {len(selected_attributes)} "
            f"attributes selected."
        )


    # --------------------------------------------------------
    # INDIVIDUAL ATTRIBUTES
    # --------------------------------------------------------

    else:

        # Display attributes in columns
        checkbox_columns = st.columns(3)


        for index, attribute in enumerate(
            selectable_attributes
        ):

            column = checkbox_columns[
                index % 3
            ]


            with column:

                selected = st.checkbox(
                    attribute,
                    key=f"attribute_{attribute}"
                )


                if selected:

                    selected_attributes.append(
                        attribute
                    )


    # --------------------------------------------------------
    # SELECTED ATTRIBUTES
    # --------------------------------------------------------

    if selected_attributes:

        st.write("Selected Attributes:")

        st.write(
            selected_attributes
        )

    else:

        st.warning(
            "Please select at least one attribute."
        )


    # ========================================================
    # GRAPH TYPE
    # ========================================================

    st.subheader(
        "Graph Type"
    )


    graph_type = st.selectbox(
        "Choose the graph type for the selected attributes",
        [
            "Line",
            "Bar",
            "Scatter"
        ],
        key="graph_type_selector"
    )


    st.session_state.graph_type = graph_type


    st.write("")


    # ========================================================
    # GENERATE ANALYSIS
    # ========================================================

    if st.button(
        "📊 Generate Analysis",
        type="primary",
        use_container_width=True
    ):

        if not selected_attributes:

            st.error(
                "Please select at least one attribute."
            )

        else:

            st.session_state.selected_attributes = (
                selected_attributes
            )


            # ------------------------------------------------
            # Backend requires selected attributes
            # ------------------------------------------------

            analysis_attributes = [
                time_column
            ] + selected_attributes


            output_directory = os.path.join(
                PROJECT_ROOT,
                "output"
            )

            os.makedirs(
                output_directory,
                exist_ok=True
            )


            export_path = os.path.join(
                output_directory,
                "final_analysis.xlsx"
            )


            # ------------------------------------------------
            # Run analysis
            # ------------------------------------------------

            with st.spinner(
                "Processing helicopter flight data..."
            ):

                result = service.analyze(
                    analysis_attributes,
                    export_path
                )


            if result["success"]:

                st.session_state.analysis_result = (
                    result
                )

                st.success(
                    "Analysis completed successfully."
                )


            else:

                st.session_state.analysis_result = None

                st.error(
                    result.get(
                        "message",
                        "Analysis failed."
                    )
                )


# ============================================================
# 04 — GRAPH OUTPUT
# ============================================================

if st.session_state.analysis_result is not None:

    result = st.session_state.analysis_result


    st.header(
        "04 — Graph Output"
    )


    df = result["selected_data"]


    selected_attributes = (
        st.session_state.selected_attributes
    )


    graph_type = (
        st.session_state.graph_type
    )


    # --------------------------------------------------------
    # FIND TIME COLUMN
    # --------------------------------------------------------

    if "Time" in df.columns:

        time_column = "Time"

    elif "Date" in df.columns:

        time_column = "Date"

    else:

        time_column = df.columns[0]


    st.write(
        f"Each selected parameter is plotted "
        f"against {time_column}."
    )


    # --------------------------------------------------------
    # GENERATE ONE GRAPH PER ATTRIBUTE
    # --------------------------------------------------------

    for attribute in selected_attributes:

        if attribute not in df.columns:

            continue


        st.subheader(
            f"{attribute} vs {time_column}"
        )


        # ----------------------------------------------------
        # Prepare graph data
        # ----------------------------------------------------

        graph_df = df[
            [time_column, attribute]
        ].copy()


        # ----------------------------------------------------
        # Convert Y values to numeric
        # ----------------------------------------------------

        graph_df[attribute] = pd.to_numeric(
            graph_df[attribute],
            errors="coerce"
        )


        graph_df = graph_df.dropna(
            subset=[attribute]
        )


        if graph_df.empty:

            st.warning(
                f"No numeric data available for {attribute}."
            )

            continue


        # ----------------------------------------------------
        # LINE
        # ----------------------------------------------------

        if graph_type == "Line":

            st.line_chart(
                graph_df,
                x=time_column,
                y=attribute
            )


        # ----------------------------------------------------
        # BAR
        # ----------------------------------------------------

        elif graph_type == "Bar":

            st.bar_chart(
                graph_df,
                x=time_column,
                y=attribute
            )


        # ----------------------------------------------------
        # SCATTER
        # ----------------------------------------------------

        elif graph_type == "Scatter":

            st.scatter_chart(
                graph_df,
                x=time_column,
                y=attribute
            )


        st.divider()


# ============================================================
# 05 — ANALYSIS SUMMARY
# ============================================================

if st.session_state.analysis_result is not None:

    result = st.session_state.analysis_result


    st.header(
        "05 — Analysis Summary"
    )


    statistics = (
        result["analysis"]["statistics"]
    )


    for column, stats in statistics.items():

        # Don't show Time as a normal parameter
        if column in ["Time", "Date"]:
            continue


        with st.expander(
            f"📊 {column}"
        ):

            col1, col2, col3, col4 = st.columns(4)


            with col1:

                st.metric(
                    "Mean",
                    round(
                        stats["mean"],
                        2
                    )
                )


            with col2:

                st.metric(
                    "Median",
                    round(
                        stats["median"],
                        2
                    )
                )


            with col3:

                st.metric(
                    "Minimum",
                    round(
                        stats["minimum"],
                        2
                    )
                )


            with col4:

                st.metric(
                    "Maximum",
                    round(
                        stats["maximum"],
                        2
                    )
                )


# ============================================================
# 06 — ANOMALY DETECTION
# ============================================================

if st.session_state.analysis_result is not None:

    result = st.session_state.analysis_result


    st.header(
        "06 — Anomaly Detection"
    )


    spikes = result["spikes"]


    anomaly_found = False


    for column, spike in spikes.items():

        if column in ["Time", "Date"]:

            continue


        if spike["spike_count"] > 0:

            anomaly_found = True


            st.warning(
                f"{column}: "
                f"{spike['spike_count']} "
                f"spike(s) detected."
            )


            st.write(
                "Spike values:",
                spike["spike_values"]
            )


    if not anomaly_found:

        st.success(
            "No anomalies detected in "
            "the selected parameters."
        )


# ============================================================
# 07 — CORRELATION ANALYSIS
# ============================================================

if st.session_state.analysis_result is not None:

    result = st.session_state.analysis_result


    st.header(
        "07 — Correlation Analysis"
    )


    correlation = (
        result["analysis"]["correlation"]
    )


    st.dataframe(
        correlation,
        use_container_width=True
    )


# ============================================================
# 08 — EXPORT
# ============================================================

if st.session_state.analysis_result is not None:

    result = st.session_state.analysis_result


    exported_file = result.get(
        "exported_file"
    )


    if exported_file:

        st.header(
            "08 — Export Analysis"
        )


        try:

            with open(
                exported_file,
                "rb"
            ) as file:

                excel_data = file.read()


            st.download_button(
                label="⬇️ Download Analysis Report",
                data=excel_data,
                file_name="helicopter_analysis.xlsx",
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True
            )


        except FileNotFoundError:

            st.error(
                "Exported analysis file could not be found."
            )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "Helicopter Analysis System • "
    "Flight Data Analytics Platform • "
    "Module 01"
)
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\main.py

```
from core.backend_service import BackendService


service = BackendService()


# ----------------------------------
# 1. Load File
# ----------------------------------

load_result = service.load_file(
    "input/test.xlsx"
)

print("=== LOAD ===")
print(load_result)


# ----------------------------------
# 2. Validate
# ----------------------------------

validation_result = service.validate()

print("\n=== VALIDATION ===")
print(validation_result)


# ----------------------------------
# 3. Get Attributes
# ----------------------------------

attribute_result = service.get_attributes()

print("\n=== ATTRIBUTES ===")
print(attribute_result)


# ----------------------------------
# 4. Analyze
# ----------------------------------

selected_attributes = [
    "Time",
    "RPM",
    "Altitude",
    "Temperature"
]

analysis_result = service.analyze(
    selected_attributes,
    "output/service_analysis.xlsx"
)

print("\n=== ANALYSIS ===")
print("Success:", analysis_result["success"])

print(
    "Exported File:",
    analysis_result.get("exported_file")
)


# ----------------------------------
# 5. Get Latest Result
# ----------------------------------

latest_result = service.get_result()

print("\n=== STORED RESULT ===")

print(
    "Success:",
    latest_result["success"]
)

if latest_result["success"]:

    print(
        "File Type:",
        latest_result["file_type"]
    )

    print(
        "Selected Attributes:",
        latest_result["selected_attributes"]
    )

else:

    print(
        "Message:",
        latest_result["message"]
    )
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\tests\__init__.py

```
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\tests\test_analyzer.py

```
import pandas as pd

from analysis.analyzer import analyze_data


def test_analysis():

    df = pd.DataFrame({
        "RPM": [1000, 1100, 1200],
        "Altitude": [100, 150, 200]
    })

    result = analyze_data(df)

    assert "RPM" in result["statistics"]
    assert "Altitude" in result["statistics"]

    assert result["statistics"]["RPM"]["mean"] == 1100
    assert result["statistics"]["Altitude"]["mean"] == 150
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\tests\test_anomaly_detector.py

```
import pandas as pd

from analysis.anomaly_detector import detect_spikes


def test_spike_detection():

    df = pd.DataFrame({
        "RPM": [
            1000,
            1050,
            1100,
            1080,
            1150,
            3000
        ]
    })

    result = detect_spikes(
        df,
        "RPM"
    )

    assert result["spike_count"] == 1
    assert result["spike_indices"] == [5]
    assert result["spike_values"] == [3000]
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\tests\test_backend_service.py

```
from core.backend_service import BackendService


def test_load_file():

    service = BackendService()

    result = service.load_file(
        "input/test.xlsx"
    )

    assert result["success"] is True
    assert result["file_type"] == "Excel"


def test_attributes():

    service = BackendService()

    service.load_file(
        "input/test.xlsx"
    )

    result = service.get_attributes()

    assert result["success"] is True

    assert result["attributes"] == [
        "Time",
        "RPM",
        "Altitude",
        "Temperature"
    ]


def test_invalid_attribute():

    service = BackendService()

    service.load_file(
        "input/test.xlsx"
    )

    result = service.analyze(
        ["InvalidColumn"]
    )

    assert result["success"] is False
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\tests\test_file_handler.py

```
from core.file_handler import load_excel, get_file_info


def test_load_excel():

    df = load_excel("input/test.xlsx")

    assert not df.empty
    assert len(df) == 5
    assert len(df.columns) == 4


def test_get_file_info():

    df = load_excel("input/test.xlsx")

    info = get_file_info(df)

    assert info["rows"] == 5
    assert info["columns"] == 4

    assert info["column_names"] == [
        "Time",
        "RPM",
        "Altitude",
        "Temperature"
    ]
```

# FILE: C:\Users\raksh\Helicopter_Analysis_System\tests\test_validator.py

```
import pandas as pd

from core.validator import validate_data


def test_valid_data():

    df = pd.DataFrame({
        "Time": [1, 2, 3],
        "RPM": [1000, 1050, 1100]
    })

    result = validate_data(df)

    assert result["is_empty"] is False
    assert result["is_valid"] is True
    assert result["duplicate_rows"] == 0


def test_missing_values():

    df = pd.DataFrame({
        "Time": [1, 2, 3],
        "RPM": [1000, None, 1100]
    })

    result = validate_data(df)

    assert "RPM" in result["missing_values"]
    assert result["missing_values"]["RPM"] == 1
```
