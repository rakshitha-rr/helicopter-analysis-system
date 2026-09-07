from core.file_handler import FileHandler
from core.validator import validate_data
from analysis.analyzer import analyze_data
from analysis.anomaly_detector import detect_all_spikes


class BackendService:

    def __init__(self):
        self.file_handler = FileHandler()

        self.data = None
        self.file_type = None
        self.file_info = None
        self.attributes = []

        self.validation_result = None
        self.analysis_result = None
        self.last_result = None

    # ========================================================
    # LOAD FILE
    # ========================================================

    def load_file(self, file_path):

        try:
            self.data = self.file_handler.load_file(
                file_path
            )

            self.file_type = (
                self.file_handler.file_type
            )

            self.file_info = (
                self.file_handler.file_info
            )

            self.attributes = (
                self.file_handler.get_attributes()
            )

            return {
                "success": True,
                "file_type": self.file_type,
                "file_info": self.file_info,
                "attributes": self.attributes,
                "converted_file":
                    self.file_handler.converted_file,
                "message":
                    "File loaded successfully."
            }

        except Exception as error:

            return {
                "success": False,
                "message": str(error)
            }

    # ========================================================
    # VALIDATE
    # ========================================================

    def validate(self):

        if self.data is None:

            return {
                "success": False,
                "message": "No file loaded."
            }

        try:

            result = validate_data(
                self.data
            )

            self.validation_result = result

            return {
                "success": True,
                "validation": result
            }

        except Exception as error:

            return {
                "success": False,
                "message": str(error)
            }

    # ========================================================
    # GET ATTRIBUTES
    # ========================================================

    def get_attributes(self):

        if self.data is None:

            return {
                "success": False,
                "attributes": [],
                "message": "No file loaded."
            }

        return {
            "success": True,
            "attributes":
                list(self.data.columns)
        }

    # ========================================================
    # ANALYZE
    # ========================================================

    def analyze(
        self,
        selected_attributes,
        export_path=None
    ):

        # ----------------------------------------------------
        # Check file
        # ----------------------------------------------------

        if self.data is None:

            return {
                "success": False,
                "message": "No file loaded."
            }

        # ----------------------------------------------------
        # Check attributes
        # ----------------------------------------------------

        if not selected_attributes:

            return {
                "success": False,
                "message": "No attributes selected."
            }

        # ----------------------------------------------------
        # Check invalid attributes
        # ----------------------------------------------------

        invalid_attributes = [

            attribute

            for attribute in selected_attributes

            if attribute not in self.data.columns

        ]

        if invalid_attributes:

            return {
                "success": False,
                "message":
                    f"Invalid attributes selected: "
                    f"{invalid_attributes}"
            }

        # ----------------------------------------------------
        # Select data
        # ----------------------------------------------------

        selected_data = self.data[
            selected_attributes
        ].copy()

        # ----------------------------------------------------
        # ANALYSIS
        # ----------------------------------------------------

        try:

            analysis = analyze_data(
                selected_data
            )

        except TypeError:

            analysis = analyze_data(
                self.data,
                selected_attributes
            )

        # ----------------------------------------------------
        # SPIKE DETECTION
        # ----------------------------------------------------

        spikes = detect_all_spikes(
            self.data,
            selected_attributes
        )

        # ----------------------------------------------------
        # EXPORT
        # ----------------------------------------------------

        exported_file = None

        if export_path:

            try:

                selected_data.to_excel(
                    export_path,
                    index=False
                )

                exported_file = export_path

            except Exception as error:

                return {
                    "success": False,
                    "message":
                        f"Export failed: {error}"
                }

        # ----------------------------------------------------
        # STATISTICS
        # ----------------------------------------------------

        statistics = analysis.get(
            "statistics",
            {}
        )

        # ----------------------------------------------------
        # CORRELATION
        # ----------------------------------------------------

        correlation = analysis.get(
            "correlation",
            selected_data.corr(
                numeric_only=True
            )
        )

        # ----------------------------------------------------
        # STORE ANALYSIS RESULT
        # ----------------------------------------------------

        self.analysis_result = {

            "statistics":
                statistics,

            "correlation":
                correlation
        }

        # ----------------------------------------------------
        # STORE COMPLETE RESULT
        # ----------------------------------------------------

        self.last_result = {

            "success": True,

            "file_type":
                self.file_type,

            "file_info":
                self.file_info,

            "selected_attributes":
                selected_attributes,

            "selected_data":
                selected_data,

            "analysis":
                self.analysis_result,

            "spikes":
                spikes,

            "exported_file":
                exported_file
        }

        return self.last_result

    # ========================================================
    # GET LAST RESULT
    # ========================================================

    def get_last_result(self):

        return self.last_result