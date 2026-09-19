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