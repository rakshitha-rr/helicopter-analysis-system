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