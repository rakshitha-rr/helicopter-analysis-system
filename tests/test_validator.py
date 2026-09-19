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