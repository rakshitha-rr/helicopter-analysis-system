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