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