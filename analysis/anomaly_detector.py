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