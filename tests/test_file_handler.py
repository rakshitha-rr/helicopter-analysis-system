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