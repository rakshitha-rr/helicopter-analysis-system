import os
import sys

from streamlit.web import bootstrap


def get_app_path():
    if getattr(sys, "frozen", False):
        # In PyInstaller --onedir, bundled data is under _internal.
        base_dir = os.path.dirname(sys.executable)
        return os.path.join(
            base_dir,
            "_internal",
            "frontend",
            "app.py",
        )

    base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "frontend", "app.py")


if __name__ == "__main__":
    app_path = get_app_path()

    print(f"Starting application: {app_path}")

    if not os.path.isfile(app_path):
        print("ERROR: Streamlit application was not found.")
        print(app_path)
        input("Press Enter to exit...")
        sys.exit(1)

    flag_options = {
        "global.developmentMode": False,
        "server.headless": True,
        "server.address": "127.0.0.1",
        "server.port": 8501,
        "browser.gatherUsageStats": False,
        "server.enableCORS": False,
        "server.enableXsrfProtection": False,
    }

    bootstrap.load_config_options(flag_options=flag_options)

    bootstrap.run(
        app_path,
        "streamlit run",
        [],
        flag_options,
    )