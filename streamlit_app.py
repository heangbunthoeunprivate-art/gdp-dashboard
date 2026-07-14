import runpy
from pathlib import Path


# Streamlit Community Cloud defaults to streamlit_app.py.
# Run the real application file so deploy works even with default settings.
runpy.run_path(str(Path(__file__).with_name("Sevenz.py")), run_name="__main__")
