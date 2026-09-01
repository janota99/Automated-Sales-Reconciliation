"""Presentation assets used by the Streamlit interface."""

from __future__ import annotations

import base64
from pathlib import Path


ASSET_DIRECTORY = Path(__file__).resolve().parent


def png_data_uri(filename: str) -> str:
    """Load a PNG relative to the application and return an HTML data URI."""
    image_path = ASSET_DIRECTORY / filename

    if not image_path.is_file():
        raise FileNotFoundError(f"Logo asset was not found: {image_path}")

    encoded = base64.b64encode(image_path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


QUICKBOOKS_LOGO_URI = png_data_uri("quickbooks.png")
INFOR_LOGO_URI = png_data_uri("infor.png")