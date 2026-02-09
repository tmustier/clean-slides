from __future__ import annotations

from typing import Any

from PIL.Image import Image

def convert_from_path(pdf_path: str, dpi: int = 200, **kwargs: Any) -> list[Image]: ...
