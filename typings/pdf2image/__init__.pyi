from typing import Any, List

from PIL.Image import Image

def convert_from_path(pdf_path: str, dpi: int = 200, **kwargs: Any) -> List[Image]: ...
