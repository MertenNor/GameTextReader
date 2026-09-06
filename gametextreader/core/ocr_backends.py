"""
OCR backend helpers.
"""

from __future__ import annotations

from typing import Any, Dict, List

import numpy as np


RAPID_LANGUAGE_MAP: Dict[str, str] = {
    "auto": "Auto",
}


class RapidOCRBackend:
    """Thin wrapper around RapidOCR with a cached engine instance."""

    def __init__(self) -> None:
        self._engine: Any = None

    @staticmethod
    def is_available() -> bool:
        try:
            from rapidocr import RapidOCR  # noqa: F401

            return True
        except Exception:
            return False

    @staticmethod
    def get_available_languages() -> Dict[str, str]:
        return dict(RAPID_LANGUAGE_MAP)

    def clear_cache(self) -> None:
        self._engine = None

    def _get_or_create_engine(self) -> Any:
        if self._engine is not None:
            return self._engine

        from rapidocr import RapidOCR

        self._engine = RapidOCR()
        return self._engine

    def image_to_string(self, image: Any, lang: str = "auto") -> str:
        engine = self._get_or_create_engine()

        # RapidOCR accepts ndarray input.
        image_array = np.array(image)
        result = engine(image_array)

        lines = self._extract_text_lines(result)
        return "\n".join(lines).strip()

    def _extract_text_lines(self, data: Any) -> List[str]:
        lines: List[str] = []

        if hasattr(data, "txts"):
            try:
                for item in getattr(data, "txts", []):
                    text_candidate = str(item).strip()
                    if text_candidate:
                        lines.append(text_candidate)
                return lines
            except Exception:
                pass

        # Some RapidOCR versions return (result, elapsed).
        if isinstance(data, tuple) and data:
            lines.extend(self._extract_text_lines(data[0]))
            return lines

        if isinstance(data, list):
            for item in data:
                lines.extend(self._extract_text_lines(item))
            return lines

        if isinstance(data, tuple) and len(data) >= 2:
            candidate = data[1]
            if isinstance(candidate, (list, tuple)) and candidate:
                text_candidate = candidate[0]
                if isinstance(text_candidate, str) and text_candidate.strip():
                    lines.append(text_candidate.strip())
            return lines

        return lines
