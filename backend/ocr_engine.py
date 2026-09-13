"""
Offline OCR & Vision Layer
Powered by local RapidOCR ONNX models.
Runs 100% locally on CPU without any cloud or API key dependencies.
"""

import os
import io
from typing import List, Dict, Any, Optional, Tuple
from PIL import Image
import numpy as np

class OfflineOCREngine:
    _instance = None
    _engine = None

    @classmethod
    def get_engine(cls):
        if cls._engine is None:
            try:
                from rapidocr import RapidOCR
                cls._engine = RapidOCR()
            except Exception as e:
                print(f"[OfflineOCREngine Warning]: Could not initialize RapidOCR: {e}")
                cls._engine = False
        return cls._engine

    @classmethod
    def extract_text_from_image_bytes(cls, img_bytes: bytes) -> Tuple[str, float]:
        """
        Runs OCR on raw image bytes.
        Returns:
            Tuple[str, float]: (extracted_text, average_confidence)
        """
        engine = cls.get_engine()
        if not engine:
            return "", 0.0

        try:
            pil_img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
            max_dim = 1280
            if max(pil_img.size) > max_dim:
                pil_img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
            np_img = np.array(pil_img)
            
            ocr_res = engine(np_img)
            if not ocr_res:
                return "", 0.0

            # Modern RapidOCROutput format
            if hasattr(ocr_res, "txts") and ocr_res.txts:
                lines = [str(t).strip() for t in ocr_res.txts if str(t).strip()]
                scores = [float(s) for s in (ocr_res.scores or []) if s is not None]
                full_text = "\n".join(lines)
                avg_conf = sum(scores) / len(scores) if scores else 0.0
                return full_text, round(avg_conf, 3)

            # Legacy tuple/list format
            items = ocr_res[0] if (isinstance(ocr_res, (list, tuple)) and len(ocr_res) == 2 and isinstance(ocr_res[0], list)) else ocr_res
            if isinstance(items, (list, tuple)):
                lines = []
                confidences = []
                for item in items:
                    if isinstance(item, (list, tuple)) and len(item) >= 2:
                        text = str(item[1]).strip()
                        score = float(item[2]) if len(item) >= 3 else 1.0
                        if text:
                            lines.append(text)
                            confidences.append(score)
                full_text = "\n".join(lines)
                avg_conf = sum(confidences) / len(confidences) if confidences else 0.0
                return full_text, round(avg_conf, 3)

            return "", 0.0
        except Exception as e:
            print(f"[OCR Processing Error]: {e}")
            return "", 0.0

    @classmethod
    def extract_text_from_image_file(cls, filepath: str) -> Tuple[str, float]:
        try:
            with open(filepath, "rb") as f:
                return cls.extract_text_from_image_bytes(f.read())
        except Exception as e:
            print(f"[OCR File Read Error {filepath}]: {e}")
            return "", 0.0
