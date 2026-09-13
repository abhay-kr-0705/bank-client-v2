"""
OpenRouter Multimodal AI Client for Banking Valuation Document Intelligence.
Supports vision OCR, few-shot prompt injection, confidence evaluation,
and robust JSON schema extraction.
"""

import os
import json
import base64
import requests
from typing import Dict, Any, List, Optional

def get_env_variable(key: str, default: str = "") -> str:
    val = os.environ.get(key)
    if val:
        return val.strip()
    env_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env")
    if os.path.exists(env_path):
        try:
            with open(env_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line.startswith(f"{key}="):
                        return line.split("=", 1)[1].strip().strip('"').strip("'")
        except Exception:
            pass
    return default

DEFAULT_OPENROUTER_KEY = get_env_variable("OPENROUTER_API_KEY", "")
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
DEFAULT_MODEL = "google/gemini-2.5-flash"

class OpenRouterClient:
    def __init__(self, api_key: Optional[str] = None, model: str = DEFAULT_MODEL):
        self.api_key = (api_key or os.environ.get("OPENROUTER_API_KEY") or DEFAULT_OPENROUTER_KEY or get_env_variable("OPENROUTER_API_KEY")).strip()
        self.model = model or DEFAULT_MODEL

    def extract_valuation_fields(
        self,
        prompt_text: str,
        image_bytes_list: Optional[List[bytes]] = None,
        max_tokens: int = 3800,
        temperature: float = 0.1
    ) -> Optional[Dict[str, Any]]:
        """
        Sends multimodal document contents to OpenRouter and returns structured JSON with field confidence.
        """
        if not self.api_key:
            print("[OpenRouter] No API key provided.")
            return None

        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "http://127.0.0.1:8000",
            "X-Title": "BankTech Valuation OCR"
        }

        # Build message contents
        content_items: List[Dict[str, Any]] = [
            {"type": "text", "text": prompt_text}
        ]

        if image_bytes_list:
            # Send up to 4 high-priority document/site images to fit within token budgets
            for idx, img_b in enumerate(image_bytes_list[:4]):
                try:
                    b64_img = base64.b64encode(img_b).decode("utf-8")
                    content_items.append({
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{b64_img}"
                        }
                    })
                except Exception as img_err:
                    print(f"[OpenRouter] Image encoding error on image {idx}: {img_err}")

        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "user",
                    "content": content_items
                }
            ],
            "max_tokens": max_tokens,
            "temperature": temperature,
            "response_format": {"type": "json_object"}
        }

        try:
            response = requests.post(OPENROUTER_URL, headers=headers, json=payload, timeout=60)
            if response.status_code != 200:
                print(f"[OpenRouter Error {response.status_code}]: {response.text}")
                # Fallback check if response format caused an issue
                if "response_format" in response.text:
                    del payload["response_format"]
                    retry_resp = requests.post(OPENROUTER_URL, headers=headers, json=payload, timeout=60)
                    if retry_resp.status_code == 200:
                        response = retry_resp
                    else:
                        return None
                else:
                    return None

            res_json = response.json()
            choices = res_json.get("choices", [])
            if not choices:
                return None

            raw_content = choices[0].get("message", {}).get("content", "")
            if not raw_content:
                return None

            # Clean markdown codeblocks if wrapped
            clean_str = raw_content.strip()
            if clean_str.startswith("```json"):
                clean_str = clean_str[7:]
            elif clean_str.startswith("```"):
                clean_str = clean_str[3:]
            if clean_str.endswith("```"):
                clean_str = clean_str[:-3]
            clean_str = clean_str.strip()

            parsed = json.loads(clean_str)
            return parsed
        except Exception as e:
            print(f"[OpenRouter Exception]: {e}")
            return None
