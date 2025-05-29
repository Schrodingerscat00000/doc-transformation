# ollama_client.py
import requests
import logging

class OllamaClient:
    def __init__(self, model="deepseek-r1:1.5b", base_url="http://localhost:11434"):
        self.model = model
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()

    def is_available(self) -> bool:
        """Check if Ollama is running and the model is available."""
        try:
            resp = self.session.get(f"{self.base_url}/api/tags", timeout=5)
            if resp.status_code != 200:
                logging.error(f"Ollama /api/tags returned {resp.status_code}")
                return False

            models = resp.json().get("models", [])
            # models is a list of dicts with "name", "tag", etc.
            return any(m.get("name", "").startswith(self.model.split(":")[0]) for m in models)
        except Exception as e:
            logging.error(f"Ollama availability check failed: {e}")
            return False

    def query(self, prompt: str, max_tokens: int = 500) -> str:
        """Send a completion request to Ollama and return the generated text."""
        payload = {
            "model": self.model,
            "prompt": prompt,
            "stream": False,
            "options": {
                "num_predict": max_tokens,
                "temperature": 0.1
            }
        }
        try:
            resp = self.session.post(
                f"{self.base_url}/api/generate",
                json=payload,
                timeout=30
            )
            if resp.status_code != 200:
                logging.error(f"Ollama /api/generate returned {resp.status_code}: {resp.text}")
                return None

            data = resp.json()
            # Ollama’s JSON has a top-level "response" field
            return data.get("response", "").strip()
        except Exception as e:
            logging.error(f"Ollama query error: {e}")
            return None
