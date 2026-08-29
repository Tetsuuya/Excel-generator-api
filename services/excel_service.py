"""
Orchestration service for generating Excel files via LLMs (Groq LPUs, OpenRouter, DeepSeek)
with AST sanitization, isolated sandbox execution, and an automated self-healing error correction loop.
"""

import os
import re
import time
from typing import Dict, Any, Tuple, Optional, List
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

from core.prompts import EXCEL_SYSTEM_PROMPT, SELF_HEALING_PROMPT_TEMPLATE
from core.executor import execute_excel_code, ExecutionError
from core.sanitizer import SecurityError


# Default supported models across providers
DEFAULT_GROQ_MODELS = [
    "qwen/qwen3.8-27b",
    "openai/gpt-oss-120b",
    "qwen/qwen3.6-27b",
    "groq/compound",
    "openai/gpt-oss-20b"
]

DEFAULT_OPENROUTER_MODELS = [
    "openrouter/free",
    "google/gemma-4-31b-it:free",
    "cohere/north-mini-code:free",
    "nvidia/nemotron-3-super-120b-a12b:free",
    "z-ai/glm-5.2:free"
]


def extract_python_code(raw_text: str) -> str:
    """
    Extracts pure Python code from markdown code fences or raw string,
    stripping any <think>...</think> tags and extraneous conversational text.
    """
    if not raw_text:
        return ""
    
    # 1. Strip reasoning / thinking tags (e.g. from DeepSeek / Qwen / Groq R1 models)
    cleaned = re.sub(r"<think>[\s\S]*?</think>", "", raw_text, flags=re.IGNORECASE).strip()
    
    # 2. Match ```python ... ```
    match = re.search(r"```(?:python)?\s*([\s\S]*?)\s*```", cleaned, re.IGNORECASE)
    if match:
        extracted = match.group(1).strip()
        if "def generate_excel" in extracted:
            return extracted

    # 3. Match any fenced code block containing def generate_excel
    all_blocks = re.findall(r"```(?:[a-zA-Z0-9_-]+)?\s*([\s\S]*?)\s*```", cleaned)
    for block in all_blocks:
        if "def generate_excel" in block:
            return block.strip()

    # 4. If no fences, find the starting import or function definition
    if "def generate_excel" in cleaned:
        import_pos = cleaned.find("import ")
        def_pos = cleaned.find("def generate_excel")
        start_idx = import_pos if (import_pos != -1 and import_pos < def_pos) else def_pos
        return cleaned[start_idx:].strip()

    return cleaned


class ExcelService:
    def __init__(
        self,
        api_key: Optional[str] = None,
        base_url: Optional[str] = None,
        default_model: Optional[str] = None
    ):
        self.groq_key = os.getenv("GROQ_API_KEY")
        self.openrouter_key = os.getenv("OPENAI_API_KEY")
        
        # Primary provider setup (Prefer Groq for blazing speed if key is present)
        if self.groq_key:
            self.primary_provider = "Groq"
            self.base_url = "https://api.groq.com/openai/v1"
            self.api_key = self.groq_key
            self.default_model = default_model or "qwen/qwen3.8-27b"
        elif self.openrouter_key:
            self.primary_provider = "OpenRouter"
            self.base_url = os.getenv("OPENAI_BASE_URL", "https://openrouter.ai/api/v1")
            self.api_key = self.openrouter_key
            self.default_model = default_model or "openrouter/free"
        else:
            self.primary_provider = "Custom"
            self.base_url = base_url or "https://openrouter.ai/api/v1"
            self.api_key = api_key
            self.default_model = default_model or "openrouter/free"

        # Clients for multi-provider fallback
        self._groq_client: Optional[OpenAI] = None
        self._openrouter_client: Optional[OpenAI] = None

    @property
    def groq_client(self) -> Optional[OpenAI]:
        if self._groq_client is None and self.groq_key:
            self._groq_client = OpenAI(
                api_key=self.groq_key,
                base_url="https://api.groq.com/openai/v1"
            )
        return self._groq_client

    @property
    def openrouter_client(self) -> Optional[OpenAI]:
        if self._openrouter_client is None and self.openrouter_key:
            self._openrouter_client = OpenAI(
                api_key=self.openrouter_key,
                base_url=os.getenv("OPENAI_BASE_URL", "https://openrouter.ai/api/v1")
            )
        return self._openrouter_client

    def get_candidate_models(self, preferred_model: Optional[str] = None) -> List[str]:
        """Returns list of model names available in the fallback cascade."""
        targets = self.get_candidate_targets(preferred_model)
        return [f"{t[0]}:{t[1]}" for t in targets]

    def get_candidate_targets(self, preferred_model: Optional[str] = None) -> List[Tuple[str, str, OpenAI]]:
        """
        Returns an ordered list of (provider_name, model_name, client_instance) targets for resilient cascade.
        """
        targets: List[Tuple[str, str, OpenAI]] = []
        
        # 1. Preferred model if specified
        if preferred_model:
            if self.groq_client and any(m in preferred_model.lower() for m in ["qwen", "groq", "gpt-oss", "llama"]):
                targets.append(("Groq", preferred_model, self.groq_client))
            elif self.openrouter_client:
                targets.append(("OpenRouter", preferred_model, self.openrouter_client))

        # 2. Add Groq models (Blazing fast LPUs)
        if self.groq_client:
            for gm in DEFAULT_GROQ_MODELS:
                if not any(t[1] == gm for t in targets):
                    targets.append(("Groq", gm, self.groq_client))

        # 3. Add OpenRouter models as secondary backup
        if self.openrouter_client:
            for om in DEFAULT_OPENROUTER_MODELS:
                if not any(t[1] == om for t in targets):
                    targets.append(("OpenRouter", om, self.openrouter_client))

        return targets

    def generate_excel(
        self,
        prompt: str,
        preferred_model: Optional[str] = None,
        max_retries: int = 2
    ) -> Dict[str, Any]:
        """
        Generates an Excel spreadsheet with automated multi-provider fallback and self-healing.
        """
        start_time = time.time()
        targets = self.get_candidate_targets(preferred_model)
        
        if not targets:
            raise ValueError("No LLM API keys configured. Please set GROQ_API_KEY or OPENAI_API_KEY in your .env.")

        last_error = ""
        last_code = ""
        used_model = targets[0][1]
        used_provider = targets[0][0]

        for attempt in range(1, max_retries + 2):
            messages = [
                {"role": "system", "content": EXCEL_SYSTEM_PROMPT},
                {"role": "user", "content": prompt}
            ]
            if last_error:
                messages.append({
                    "role": "user",
                    "content": f"Your previous code failed with error:\n{last_error[:300]}\n\nPrevious Code:\n```python\n{last_code}\n```\n\nPlease fix the issue and return the complete, working Python script inside ```python ... ``` ending with wb.save(output_path)."
                })

            response_text = ""
            for provider_name, model_name, client in targets:
                try:
                    used_model = model_name
                    used_provider = provider_name
                    
                    completion = client.chat.completions.create(
                        model=model_name,
                        messages=messages,
                        temperature=0.1,
                        max_tokens=6500,
                    )
                    
                    if completion and getattr(completion, "choices", None) and len(completion.choices) > 0:
                        first_choice = completion.choices[0]
                        if hasattr(first_choice, "message") and first_choice.message:
                            response_text = first_choice.message.content or ""
                    
                    if response_text and response_text.strip():
                        break
                    else:
                        print(f"[Fallback] {provider_name} '{model_name}' returned empty content. Trying next...")
                except Exception as e:
                    err_str = str(e).lower()
                    print(f"[Fallback] {provider_name} '{model_name}' error: {e}. Trying next candidate...")
                    if any(k in err_str for k in ["rate_limit", "429", "413", "overloaded", "tpm", "rpm"]):
                        wait_match = re.search(r"try again in ([\d\.]+)s", err_str)
                        wait_sec = float(wait_match.group(1)) + 0.5 if wait_match else 1.0
                        time.sleep(min(wait_sec, 4.0))
                    continue

            if not response_text:
                if attempt <= max_retries:
                    time.sleep(2.0)
                    continue
                raise RuntimeError(
                    f"Failed to receive response from any LLM provider (Groq / OpenRouter). "
                    "Please check your API keys and internet connection."
                )

            code = extract_python_code(response_text)
            last_code = code

            try:
                # Sanitize and execute in sandbox with 60s window
                excel_bytes = execute_excel_code(code, timeout_seconds=60)
                
                duration = round(time.time() - start_time, 2)
                return {
                    "excel_bytes": excel_bytes,
                    "code": code,
                    "model": f"{used_provider}:{used_model}",
                    "attempts": attempt,
                    "duration_seconds": duration,
                    "success": True
                }

            except (SecurityError, ExecutionError, Exception) as err:
                last_error = str(err)
                print(f"[Self-Healing Triggered] Execution error on attempt {attempt}: {err}")
                if attempt > max_retries:
                    break

        raise RuntimeError(
            f"Failed to generate valid Excel workbook after {max_retries + 1} attempts.\n"
            f"Last Error: {last_error}\n"
            f"Last Code:\n{last_code}"
        )
