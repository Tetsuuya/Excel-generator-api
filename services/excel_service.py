"""
Orchestration service for generating Excel files via LLMs (DeepSeek, Groq, Vercel AI Gateway, OpenAI-compatible APIs)
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
DEFAULT_OPENROUTER_MODELS = [
    "openrouter/free",
    "cohere/north-mini-code:free",
    "google/gemma-4-31b-it:free",
    "nvidia/nemotron-3-super-120b-a12b:free",
    "z-ai/glm-5.2:free"
]

DEFAULT_DEEPSEEK_MODELS = [
    "deepseek-chat",
    "deepseek-reasoner"
]

DEFAULT_GROQ_MODELS = [
    "qwen/qwen3.6-27b",
    "openai/gpt-oss-120b",
    "groq/compound",
    "openai/gpt-oss-20b"
]


def extract_python_code(raw_text: str) -> str:
    """
    Extracts pure Python code from markdown code fences or raw string,
    stripping any <think>...</think> tags and extraneous conversational text.
    """
    if not raw_text:
        return ""
    
    # 1. Strip reasoning / thinking tags (e.g. from DeepSeek / Qwen models)
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
        self.base_url = base_url or os.getenv("OPENAI_BASE_URL")
        
        # Determine API key based on configured provider
        if self.base_url and "openrouter" in self.base_url.lower():
            self.api_key = api_key or os.getenv("OPENAI_API_KEY")
            if not default_model:
                self.default_model = "openrouter/free"
        elif self.base_url and "deepseek" in self.base_url.lower():
            self.api_key = api_key or os.getenv("DEEPSEEK_API_KEY") or os.getenv("OPENAI_API_KEY")
            if not default_model:
                self.default_model = "deepseek-chat"
        elif self.base_url and "groq" in self.base_url.lower():
            self.api_key = api_key or os.getenv("GROQ_API_KEY") or os.getenv("OPENAI_API_KEY")
            if not default_model:
                self.default_model = "qwen/qwen3.6-27b"
        else:
            if os.getenv("OPENAI_API_KEY"):
                self.api_key = api_key or os.getenv("OPENAI_API_KEY")
                self.base_url = self.base_url or "https://openrouter.ai/api/v1"
                self.default_model = default_model or "openrouter/free"
            elif os.getenv("DEEPSEEK_API_KEY"):
                self.api_key = api_key or os.getenv("DEEPSEEK_API_KEY")
                self.base_url = self.base_url or "https://api.deepseek.com"
                self.default_model = default_model or "deepseek-chat"
            elif os.getenv("GROQ_API_KEY"):
                self.api_key = api_key or os.getenv("GROQ_API_KEY")
                self.base_url = self.base_url or "https://api.groq.com/openai/v1"
                self.default_model = default_model or "qwen/qwen3.6-27b"
            else:
                self.api_key = api_key
                self.base_url = self.base_url or "https://openrouter.ai/api/v1"
                self.default_model = default_model or "openrouter/free"

        if default_model:
            self.default_model = default_model

        self._client: Optional[OpenAI] = None

    @property
    def client(self) -> OpenAI:
        if self._client is None:
            if not self.api_key:
                raise ValueError(
                    "No API Key provided. Please set OPENAI_API_KEY, GROQ_API_KEY, or DEEPSEEK_API_KEY in your .env file."
                )
            self._client = OpenAI(
                api_key=self.api_key,
                base_url=self.base_url
            )
        return self._client

    def get_candidate_models(self, preferred_model: Optional[str] = None) -> List[str]:
        """Returns ordered list of fallback models depending on provider."""
        models: List[str] = []
        if preferred_model:
            models.append(preferred_model)
        
        base = (self.base_url or "").lower()
        if "openrouter" in base or (preferred_model and ":free" in preferred_model):
            for m in DEFAULT_OPENROUTER_MODELS:
                if m not in models:
                    models.append(m)
        elif "deepseek.com" in base or (preferred_model and "deepseek" in preferred_model):
            for m in DEFAULT_DEEPSEEK_MODELS:
                if m not in models:
                    models.append(m)
        else:
            for m in DEFAULT_GROQ_MODELS:
                if m not in models:
                    models.append(m)
                    
        return models

    def generate_excel(
        self,
        prompt: str,
        preferred_model: Optional[str] = None,
        max_retries: int = 2
    ) -> Dict[str, Any]:
        """
        Generates an Excel spreadsheet from a natural language prompt with automated AST verification and self-healing.
        """
        start_time = time.time()
        
        selected_model = preferred_model or self.default_model or "qwen/qwen3.6-27b"
        models_to_try = self.get_candidate_models(selected_model)
        
        last_error = ""
        last_code = ""
        current_model = models_to_try[0]
        
        for attempt in range(1, max_retries + 2):
            messages = [
                {"role": "system", "content": EXCEL_SYSTEM_PROMPT},
                {"role": "user", "content": prompt}
            ]
            if last_error:
                messages.append({
                    "role": "user",
                    "content": f"Your previous code failed with error: {last_error[:300]}\n\nPrevious Code:\n```python\n{last_code}\n```\n\nPlease fix the issue and return the complete, working Python script inside ```python ... ``` ending with wb.save(output_path)."
                })

            response_text = ""
            for model_candidate in models_to_try:
                try:
                    current_model = model_candidate
                    model_max_tokens = 7500
                    
                    completion = self.client.chat.completions.create(
                        model=current_model,
                        messages=messages,
                        temperature=0.1,
                        max_tokens=model_max_tokens,
                    )
                    response_text = completion.choices[0].message.content or ""
                    if response_text:
                        break
                except Exception as e:
                    err_str = str(e).lower()
                    if any(k in err_str for k in ["rate_limit", "429", "413", "request_too_large", "overloaded", "tpm", "rpm"]):
                        wait_match = re.search(r"try again in ([\d\.]+)s", err_str)
                        wait_sec = float(wait_match.group(1)) + 0.5 if wait_match else 1.5
                        print(f"[Fallback] Model '{model_candidate}' hit rate limit. Waiting {wait_sec:.1f}s and trying next candidate...")
                        time.sleep(min(wait_sec, 6.0))
                        continue
                    elif "decommissioned" in err_str or "not found" in err_str:
                        print(f"[Fallback] Model '{model_candidate}' unavailable on this endpoint. Skipping to next candidate...")
                        continue
                    else:
                        raise e

            if not response_text:
                if attempt <= max_retries:
                    print(f"[Retry Attempt {attempt}] Waiting 3s for rate limit window to clear...")
                    time.sleep(3.0)
                    continue
                raise RuntimeError(
                    f"Failed to receive response from LLM endpoint ({self.base_url}). "
                    "Please check your API key and connection."
                )

            code = extract_python_code(response_text)
            last_code = code

            try:
                # Sanitize and execute the generated Python code
                excel_bytes = execute_excel_code(code, timeout_seconds=30)
                
                duration = round(time.time() - start_time, 2)
                return {
                    "excel_bytes": excel_bytes,
                    "code": code,
                    "model": current_model,
                    "attempts": attempt,
                    "duration_seconds": duration,
                    "success": True
                }

            except (SecurityError, ExecutionError, Exception) as err:
                last_error = str(err)
                if attempt > max_retries:
                    break

        raise RuntimeError(
            f"Failed to generate valid Excel workbook after {max_retries + 1} attempts.\n"
            f"Last Error: {last_error}\n"
            f"Last Code:\n{last_code}"
        )

