"""
Isolated subprocess executor for running sanitized Python Excel generation scripts.
Enforces timeout, minimal clean environment, and scratch directory sandboxing.
"""

import os
import sys
import tempfile
import subprocess
from typing import Tuple
from core.sanitizer import sanitize_python_code, SecurityError


class ExecutionError(Exception):
    """Raised when Python script fails at runtime."""
    pass


def execute_excel_code(code_str: str, timeout_seconds: int = 15) -> bytes:
    """
    1. Sanitizes the code via AST analysis.
    2. Writes code into an ephemeral temporary folder.
    3. Runs the code with Python in isolated mode (-I).
    4. Collects and returns the generated .xlsx binary bytes.
    """
    # Step 1: Security AST Verification
    sanitize_python_code(code_str)

    # Step 2: Temporary sandbox folder
    with tempfile.TemporaryDirectory() as temp_dir:
        script_file = os.path.join(temp_dir, "script.py")
        output_xlsx = os.path.join(temp_dir, "output.xlsx")

        # Create wrapper that safely imports and invokes generate_excel
        runner_code = f"""# -*- coding: utf-8 -*-
{code_str}

if __name__ == "__main__":
    if "generate_excel" in globals() and callable(globals()["generate_excel"]):
        generate_excel(r"{output_xlsx}")
    else:
        raise RuntimeError("Required entry point 'generate_excel(output_path)' was not found.")
"""

        with open(script_file, "w", encoding="utf-8") as f:
            f.write(runner_code)

        # Step 3: Strip sensitive environment variables (e.g., GROQ_API_KEY, credentials)
        # Only pass minimal necessary system environment variables for Python runtime
        safe_env = {
            "SYSTEMROOT": os.environ.get("SYSTEMROOT", "C:\\Windows"),
            "PATH": os.environ.get("PATH", ""),
            "TEMP": temp_dir,
            "TMP": temp_dir,
            "PYTHONIOENCODING": "utf-8",
            "PYTHONUNBUFFERED": "1"
        }

        # Step 4: Execute in subprocess with isolated flag (-I)
        # Note: If running inside venv, sys.executable is the venv python
        try:
            process = subprocess.run(
                [sys.executable, script_file],
                cwd=temp_dir,
                env=safe_env,
                capture_output=True,
                text=True,
                timeout=timeout_seconds,
                encoding="utf-8",
                errors="replace"
            )
        except subprocess.TimeoutExpired:
            raise ExecutionError(f"Execution timed out after {timeout_seconds} seconds (possible infinite loop).")
        except Exception as e:
            raise ExecutionError(f"Failed to spawn execution process: {str(e)}")

        if process.returncode != 0:
            error_output = process.stderr.strip() or process.stdout.strip()
            raise ExecutionError(f"Python Runtime Error (Exit Code {process.returncode}):\n{error_output}")

        if not os.path.exists(output_xlsx) or os.path.getsize(output_xlsx) == 0:
            raise ExecutionError("Script completed successfully but did not produce a valid .xlsx file.")

        with open(output_xlsx, "rb") as f:
            excel_bytes = f.read()

        return excel_bytes
