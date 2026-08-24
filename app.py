"""
FastAPI application for AI-Powered Secure Excel Generation.
"""

import os
import re
from io import BytesIO
from typing import Optional
from fastapi import FastAPI, HTTPException, status
from fastapi.responses import Response, JSONResponse, StreamingResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from dotenv import load_dotenv

from core.sanitizer import sanitize_python_code, SecurityError
from core.executor import execute_excel_code, ExecutionError
from services.excel_service import ExcelService

load_dotenv()

app = FastAPI(
    title="AI-Powered Secure Python Excel Generator",
    description="Generates executive-ready Excel workbooks (.xlsx) from natural language prompts using Groq LPUs, AST sanitization, and isolated execution.",
    version="1.0.0"
)

# Enable CORS for cross-service calls (e.g., from Node.js Docs-Service or frontend)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition", "X-Generated-Model", "X-Attempts-Required", "X-Duration-Seconds", "*"],
)

excel_service = ExcelService()


class GenerateExcelRequest(BaseModel):
    prompt: str = Field(..., description="Description of the spreadsheet, tables, formulas, and charts to create.")
    model: str = Field(default="groq/compound", description="Groq model to use.")
    filename: Optional[str] = Field(default="generated_report.xlsx", description="Suggested filename for download.")


class ExecuteCustomCodeRequest(BaseModel):
    code: str = Field(..., description="Raw Python code containing generate_excel(output_path) function.")
    filename: Optional[str] = Field(default="custom_report.xlsx", description="Suggested filename for download.")


@app.get("/", include_in_schema=False)
def serve_index_html():
    index_path = os.path.join(os.path.dirname(__file__), "public", "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path, media_type="text/html")
    return {"message": "AI Excel Generator API Online"}


@app.get("/health", tags=["System"])
def health_check():
    has_api_key = bool(os.getenv("GROQ_API_KEY"))
    return {
        "status": "online",
        "service": "excel-generator",
        "has_groq_api_key": has_api_key,
        "default_model": "groq/compound",
        "available_models": ["groq/compound", "openai/gpt-oss-120b", "qwen/qwen3.6-27b", "openai/gpt-oss-20b"],
        "sandbox": "AST_sanitizer + isolated_process"
    }


@app.post("/api/generate-excel", tags=["Excel Generation"])
def generate_excel_endpoint(req: GenerateExcelRequest):
    """
    Generates an Excel workbook from a prompt and streams the binary .xlsx file directly.
    """
    try:
        result = excel_service.generate_excel(
            prompt=req.prompt,
            preferred_model=req.model
        )
        
        # Ensure clean filename
        raw_name = req.filename or "generated_report.xlsx"
        clean_name = re.sub(r'[\\/*?:"<>|]', "", raw_name).strip() or "generated_report.xlsx"
        if not clean_name.lower().endswith(".xlsx"):
            clean_name += ".xlsx"
        
        return Response(
            content=result["excel_bytes"],
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f'attachment; filename="{clean_name}"',
                "X-Generated-Model": str(result["model"]),
                "X-Attempts-Required": str(result["attempts"]),
                "X-Duration-Seconds": str(result["duration_seconds"])
            }
        )
    except (SecurityError, ExecutionError, ValueError) as err:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(err))
    except Exception as err:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=f"Generation failed: {str(err)}")


@app.post("/api/generate-excel/preview", tags=["Excel Generation"])
def preview_excel_generation(req: GenerateExcelRequest):
    """
    Generates Excel code, validates it, and returns the generated Python code and metadata without downloading.
    """
    try:
        result = excel_service.generate_excel(
            prompt=req.prompt,
            preferred_model=req.model
        )
        
        return {
            "success": True,
            "model": result["model"],
            "attempts": result["attempts"],
            "duration_seconds": result["duration_seconds"],
            "code": result["code"],
            "file_size_bytes": len(result["excel_bytes"])
        }
    except Exception as err:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(err))


@app.post("/api/execute-code", tags=["Direct Execution"])
def execute_code_endpoint(req: ExecuteCustomCodeRequest):
    """
    Directly sanitizes and executes provided Python code in the secure sandbox.
    """
    try:
        excel_bytes = execute_excel_code(req.code)
        filename = req.filename if req.filename.endswith(".xlsx") else f"{req.filename}.xlsx"
        
        return Response(
            content=excel_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )
    except (SecurityError, ExecutionError) as err:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(err))
    except Exception as err:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(err))


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8001, reload=True)
