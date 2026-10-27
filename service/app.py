#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BIDV Meeting Minutes Service
Web UI & REST API for generating administrative meeting minutes & conclusions
complying with BIDV standards and powered by gemini-stealth-bridge (GSB).
"""

import os
import shutil
import uuid
import datetime
import base64
import requests
from typing import Optional, List
from pydantic import BaseModel, Field
from fastapi import FastAPI, File, UploadFile, Form, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

import bidv_engine

app = FastAPI(
    title="BIDV Meeting Minutes & Conclusions API Service",
    description="REST API dịch vụ tổng hợp Biên bản cuộc họp & Thông báo kết luận chuẩn thể thức hành chính BIDV (kèm xuất file .docx và .pdf)",
    version="1.1.0",
    docs_url="/docs",
    redoc_url="/redoc"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

OUTPUT_DIR = os.environ.get("OUTPUT_DIR", "/app/outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)

GSB_URL = os.environ.get("GSB_URL", "http://gsb:20133/v1/chat/completions")
GSB_BASE_URL = GSB_URL.split("/v1/")[0] if "/v1/" in GSB_URL else "http://gsb:20133"
GSB_API_KEY = os.environ.get("GSB_API_KEY", "public")
DEFAULT_MODEL = os.environ.get("DEFAULT_MODEL", "gemini-argon")


# -------------------------------------------------------------
# Data Models for Client API
# -------------------------------------------------------------
class GenerateMinutesJSONRequest(BaseModel):
    transcript: str = Field(
        ..., 
        description="Toàn bộ nội dung bản ghi chép chi tiết, phân vai hoặc transcript cuộc họp"
    )
    doc_type: str = Field(
        "BB", 
        description="Loại văn bản hành chính: 'BB' (Biên bản cuộc họp) hoặc 'TBKL' (Thông báo kết luận)"
    )
    meeting_title: Optional[str] = Field(
        "", 
        description="Tiêu đề / trích yếu phiên họp (nếu để trống sẽ dùng tiêu đề mặc định BIDV)"
    )
    doc_number: Optional[str] = Field(
        "2710", 
        description="Số hiệu văn bản (ví dụ: '2710' -> 2710/BB-BIDV hoặc 2710/TB-BIDV)"
    )
    signer_name: Optional[str] = Field(
        "HOÀNG VIỆT HÙNG", 
        description="Họ tên người ký / Lãnh đạo chủ trì"
    )
    signer_title: Optional[str] = Field(
        "ĐẠI DIỆN ĐỒNG CHỦ TRÌ PHIÊN HỌP\nTỔNG GIÁM ĐỐC BIDV", 
        description="Chức danh người ký"
    )
    recipients: Optional[List[str]] = Field(
        None,
        description="Danh sách các đơn vị nhận (Nơi nhận), ví dụ: ['- Ban Lãnh đạo...', '- Lưu: VT, VP...']"
    )
    model: Optional[str] = Field(
        "gemini-argon", 
        description="Mô hình AI: 'gemini-argon' (khuyên dùng), 'gemini-pro', 'gemini-flash', 'gemini-flash-lite'"
    )
    include_base64: Optional[bool] = Field(
        False, 
        description="Nếu true, trả về trực tiếp chuỗi base64 của file .docx và .pdf trong JSON response"
    )

    class Config:
        json_schema_extra = {
            "example": {
                "transcript": "Vào lúc 09h00 ngày 27/10/2026, tại Trụ sở chính BIDV, diễn ra cuộc họp giữa BIDV và Tencent. Đồng chí Hoàng Việt Hùng - Tổng Giám đốc BIDV chủ trì...",
                "doc_type": "BB",
                "meeting_title": "Phiên họp Cấp cao Song phương BIDV - Tencent về Hợp tác Chiến lược",
                "doc_number": "2710",
                "signer_name": "HOÀNG VIỆT HÙNG",
                "signer_title": "ĐẠI DIỆN ĐỒNG CHỦ TRÌ PHIÊN HỌP\nTỔNG GIÁM ĐỐC BIDV",
                "model": "gemini-argon",
                "include_base64": False
            }
        }


def _process_generation(
    transcript_text: str,
    doc_type: str = "BB",
    meeting_title: str = "",
    doc_number: str = "2710",
    signer_name: str = "HOÀNG VIỆT HÙNG",
    signer_title: str = "ĐẠI DIỆN ĐỒNG CHỦ TRÌ PHIÊN HỌP\nTỔNG GIÁM ĐỐC BIDV",
    recipients: Optional[List[str]] = None,
    model: str = "gemini-argon",
    include_base64: bool = False,
    request_host: str = ""
):
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    doc_prefix = "Bien_ban_hop" if doc_type == "BB" else "Thong_bao_ket_luan"
    base_name = f"{doc_prefix}_BIDV_{timestamp}_{uuid.uuid4().hex[:6]}"
    docx_path = os.path.join(OUTPUT_DIR, f"{base_name}.docx")

    # 1. Call Gemini via GSB
    ai_response = bidv_engine.call_ai(transcript_text, model=model, doc_type=doc_type)

    # 2. Build docx & convert to pdf
    out_docx, out_pdf = bidv_engine.build_docx_and_pdf(
        ai_text=ai_response,
        docx_path=docx_path,
        doc_type=doc_type,
        meeting_title=meeting_title,
        doc_number=doc_number,
        signer_title=signer_title,
        signer_name=signer_name,
        recipients=recipients
    )

    docx_filename = os.path.basename(out_docx)
    pdf_filename = os.path.basename(out_pdf) if out_pdf else None

    # Base URL builder
    host_prefix = f"http://{request_host}" if request_host else ""
    docx_download_url = f"{host_prefix}/api/download/{docx_filename}"
    pdf_download_url = f"{host_prefix}/api/download/{pdf_filename}" if pdf_filename else None

    res = {
        "status": "success",
        "message": "Đã tạo biên bản thành công theo đúng thể thức chuẩn BIDV",
        "model_used": model,
        "doc_type": doc_type,
        "docx_download_url": docx_download_url,
        "pdf_download_url": pdf_download_url,
        "docx_filename": docx_filename,
        "pdf_filename": pdf_filename,
        "text_content": ai_response,
        "preview_text": ai_response[:1000] + ("..." if len(ai_response) > 1000 else "")
    }

    if include_base64:
        with open(out_docx, "rb") as f_d:
            res["docx_base64"] = base64.b64encode(f_d.read()).decode("utf-8")
        if out_pdf and os.path.exists(out_pdf):
            with open(out_pdf, "rb") as f_p:
                res["pdf_base64"] = base64.b64encode(f_p.read()).decode("utf-8")
        else:
            res["pdf_base64"] = None

    return res


# -------------------------------------------------------------
# REST API Endpoints for Client Applications
# -------------------------------------------------------------
@app.get("/api/health", tags=["System"])
def health_check():
    """Kiểm tra tình trạng hoạt động của service và kết nối tới Gemini Stealth Bridge"""
    gsb_ok = False
    gsb_status = "unreachable"
    models = []
    try:
        r = requests.get(
            f"{GSB_BASE_URL}/v1/models",
            headers={"Authorization": f"Bearer {GSB_API_KEY}"},
            timeout=5
        )
        if r.status_code == 200:
            gsb_ok = True
            gsb_status = "connected"
            data = r.json()
            models = [m.get("id") for m in data.get("data", [])]
    except Exception as e:
        gsb_status = str(e)

    return {
        "status": "ok",
        "service": "bidv-minutes",
        "gsb_connected": gsb_ok,
        "gsb_status": gsb_status,
        "available_models": models,
        "default_model": DEFAULT_MODEL,
        "output_dir": OUTPUT_DIR
    }


@app.get("/api/v1/models", tags=["AI Models"])
def get_available_models():
    """Lấy danh sách các mô hình Gemini đang sẵn sàng trên Bridge"""
    try:
        r = requests.get(
            f"{GSB_BASE_URL}/v1/models",
            headers={"Authorization": f"Bearer {GSB_API_KEY}"},
            timeout=5
        )
        if r.status_code == 200:
            return r.json()
        raise HTTPException(status_code=r.status_code, detail=r.text)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Không thể lấy danh sách mô hình từ GSB: {str(e)}")


@app.post("/api/v1/generate", tags=["Minutes Generation"])
def generate_minutes_json(req: GenerateMinutesJSONRequest, request: Request):
    """
    **Primary API cho Client**: Gọi qua JSON Payload (`application/json`).
    Nhận transcript cuộc họp, tự động phân tích sâu với Gemini Argon và trả về link tải / chuỗi base64 file .docx & .pdf.
    """
    if not req.transcript or not req.transcript.strip():
        raise HTTPException(status_code=400, detail="Trường 'transcript' không được để trống.")

    try:
        return _process_generation(
            transcript_text=req.transcript.strip(),
            doc_type=req.doc_type,
            meeting_title=req.meeting_title or "",
            doc_number=req.doc_number or "2710",
            signer_name=req.signer_name or "HOÀNG VIỆT HÙNG",
            signer_title=req.signer_title or "ĐẠI DIỆN ĐỒNG CHỦ TRÌ PHIÊN HỌP\nTỔNG GIÁM ĐỐC BIDV",
            recipients=req.recipients,
            model=req.model or DEFAULT_MODEL,
            include_base64=req.include_base64 or False,
            request_host=request.headers.get("host", "")
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi xử lý tạo biên bản: {str(e)}")


@app.post("/api/v1/upload", tags=["Minutes Generation"])
async def generate_minutes_upload(
    request: Request,
    file: UploadFile = File(..., description="Tệp tài liệu cuộc họp (.docx, .txt, .md)"),
    doc_type: str = Form("BB", description="'BB' hoặc 'TBKL'"),
    meeting_title: str = Form("", description="Tiêu đề phiên họp"),
    doc_number: str = Form("2710", description="Số hiệu"),
    signer_name: str = Form("HOÀNG VIỆT HÙNG", description="Tên người ký"),
    signer_title: str = Form("ĐẠI DIỆN ĐỒNG CHỦ TRÌ PHIÊN HỌP\nTỔNG GIÁM ĐỐC BIDV", description="Chức vụ người ký"),
    model: str = Form("gemini-argon", description="Mô hình AI"),
    include_base64: bool = Form(False, description="Trả về base64")
):
    """
    **Upload API cho Client**: Tải lên trực tiếp file Word (.docx) hoặc file text ghi chép qua `multipart/form-data`.
    """
    temp_id = uuid.uuid4().hex[:8]
    temp_path = os.path.join(OUTPUT_DIR, f"temp_{temp_id}_{file.filename}")
    with open(temp_path, "wb") as f_out:
        shutil.copyfileobj(file.file, f_out)
    
    try:
        transcript_text = bidv_engine.extract_text_from_file(temp_path)
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

    if not transcript_text.strip():
        raise HTTPException(status_code=400, detail="Không thể trích xuất nội dung văn bản từ tệp tin tải lên.")

    try:
        return _process_generation(
            transcript_text=transcript_text,
            doc_type=doc_type,
            meeting_title=meeting_title,
            doc_number=doc_number,
            signer_name=signer_name,
            signer_title=signer_title,
            model=model,
            include_base64=include_base64,
            request_host=request.headers.get("host", "")
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi xử lý tạo biên bản: {str(e)}")


@app.get("/api/download/{filename}", tags=["File Downloads"])
def download_file(filename: str):
    """Tải về file tài liệu .docx hoặc .pdf đã sinh"""
    safe_name = os.path.basename(filename)
    path = os.path.join(OUTPUT_DIR, safe_name)
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="File không tồn tại trên hệ thống.")
    
    media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document" if safe_name.endswith(".docx") else "application/pdf"
    return FileResponse(path, media_type=media_type, filename=safe_name)


# Legacy endpoint support
@app.post("/api/generate", include_in_schema=False)
async def generate_minutes_legacy(
    request: Request,
    file: Optional[UploadFile] = File(None),
    raw_text: Optional[str] = Form(None),
    doc_type: str = Form("BB"),
    meeting_title: str = Form(""),
    doc_number: str = Form("2710"),
    signer_name: str = Form("HOÀNG VIỆT HÙNG"),
    signer_title: str = Form("ĐẠI DIỆN ĐỒNG CHỦ TRÌ PHIÊN HỌP\nTỔNG GIÁM ĐỐC BIDV"),
    model: str = Form("gemini-argon")
):
    transcript_text = ""
    if file and file.filename:
        temp_id = uuid.uuid4().hex[:8]
        temp_path = os.path.join(OUTPUT_DIR, f"temp_{temp_id}_{file.filename}")
        with open(temp_path, "wb") as f_out:
            shutil.copyfileobj(file.file, f_out)
        try:
            transcript_text = bidv_engine.extract_text_from_file(temp_path)
        finally:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass

    if not transcript_text and raw_text:
        transcript_text = raw_text.strip()

    if not transcript_text:
        raise HTTPException(status_code=400, detail="Vui lòng cung cấp nội dung transcript.")

    res = _process_generation(
        transcript_text=transcript_text,
        doc_type=doc_type,
        meeting_title=meeting_title,
        doc_number=doc_number,
        signer_name=signer_name,
        signer_title=signer_title,
        model=model,
        request_host=request.headers.get("host", "")
    )
    # Match legacy format
    res["docx_url"] = f"/api/download/{res['docx_filename']}"
    res["pdf_url"] = f"/api/download/{res['pdf_filename']}" if res['pdf_filename'] else None
    return res


@app.get("/", response_class=HTMLResponse)
def index_page():
    return """<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BIDV AI Minutes Generator</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bidv-green: #006a55;
            --bidv-green-dark: #004d3e;
            --bidv-green-light: #e6f4f1;
            --bidv-gold: #c5a059;
            --bidv-gold-dark: #a8853e;
            --gray-50: #f8fafc;
            --gray-100: #f1f5f9;
            --gray-200: #e2e8f0;
            --gray-700: #334155;
            --gray-900: #0f172a;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Inter', sans-serif; }
        body { background: var(--gray-50); color: var(--gray-900); min-height: 100vh; display: flex; flex-direction: column; }
        
        header {
            background: linear-gradient(135deg, var(--bidv-green-dark) 0%, var(--bidv-green) 100%);
            color: white;
            padding: 1.25rem 2rem;
            display: flex;
            align-items: center;
            justify-content: space-between;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
        }
        .brand { display: flex; align-items: center; gap: 1rem; }
        .brand-logo {
            background: white;
            color: var(--bidv-green);
            font-weight: 800;
            font-size: 1.25rem;
            padding: 0.4rem 0.8rem;
            border-radius: 6px;
            letter-spacing: 1px;
        }
        .brand-title { font-size: 1.15rem; font-weight: 600; }
        .brand-subtitle { font-size: 0.8rem; opacity: 0.85; }
        .bridge-status {
            display: flex;
            align-items: center;
            gap: 0.5rem;
            background: rgba(255,255,255,0.12);
            padding: 0.4rem 0.8rem;
            border-radius: 20px;
            font-size: 0.85rem;
        }
        .status-dot {
            width: 9px;
            height: 9px;
            border-radius: 50%;
            background: #22c55e;
            box-shadow: 0 0 8px #22c55e;
        }

        main { max-width: 1000px; margin: 2rem auto; padding: 0 1rem; flex: 1; width: 100%; }

        .card {
            background: white;
            border-radius: 12px;
            border: 1px solid var(--gray-200);
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            padding: 2rem;
            margin-bottom: 2rem;
        }
        .card-header {
            border-bottom: 2px solid var(--bidv-green-light);
            padding-bottom: 1rem;
            margin-bottom: 1.5rem;
        }
        .card-title {
            font-size: 1.25rem;
            font-weight: 700;
            color: var(--bidv-green);
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .form-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 1.25rem;
            margin-bottom: 1.5rem;
        }
        .full-width { grid-column: span 2; }

        label {
            display: block;
            font-size: 0.85rem;
            font-weight: 600;
            color: var(--gray-700);
            margin-bottom: 0.4rem;
        }
        input, select, textarea {
            width: 100%;
            padding: 0.65rem 0.85rem;
            border: 1px solid var(--gray-200);
            border-radius: 6px;
            font-size: 0.9rem;
            outline: none;
            transition: border-color 0.15s;
        }
        input:focus, select:focus, textarea:focus {
            border-color: var(--bidv-green);
            box-shadow: 0 0 0 3px rgba(0,106,85,0.15);
        }

        /* Dropzone */
        .dropzone {
            border: 2px dashed #cbd5e1;
            border-radius: 8px;
            padding: 2rem 1rem;
            text-align: center;
            background: #f8fafc;
            cursor: pointer;
            transition: all 0.2s;
        }
        .dropzone:hover, .dropzone.dragover {
            border-color: var(--bidv-green);
            background: var(--bidv-green-light);
        }
        .dropzone-icon { font-size: 2.2rem; margin-bottom: 0.5rem; }
        .dropzone-text { font-size: 0.95rem; font-weight: 500; color: var(--gray-700); }
        .dropzone-sub { font-size: 0.8rem; color: #64748b; margin-top: 0.25rem; }
        .file-chosen {
            display: none;
            margin-top: 0.75rem;
            background: white;
            border: 1px solid var(--bidv-green);
            border-radius: 6px;
            padding: 0.5rem 1rem;
            font-size: 0.85rem;
            font-weight: 600;
            color: var(--bidv-green);
        }

        .btn-submit {
            background: linear-gradient(135deg, var(--bidv-green) 0%, var(--bidv-green-dark) 100%);
            color: white;
            border: none;
            padding: 0.85rem 2rem;
            border-radius: 8px;
            font-size: 1rem;
            font-weight: 600;
            cursor: pointer;
            width: 100%;
            box-shadow: 0 4px 6px -1px rgba(0,106,85,0.25);
            transition: transform 0.1s, opacity 0.2s;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.5rem;
        }
        .btn-submit:hover { opacity: 0.95; }
        .btn-submit:active { transform: scale(0.99); }
        .btn-submit:disabled { opacity: 0.6; cursor: not-allowed; }

        /* Status & Results */
        .status-box {
            display: none;
            margin-top: 1.5rem;
            padding: 1.25rem;
            border-radius: 8px;
            background: #f1f5f9;
            border-left: 4px solid var(--bidv-green);
        }
        .status-spinner {
            display: inline-block;
            width: 18px;
            height: 18px;
            border: 3px solid rgba(0,106,85,0.2);
            border-top-color: var(--bidv-green);
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            vertical-align: middle;
            margin-right: 0.5rem;
        }
        @keyframes spin { to { transform: rotate(360deg); } }

        .result-box {
            display: none;
            margin-top: 1.5rem;
            background: #f0fdf4;
            border: 1px solid #bbf7d0;
            border-radius: 8px;
            padding: 1.5rem;
        }
        .result-title {
            color: #166534;
            font-size: 1.1rem;
            font-weight: 700;
            margin-bottom: 0.75rem;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }
        .btn-group { display: flex; gap: 1rem; margin-top: 1rem; }
        .btn-download {
            flex: 1;
            padding: 0.75rem;
            text-align: center;
            border-radius: 6px;
            font-weight: 600;
            text-decoration: none;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.5rem;
        }
        .btn-docx { background: #1d4ed8; color: white; }
        .btn-pdf { background: #dc2626; color: white; }
        .btn-docx:hover { background: #1e40af; }
        .btn-pdf:hover { background: #b91c1c; }

        .preview-area {
            margin-top: 1.25rem;
            background: white;
            border: 1px solid var(--gray-200);
            border-radius: 6px;
            padding: 1rem;
            font-size: 0.85rem;
            max-height: 350px;
            overflow-y: auto;
            white-space: pre-wrap;
            line-height: 1.6;
        }

        footer {
            text-align: center;
            padding: 1.5rem;
            font-size: 0.8rem;
            color: #64748b;
            border-top: 1px solid var(--gray-200);
            background: white;
        }
    </style>
</head>
<body>
    <header>
        <div class="brand">
            <div class="brand-logo">BIDV</div>
            <div>
                <div class="brand-title">Hệ thống Tổng hợp Biên bản Cuộc họp & TBKL</div>
                <div class="brand-subtitle">Chuẩn thể thức hành chính BIDV | Powered by Gemini Argon & Stealth Bridge</div>
            </div>
        </div>
        <div class="bridge-status">
            <div class="status-dot" id="status-dot"></div>
            <span id="status-text">Đang kiểm tra GSB...</span>
        </div>
    </header>

    <main>
        <div class="card">
            <div class="card-header">
                <div class="card-title">
                    <span>📄</span> Thiết lập và Tải lên Tài liệu Phiên họp
                </div>
            </div>

            <form id="minutes-form">
                <div class="form-grid">
                    <div>
                        <label for="doc_type">Loại văn bản hành chính</label>
                        <select id="doc_type" name="doc_type">
                            <option value="BB" selected>Biên bản cuộc họp (BB-BIDV)</option>
                            <option value="TBKL">Thông báo kết luận (TB-BIDV)</option>
                        </select>
                    </div>
                    <div>
                        <label for="model">Mô hình AI xử lý</label>
                        <select id="model" name="model">
                            <option value="gemini-argon" selected>Gemini Argon (Khuyên dùng - Độ sâu chỉ đạo cao nhất)</option>
                            <option value="gemini-pro">Gemini Pro</option>
                            <option value="gemini-flash">Gemini Flash</option>
                            <option value="gemini-flash-lite">Gemini Flash Lite (Nhanh)</option>
                        </select>
                    </div>

                    <div class="full-width">
                        <label for="meeting_title">Tiêu đề phiên họp (Trích yếu / Subtitle)</label>
                        <input type="text" id="meeting_title" name="meeting_title" placeholder="VD: Phiên họp Cấp cao Song phương BIDV - Tencent về Hợp tác Chiến lược..." value="Phiên họp Cấp cao Song phương giữa Ban Lãnh đạo Ngân hàng TMCP Đầu tư và&#10;Phát triển Việt Nam (BIDV) và Ban Lãnh đạo Tập đoàn Tencent (Trung Quốc)&#10;về Hợp tác Chiến lược Toàn diện trong lĩnh vực Chuyển đổi số, AI, Cloud và Thanh toán">
                    </div>

                    <div>
                        <label for="doc_number">Số hiệu văn bản</label>
                        <input type="text" id="doc_number" name="doc_number" value="2710">
                    </div>
                    <div>
                        <label for="signer_name">Họ tên người ký</label>
                        <input type="text" id="signer_name" name="signer_name" value="HOÀNG VIỆT HÙNG">
                    </div>

                    <div class="full-width">
                        <label for="signer_title">Chức danh người ký</label>
                        <input type="text" id="signer_title" name="signer_title" value="ĐẠI DIỆN ĐỒNG CHỦ TRÌ PHIÊN HỌP&#10;TỔNG GIÁM ĐỐC BIDV">
                    </div>

                    <div class="full-width">
                        <label>Tệp tin bản ghi chép cuộc họp (.docx, .txt, .md)</label>
                        <div class="dropzone" id="dropzone">
                            <input type="file" id="file" name="file" accept=".docx,.txt,.md" style="display: none;">
                            <div class="dropzone-icon">📁</div>
                            <div class="dropzone-text">Kéo thả tệp tin vào đây hoặc click để chọn tệp</div>
                            <div class="dropzone-sub">Hỗ trợ file Word (.docx) hoặc file văn bản ghi chép chi tiết</div>
                            <div class="file-chosen" id="file-chosen"></div>
                        </div>
                    </div>

                    <div class="full-width">
                        <label for="raw_text">Hoặc dán trực tiếp nội dung ghi chép phiên họp vào đây (tùy chọn)</label>
                        <textarea id="raw_text" name="raw_text" rows="4" placeholder="Nếu không tải tệp, có thể dán toàn bộ biên bản ghi âm / transcript vào ô này..."></textarea>
                    </div>
                </div>

                <button type="submit" class="btn-submit" id="btn-submit">
                    <span>⚡ Bắt đầu Tổng hợp & Xuất Biên bản BIDV</span>
                </button>
            </form>

            <div class="status-box" id="status-box">
                <span class="status-spinner"></span>
                <span id="status-desc">Đang gửi yêu cầu phân tích tới Gemini Argon thông qua GSB...</span>
            </div>

            <div class="result-box" id="result-box">
                <div class="result-title">
                    <span>✅</span> Biên bản đã được khởi tạo thành công theo đúng chuẩn BIDV!
                </div>
                <div style="font-size: 0.9rem; color: #166534;" id="result-meta">
                    Văn bản đã bóc tách toàn bộ mã định dạng, căn lề Letter 0.75"-1.0", định dạng font Times New Roman 13pt và tạo cấu trúc 2.1 / 2.2 chuẩn thể thức.
                </div>
                <div class="btn-group">
                    <a href="#" class="btn-download btn-docx" id="download-docx" target="_blank">
                        <span>📥</span> Tải văn bản Word (.docx)
                    </a>
                    <a href="#" class="btn-download btn-pdf" id="download-pdf" target="_blank">
                        <span>📄</span> Tải bản PDF (.pdf)
                    </a>
                </div>
                <div class="preview-area" id="preview-text"></div>
            </div>
        </div>
    </main>

    <footer>
        Ngân hàng TMCP Đầu tư và Phát triển Việt Nam (BIDV) — Ban Công nghệ & Dữ liệu | Gemini Stealth Bridge
    </footer>

    <script>
        // Check health
        async function checkHealth() {
            try {
                const res = await fetch('/api/health');
                const data = await res.json();
                const dot = document.getElementById('status-dot');
                const text = document.getElementById('status-text');
                if (data.gsb_connected) {
                    dot.style.background = '#22c55e';
                    dot.style.boxShadow = '0 0 8px #22c55e';
                    text.innerText = 'GSB Ready (' + data.default_model + ')';
                } else {
                    dot.style.background = '#ef4444';
                    dot.style.boxShadow = '0 0 8px #ef4444';
                    text.innerText = 'GSB ' + data.gsb_status;
                }
            } catch (e) {
                document.getElementById('status-dot').style.background = '#ef4444';
                document.getElementById('status-text').innerText = 'Không kết nối được dịch vụ';
            }
        }
        checkHealth();

        // Dropzone
        const dropzone = document.getElementById('dropzone');
        const fileInput = document.getElementById('file');
        const fileChosen = document.getElementById('file-chosen');

        dropzone.addEventListener('click', () => fileInput.click());
        dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
        dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
        dropzone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropzone.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) {
                fileInput.files = e.dataTransfer.files;
                showFileName(e.dataTransfer.files[0].name);
            }
        });
        fileInput.addEventListener('change', () => {
            if (fileInput.files.length > 0) {
                showFileName(fileInput.files[0].name);
            }
        });

        function showFileName(name) {
            fileChosen.style.display = 'inline-block';
            fileChosen.innerText = '✓ Đã chọn: ' + name;
        }

        // Form submit
        const form = document.getElementById('minutes-form');
        const btnSubmit = document.getElementById('btn-submit');
        const statusBox = document.getElementById('status-box');
        const statusDesc = document.getElementById('status-desc');
        const resultBox = document.getElementById('result-box');
        const downloadDocx = document.getElementById('download-docx');
        const downloadPdf = document.getElementById('download-pdf');
        const previewText = document.getElementById('preview-text');

        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            const formData = new FormData(form);

            // Validation
            if (!fileInput.files.length && !document.getElementById('raw_text').value.trim()) {
                alert('Vui lòng chọn một tệp tài liệu ghi chép hoặc dán nội dung vào ô văn bản.');
                return;
            }

            btnSubmit.disabled = true;
            statusBox.style.display = 'block';
            resultBox.style.display = 'none';
            statusDesc.innerText = 'Đang phân tích sâu biên bản và tổng kết theo chuẩn BIDV qua Gemini Argon... (quá trình mất ~30-60 giây)';

            try {
                const response = await fetch('/api/generate', {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.detail || 'Có lỗi xảy ra khi tạo văn bản');
                }

                const data = await response.json();
                statusBox.style.display = 'none';
                resultBox.style.display = 'block';

                downloadDocx.href = data.docx_url;
                if (data.pdf_url) {
                    downloadPdf.href = data.pdf_url;
                    downloadPdf.style.display = 'flex';
                } else {
                    downloadPdf.style.display = 'none';
                }

                previewText.innerText = data.preview_text;
                resultBox.scrollIntoView({ behavior: 'smooth' });

            } catch (err) {
                alert('Lỗi: ' + err.message);
                statusBox.style.display = 'none';
            } finally {
                btnSubmit.disabled = false;
            }
        });
    </script>
</body>
</html>
"""
