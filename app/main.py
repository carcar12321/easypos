from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re
from uuid import uuid4

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates

from app.config import AppConfig, load_config
from app.generator import GenerationArtifact, generate_voucher, save_upload_to_tempfile
from app.parsers import ParsingError


BASE_DIR = Path(__file__).resolve().parent.parent
CONFIG_PATH = BASE_DIR / "config" / "voucher_config.json"
DEFAULT_TEMPLATE_PATH = BASE_DIR / "자동전표 양식.xlsx"
GENERATED_DIR = BASE_DIR / ".generated"
GENERATED_DIR.mkdir(exist_ok=True)

app = FastAPI(title="POS Voucher Generator")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
config: AppConfig = load_config(CONFIG_PATH)


@dataclass(frozen=True)
class ResultRecord:
    artifact: GenerationArtifact


RESULTS: dict[str, ResultRecord] = {}


def _normalize_account_date_input(raw_value: str) -> str | None:
    value = raw_value.strip()
    if not value:
        return None

    digits = re.sub(r"[^0-9]", "", value)
    if len(digits) == 6:
        digits = f"20{digits}"
    if len(digits) != 8:
        raise ParsingError("회계일자는 YYYYMMDD 또는 YYMMDD 형태로 입력해 주세요.")
    try:
        datetime.strptime(digits, "%Y%m%d")
    except ValueError as error:
        raise ParsingError("유효하지 않은 회계일자입니다. 예: 20260329 또는 260329") from error
    return digits


@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "businesses": sorted(config.businesses.values(), key=lambda item: item.display_name),
            "result": None,
            "error": None,
            "account_date_input": "",
        },
    )


@app.post("/generate", response_class=HTMLResponse)
async def generate(
    request: Request,
    business_key: str = Form(...),
    account_date_input: str = Form(default=""),
    card_sales_file: UploadFile = File(...),
    daily_sales_file: UploadFile = File(...),
    template_file: UploadFile | None = File(default=None),
) -> HTMLResponse:
    business = config.businesses.get(business_key)
    if not business:
        raise HTTPException(status_code=400, detail="알 수 없는 사업장입니다.")

    if not card_sales_file.filename or not daily_sales_file.filename:
        raise HTTPException(status_code=400, detail="업로드 파일명이 비어 있습니다.")

    try:
        manual_account_date = _normalize_account_date_input(account_date_input)
        card_path = save_upload_to_tempfile(card_sales_file.filename, await card_sales_file.read())
        daily_path = save_upload_to_tempfile(daily_sales_file.filename, await daily_sales_file.read())

        if template_file and template_file.filename:
            template_path = save_upload_to_tempfile(template_file.filename, await template_file.read())
        else:
            if not DEFAULT_TEMPLATE_PATH.exists():
                raise ParsingError(
                    "기본 템플릿 파일(자동전표 양식.xlsx)을 찾지 못했습니다. 전표 템플릿 파일을 업로드해 주세요."
                )
            template_path = DEFAULT_TEMPLATE_PATH

        artifact = generate_voucher(
            business=business,
            card_file_path=card_path,
            daily_file_path=daily_path,
            template_file_path=template_path,
            config=config,
            output_dir=GENERATED_DIR,
            manual_account_date=manual_account_date,
        )
    except ParsingError as error:
        return templates.TemplateResponse(
            request,
            "index.html",
            {
                "businesses": sorted(config.businesses.values(), key=lambda item: item.display_name),
                "result": None,
                "error": str(error),
                "account_date_input": account_date_input,
            },
            status_code=400,
        )

    result_id = uuid4().hex
    RESULTS[result_id] = ResultRecord(artifact=artifact)
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "businesses": sorted(config.businesses.values(), key=lambda item: item.display_name),
            "error": None,
            "account_date_input": account_date_input,
            "result": {
                "id": result_id,
                "output_filename": artifact.output_filename,
                "generated_store_names": artifact.generated_store_names,
                "notes": artifact.notes,
            },
        },
    )


@app.get("/download/{result_id}")
async def download(result_id: str) -> FileResponse:
    record = RESULTS.get(result_id)
    if not record:
        raise HTTPException(status_code=404, detail="결과 파일을 찾지 못했습니다.")
    artifact = record.artifact
    return FileResponse(
        path=artifact.output_path,
        filename=artifact.output_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}
