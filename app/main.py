from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re
from uuid import uuid4

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates

from app.config import AppConfig, load_config, load_store_name_mapping
from app.generator import (
    GenerationArtifact,
    SalesFormatInputArtifact,
    SettlementVerificationArtifact,
    generate_sales_input_from_sales_format,
    generate_voucher,
    save_upload_to_tempfile,
    verify_settlement_against_voucher,
)
from app.parsers import ParsingError


BASE_DIR = Path(__file__).resolve().parent.parent
CONFIG_PATH = BASE_DIR / "config" / "voucher_config.json"
STORE_NAME_MAPPING_PATH = BASE_DIR / "config" / "store_name_mapping.json"
DEFAULT_TEMPLATE_PATH = BASE_DIR / "자동전표 양식.xlsx"
GENERATED_DIR = BASE_DIR / ".generated"
GENERATED_DIR.mkdir(exist_ok=True)

app = FastAPI(title="POS Voucher Generator")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
config: AppConfig = load_config(CONFIG_PATH)
store_name_mapping = load_store_name_mapping(STORE_NAME_MAPPING_PATH)


@dataclass(frozen=True)
class ResultRecord:
    artifact: GenerationArtifact


RESULTS: dict[str, ResultRecord] = {}


@dataclass(frozen=True)
class SalesInputResultRecord:
    artifact: SalesFormatInputArtifact


SALES_INPUT_RESULTS: dict[str, SalesInputResultRecord] = {}


def _base_context() -> dict[str, object]:
    return {
        "businesses": sorted(config.businesses.values(), key=lambda item: item.display_name),
        "result": None,
        "error": None,
        "account_date_input": "",
        "sales_input_result": None,
        "sales_input_error": None,
        "sales_date_from_input": "",
        "sales_date_to_input": "",
        "verify_result": None,
        "verify_error": None,
        "verify_account_date_input": "",
    }


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
    return templates.TemplateResponse(request, "index.html", _base_context())


@app.post("/generate", response_class=HTMLResponse)
async def generate(
    request: Request,
    business_key: str = Form(...),
    account_date_input: str = Form(default=""),
    card_sales_file: UploadFile = File(...),
    daily_sales_file: UploadFile = File(...),
    settlement_order_file: UploadFile | None = File(default=None),
    template_file: UploadFile | None = File(default=None),
) -> HTMLResponse:
    business = config.businesses.get(business_key)
    if not business:
        raise HTTPException(status_code=400, detail="알 수 없는 사업장입니다.")

    if not card_sales_file.filename or not daily_sales_file.filename:
        raise HTTPException(status_code=400, detail="업로드 파일명이 비어 있습니다.")

    try:
        settlement_path: Path | None = None
        manual_account_date = _normalize_account_date_input(account_date_input)
        card_path = save_upload_to_tempfile(card_sales_file.filename, await card_sales_file.read())
        daily_path = save_upload_to_tempfile(daily_sales_file.filename, await daily_sales_file.read())
        if settlement_order_file and settlement_order_file.filename:
            settlement_path = save_upload_to_tempfile(
                settlement_order_file.filename,
                await settlement_order_file.read(),
            )

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
            settlement_order_file_path=settlement_path,
            settlement_store_aliases=store_name_mapping.get(business_key, {}),
        )
    except ParsingError as error:
        context = _base_context()
        context["error"] = str(error)
        context["account_date_input"] = account_date_input
        return templates.TemplateResponse(
            request,
            "index.html",
            context,
            status_code=400,
        )

    result_id = uuid4().hex
    RESULTS[result_id] = ResultRecord(artifact=artifact)
    context = _base_context()
    context["account_date_input"] = account_date_input
    context["result"] = {
        "id": result_id,
        "output_filename": artifact.output_filename,
        "generated_store_names": artifact.generated_store_names,
        "notes": artifact.notes,
    }
    return templates.TemplateResponse(
        request,
        "index.html",
        context,
    )


@app.post("/generate-sales-input", response_class=HTMLResponse)
async def generate_sales_input(
    request: Request,
    business_key: str = Form(...),
    sales_date_from_input: str = Form(default=""),
    sales_date_to_input: str = Form(default=""),
    sales_format_file: UploadFile = File(...),
) -> HTMLResponse:
    business = config.businesses.get(business_key)
    if not business:
        raise HTTPException(status_code=400, detail="알 수 없는 사업장입니다.")
    if not sales_format_file.filename:
        raise HTTPException(status_code=400, detail="업로드 파일명이 비어 있습니다.")

    context = _base_context()
    context["sales_date_from_input"] = sales_date_from_input
    context["sales_date_to_input"] = sales_date_to_input

    try:
        date_from = _normalize_account_date_input(sales_date_from_input) if sales_date_from_input.strip() else None
        date_to = _normalize_account_date_input(sales_date_to_input) if sales_date_to_input.strip() else None
        if date_from and date_to and date_from > date_to:
            raise ParsingError("날짜 범위가 올바르지 않습니다. 시작일은 종료일보다 이전이어야 합니다.")

        sales_format_path = save_upload_to_tempfile(
            sales_format_file.filename,
            await sales_format_file.read(),
        )
        artifact = generate_sales_input_from_sales_format(
            business=business,
            sales_format_file_path=sales_format_path,
            output_dir=GENERATED_DIR,
            date_from=date_from,
            date_to=date_to,
        )
    except ParsingError as error:
        context["sales_input_error"] = str(error)
        return templates.TemplateResponse(request, "index.html", context, status_code=400)

    result_id = uuid4().hex
    SALES_INPUT_RESULTS[result_id] = SalesInputResultRecord(artifact=artifact)
    context["sales_input_result"] = {
        "id": result_id,
        "output_filename": artifact.output_filename,
        "row_count": artifact.row_count,
        "store_count": artifact.store_count,
        "date_min": artifact.date_min,
        "date_max": artifact.date_max,
        "notes": artifact.notes,
    }
    return templates.TemplateResponse(request, "index.html", context)


@app.post("/verify-settlement", response_class=HTMLResponse)
async def verify_settlement(
    request: Request,
    business_key: str = Form(...),
    verify_account_date_input: str = Form(default=""),
    voucher_file: UploadFile = File(...),
    settlement_order_file_verify: UploadFile = File(...),
) -> HTMLResponse:
    business = config.businesses.get(business_key)
    if not business:
        raise HTTPException(status_code=400, detail="알 수 없는 사업장입니다.")
    if not voucher_file.filename or not settlement_order_file_verify.filename:
        raise HTTPException(status_code=400, detail="검증 파일명이 비어 있습니다.")

    context = _base_context()
    context["verify_account_date_input"] = verify_account_date_input
    try:
        manual_account_date = _normalize_account_date_input(verify_account_date_input)
        voucher_path = save_upload_to_tempfile(voucher_file.filename, await voucher_file.read())
        settlement_path = save_upload_to_tempfile(
            settlement_order_file_verify.filename,
            await settlement_order_file_verify.read(),
        )
        settlement_partner = config.payment_by_key.get(config.settlement_partner_payment_key)
        if settlement_partner is None:
            raise ParsingError(
                f"정산주문목록용 결제수단 키({config.settlement_partner_payment_key})를 payment_methods에서 찾지 못했습니다."
            )

        artifact: SettlementVerificationArtifact = verify_settlement_against_voucher(
            business=business,
            voucher_file_path=voucher_path,
            settlement_order_file_path=settlement_path,
            settlement_partner_code=settlement_partner.partner_code,
            settlement_management_data=config.settlement_management_data,
            settlement_store_aliases=store_name_mapping.get(business_key, {}),
            manual_account_date=manual_account_date,
        )
        context["verify_result"] = {
            "account_date": artifact.account_date,
            "settlement_store_count": artifact.settlement_store_count,
            "voucher_store_count": artifact.voucher_store_count,
            "matched_merchant_count": artifact.matched_merchant_count,
            "unmatched_merchants": artifact.unmatched_merchants,
            "differences": artifact.differences,
        }
    except ParsingError as error:
        context["verify_error"] = str(error)
        return templates.TemplateResponse(request, "index.html", context, status_code=400)

    return templates.TemplateResponse(request, "index.html", context)


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


@app.get("/download-sales-input/{result_id}")
async def download_sales_input(result_id: str) -> FileResponse:
    record = SALES_INPUT_RESULTS.get(result_id)
    if not record:
        raise HTTPException(status_code=404, detail="자동 입력 결과 파일을 찾지 못했습니다.")
    artifact = record.artifact
    return FileResponse(
        path=artifact.output_path,
        filename=artifact.output_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}
