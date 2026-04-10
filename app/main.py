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
    SalesTemplateFillArtifact,
    SettlementVerificationArtifact,
    generate_sales_template_auto_input,
    generate_voucher,
    save_upload_to_tempfile,
    verify_settlement_against_voucher,
)
from app.parsers import ParsingError


BASE_DIR = Path(__file__).resolve().parent.parent
CONFIG_PATH = BASE_DIR / "config" / "voucher_config.json"
STORE_NAME_MAPPING_PATH = BASE_DIR / "config" / "store_name_mapping.json"
DEFAULT_TEMPLATE_PATH = BASE_DIR / "자동전표 양식.xlsx"
DEFAULT_VALIDATION_TEMPLATE_PATH = BASE_DIR / "검증시트.xlsx"
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
    artifact: SalesTemplateFillArtifact


SALES_INPUT_RESULTS: dict[str, SalesInputResultRecord] = {}


def _default_business_key() -> str:
    if "seoul_station" in config.businesses:
        return "seoul_station"
    if not config.businesses:
        return ""
    return sorted(config.businesses.keys())[0]


def _base_context() -> dict[str, object]:
    return {
        "businesses": sorted(config.businesses.values(), key=lambda item: item.display_name),
        "selected_business_key": _default_business_key(),
        "result": None,
        "error": None,
        "account_date_input": "",
        "sales_input_result": None,
        "sales_input_error": None,
        "sales_account_date_input": "",
        "verify_result": None,
        "verify_error": None,
        "verify_account_date_input": "",
    }


def _normalize_account_date_input(raw_value: str) -> str:
    value = raw_value.strip()
    if not value:
        raise ParsingError("회계일자는 YYYYMMDD 형식으로 반드시 입력해 주세요.")

    if not re.fullmatch(r"\d{8}", value):
        raise ParsingError("회계일자는 YYYYMMDD 8자리 숫자로 입력해 주세요. 예: 20260329")
    try:
        datetime.strptime(value, "%Y%m%d")
    except ValueError as error:
        raise ParsingError("유효하지 않은 회계일자입니다. 예: 20260329") from error
    return value


def _is_openxml_xlsx(path: Path) -> bool:
    if not path.exists() or path.suffix.lower() != ".xlsx":
        return False
    try:
        with path.open("rb") as handle:
            return handle.read(4) == b"PK\x03\x04"
    except OSError:
        return False


def _find_default_sales_template_path() -> Path | None:
    candidates = sorted(BASE_DIR.glob("*임대을매출*POS매출*.xlsx"))
    if not candidates:
        return None
    return candidates[-1]


def _find_default_validation_template_path() -> Path | None:
    candidates = [
        BASE_DIR / "validation_template.xlsx",
        DEFAULT_VALIDATION_TEMPLATE_PATH,
        BASE_DIR / "config" / "검증시트.xlsx",
        BASE_DIR / "config" / "validation_template.xlsx",
    ]
    for candidate in candidates:
        if _is_openxml_xlsx(candidate):
            return candidate
    return None


@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(request, "index.html", _base_context())


@app.post("/generate", response_class=HTMLResponse)
async def generate(
    request: Request,
    business_key: str = Form(...),
    account_date_input: str = Form(...),
    card_sales_file: UploadFile = File(...),
    daily_sales_file: UploadFile = File(...),
    settlement_order_file: UploadFile | None = File(default=None),
    template_file: UploadFile | None = File(default=None),
    sales_template_file: UploadFile | None = File(default=None),
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

        if sales_template_file and sales_template_file.filename:
            sales_template_path = save_upload_to_tempfile(
                sales_template_file.filename,
                await sales_template_file.read(),
            )
        else:
            default_sales_template_path = _find_default_sales_template_path()
            if default_sales_template_path is None:
                raise ParsingError(
                    "기본 매출 양식 템플릿을 찾지 못했습니다. 매출 양식 파일을 직접 업로드해 주세요."
                )
            sales_template_path = default_sales_template_path

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
        sales_artifact = generate_sales_template_auto_input(
            business=business,
            card_file_path=card_path,
            daily_file_path=daily_path,
            sales_template_file_path=sales_template_path,
            config=config,
            output_dir=GENERATED_DIR,
            manual_account_date=manual_account_date,
            validation_template_file_path=_find_default_validation_template_path(),
            settlement_order_file_path=settlement_path,
            settlement_store_aliases=store_name_mapping.get(business_key, {}),
        )
    except ParsingError as error:
        context = _base_context()
        context["selected_business_key"] = business_key
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
    sales_result_id = uuid4().hex
    SALES_INPUT_RESULTS[sales_result_id] = SalesInputResultRecord(artifact=sales_artifact)
    context = _base_context()
    context["selected_business_key"] = business_key
    context["account_date_input"] = account_date_input
    context["result"] = {
        "id": result_id,
        "output_filename": artifact.output_filename,
        "sales_input_id": sales_result_id,
        "sales_output_filename": sales_artifact.output_filename,
        "generated_store_names": artifact.generated_store_names,
        "notes": tuple(artifact.notes) + tuple(sales_artifact.notes),
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
    sales_account_date_input: str = Form(...),
    card_sales_file_sales: UploadFile = File(...),
    daily_sales_file_sales: UploadFile = File(...),
    settlement_order_file_sales: UploadFile | None = File(default=None),
    sales_template_file: UploadFile | None = File(default=None),
) -> HTMLResponse:
    business = config.businesses.get(business_key)
    if not business:
        raise HTTPException(status_code=400, detail="알 수 없는 사업장입니다.")
    if not card_sales_file_sales.filename or not daily_sales_file_sales.filename:
        raise HTTPException(status_code=400, detail="업로드 파일명이 비어 있습니다.")

    context = _base_context()
    context["selected_business_key"] = business_key
    context["sales_account_date_input"] = sales_account_date_input

    try:
        manual_account_date = _normalize_account_date_input(sales_account_date_input)

        card_path = save_upload_to_tempfile(
            card_sales_file_sales.filename,
            await card_sales_file_sales.read(),
        )
        daily_path = save_upload_to_tempfile(
            daily_sales_file_sales.filename,
            await daily_sales_file_sales.read(),
        )
        settlement_path: Path | None = None
        if settlement_order_file_sales and settlement_order_file_sales.filename:
            settlement_path = save_upload_to_tempfile(
                settlement_order_file_sales.filename,
                await settlement_order_file_sales.read(),
            )

        if sales_template_file and sales_template_file.filename:
            sales_template_path = save_upload_to_tempfile(
                sales_template_file.filename,
                await sales_template_file.read(),
            )
        else:
            default_sales_template_path = _find_default_sales_template_path()
            if default_sales_template_path is None:
                raise ParsingError(
                    "기본 매출 양식 템플릿을 찾지 못했습니다. 매출 양식 파일을 직접 업로드해 주세요."
                )
            sales_template_path = default_sales_template_path

        artifact = generate_sales_template_auto_input(
            business=business,
            card_file_path=card_path,
            daily_file_path=daily_path,
            sales_template_file_path=sales_template_path,
            config=config,
            output_dir=GENERATED_DIR,
            manual_account_date=manual_account_date,
            validation_template_file_path=_find_default_validation_template_path(),
            settlement_order_file_path=settlement_path,
            settlement_store_aliases=store_name_mapping.get(business_key, {}),
        )
    except ParsingError as error:
        context["sales_input_error"] = str(error)
        return templates.TemplateResponse(request, "index.html", context, status_code=400)

    result_id = uuid4().hex
    SALES_INPUT_RESULTS[result_id] = SalesInputResultRecord(artifact=artifact)
    context["sales_input_result"] = {
        "id": result_id,
        "output_filename": artifact.output_filename,
        "account_date": artifact.account_date,
        "filled_store_names": artifact.filled_store_names,
        "filled_store_count": len(artifact.filled_store_names),
        "skipped_store_names": artifact.skipped_store_names,
        "notes": artifact.notes,
    }
    return templates.TemplateResponse(request, "index.html", context)


@app.post("/verify-settlement", response_class=HTMLResponse)
async def verify_settlement(
    request: Request,
    business_key: str = Form(...),
    verify_account_date_input: str = Form(...),
    voucher_file: UploadFile = File(...),
    settlement_order_file_verify: UploadFile = File(...),
) -> HTMLResponse:
    business = config.businesses.get(business_key)
    if not business:
        raise HTTPException(status_code=400, detail="알 수 없는 사업장입니다.")
    if not voucher_file.filename or not settlement_order_file_verify.filename:
        raise HTTPException(status_code=400, detail="검증 파일명이 비어 있습니다.")

    context = _base_context()
    context["selected_business_key"] = business_key
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
