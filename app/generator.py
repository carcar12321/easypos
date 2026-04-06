from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any

from openpyxl import Workbook, load_workbook

from app.config import AppConfig, BusinessConfig, PaymentMethodConfig, StoreMapping
from app.parsers import (
    DailySalesRow,
    ParsedCardSales,
    ParsedDailySales,
    ParsedSalesFormatInput,
    ParsedSettlementOrders,
    ParsedVoucherSettlement,
    ParsingError,
    parse_card_sales,
    parse_daily_sales,
    parse_generated_voucher_settlement,
    parse_sales_format_input,
    parse_settlement_orders,
)


@dataclass(frozen=True)
class VoucherLine:
    docu_no: str
    doline_sq: int
    row_line_sq: int
    actg_dt: str
    wrt_dt: str
    consul_dc: str | None
    drcrfg_cd: str
    acct_cd: str
    docu_amt: float
    docu_amt2: str
    note_dc: str
    partner_cd: str
    partner_nm: str
    cc_cd: str | None
    management_data: str | None
    row_no: str
    company_cd: str
    pc_cd: str
    gaap_cd: str
    docu_cd: str
    wrt_dept_cd: str
    wrt_emp_no: str
    insr_fg_cd: str

    def as_row(self) -> list[Any]:
        return [
            self.docu_no,
            self.doline_sq,
            self.row_line_sq,
            self.actg_dt,
            self.wrt_dt,
            self.consul_dc,
            self.drcrfg_cd,
            self.acct_cd,
            self.docu_amt,
            self.docu_amt2,
            self.note_dc,
            self.partner_cd,
            self.partner_nm,
            self.cc_cd,
            self.management_data,
            self.row_no,
            self.company_cd,
            self.pc_cd,
            self.gaap_cd,
            self.docu_cd,
            self.wrt_dept_cd,
            self.wrt_emp_no,
            self.insr_fg_cd,
        ]


@dataclass(frozen=True)
class GenerationArtifact:
    output_path: Path
    output_filename: str
    generated_store_names: tuple[str, ...]
    notes: tuple[str, ...]


@dataclass(frozen=True)
class SettlementVerificationArtifact:
    account_date: str
    settlement_store_count: int
    voucher_store_count: int
    matched_merchant_count: int
    unmatched_merchants: tuple[str, ...]
    differences: tuple[str, ...]


@dataclass(frozen=True)
class SalesFormatInputArtifact:
    output_path: Path
    output_filename: str
    row_count: int
    store_count: int
    date_min: str | None
    date_max: str | None
    notes: tuple[str, ...]


def save_upload_to_tempfile(filename: str, content: bytes) -> Path:
    suffix = Path(filename).suffix or ".xlsx"
    with NamedTemporaryFile(delete=False, suffix=suffix) as handle:
        handle.write(content)
        return Path(handle.name)


def _short_date(account_date: str) -> str:
    return account_date[2:]


def _format_amount_text(amount: float) -> str:
    rounded = round(amount)
    if abs(rounded - amount) < 1e-9:
        return str(int(rounded))
    return str(amount)


def _build_consul_dc(business: BusinessConfig, account_date: str) -> str:
    return business.consul_dc_format.format(
        date_short=_short_date(account_date),
        date_full=account_date,
        business=business.display_name,
    )


def _build_note_dc(business: BusinessConfig, account_date: str, output_name: str) -> str:
    return business.note_dc_format.format(
        date_short=_short_date(account_date),
        date_full=account_date,
        business=business.display_name,
        store=output_name,
    )


def _line_common(business: BusinessConfig, account_date: str) -> dict[str, str]:
    return {
        "docu_no": f"{business.doc_prefix}{account_date}",
        "actg_dt": account_date,
        "wrt_dt": account_date,
        "row_no": "1",
        "company_cd": business.company_code,
        "pc_cd": business.pc_code,
        "gaap_cd": business.gaap_code,
        "docu_cd": business.docu_code,
        "wrt_dept_cd": business.wrt_dept_code,
        "wrt_emp_no": business.wrt_emp_no,
        "insr_fg_cd": business.insr_fg_cd,
    }


def _payment_amounts_for_store(
    store: StoreMapping,
    card_sales: ParsedCardSales,
    daily_sales: ParsedDailySales,
    payment_methods: tuple[PaymentMethodConfig, ...],
) -> dict[str, float]:
    return dict(card_sales.by_store.get(store.source_name, {}))


def _build_store_lines(
    business: BusinessConfig,
    store: StoreMapping,
    daily_row: DailySalesRow,
    payment_amounts: dict[str, float],
    payment_methods: tuple[PaymentMethodConfig, ...],
    start_seq: int,
    account_date: str,
    include_consul_dc: bool,
    settlement_amount: float = 0.0,
    settlement_payment: PaymentMethodConfig | None = None,
    settlement_management_data: str | None = None,
) -> list[VoucherLine]:
    common = _line_common(business, account_date)
    note_dc = _build_note_dc(business, account_date, store.output_name)
    discount_amount = abs(daily_row.discount_amount or 0.0)
    receivable_total = sum(payment_amounts.values()) + settlement_amount
    debit_total = daily_row.cash_sales + receivable_total
    # 새 규칙:
    # - 할인(401511)은 할인액을 음수로
    # - 현금(103300) + 카드승인(102700)을 먼저 채운 뒤
    # - 남는 금액으로 401510을 계산해 차대합을 맞춘다.
    gross_sales_amount = debit_total + discount_amount
    lines: list[VoucherLine] = []
    seq = start_seq

    def add_line(
        *,
        drcrfg_cd: str,
        acct_cd: str,
        amount: float,
        partner_cd: str,
        partner_nm: str,
        cc_cd: str | None,
        management_data: str | None,
        consul_dc: str | None = None,
    ) -> None:
        nonlocal seq
        lines.append(
            VoucherLine(
                doline_sq=seq,
                row_line_sq=seq,
                consul_dc=consul_dc,
                drcrfg_cd=drcrfg_cd,
                acct_cd=acct_cd,
                docu_amt=amount,
                docu_amt2=_format_amount_text(amount),
                note_dc=note_dc,
                partner_cd=partner_cd,
                partner_nm=partner_nm,
                cc_cd=cc_cd,
                management_data=management_data,
                **common,
            )
        )
        seq += 1

    add_line(
        drcrfg_cd="2",
        acct_cd=business.account_codes["gross_sales"],
        amount=gross_sales_amount,
        partner_cd=store.partner_code,
        partner_nm=store.partner_name,
        cc_cd="1300",
        management_data=None,
        consul_dc=_build_consul_dc(business, account_date) if include_consul_dc else None,
    )

    if discount_amount:
        add_line(
            drcrfg_cd="2",
            acct_cd=business.account_codes["discount"],
            amount=(-1) * discount_amount,
            partner_cd=store.partner_code,
            partner_nm=store.partner_name,
            cc_cd="1300",
            management_data=None,
        )

    if daily_row.cash_sales:
        add_line(
            drcrfg_cd="1",
            acct_cd=business.account_codes["cash"],
            amount=daily_row.cash_sales,
            partner_cd=store.partner_code,
            partner_nm=store.partner_name,
            cc_cd=None,
            management_data=None,
        )

    for payment in payment_methods:
        amount = payment_amounts.get(payment.key, 0.0)
        if not amount:
            continue
        add_line(
            drcrfg_cd="1",
            acct_cd=business.account_codes["receivable"],
            amount=amount,
            partner_cd=payment.partner_code,
            partner_nm=payment.partner_name,
            cc_cd=None,
            management_data=payment.management_data,
        )

    if settlement_amount and settlement_payment is not None:
        add_line(
            drcrfg_cd="1",
            acct_cd=business.account_codes["receivable"],
            amount=settlement_amount,
            partner_cd=settlement_payment.partner_code,
            partner_nm=settlement_payment.partner_name,
            cc_cd=None,
            management_data=settlement_management_data or settlement_payment.management_data,
        )

    return lines


def _write_output(template_path: Path, output_path: Path, lines: list[VoucherLine]) -> None:
    workbook = load_workbook(template_path)
    worksheet = workbook.worksheets[0]

    if worksheet.max_row >= 4:
        worksheet.delete_rows(4, worksheet.max_row - 3)

    for row_index, line in enumerate(lines, start=4):
        for column_index, value in enumerate(line.as_row(), start=1):
            worksheet.cell(row_index, column_index).value = value

    workbook.save(output_path)


def _write_sales_input_output(output_path: Path, parsed: ParsedSalesFormatInput) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Sheet1"
    worksheet.append(
        [
            "일자",
            "가맹점명",
            "가맹점코드",
            "카드사명",
            "카드승인금액",
            "카드구분(승인/취소)",
            "현금매출액",
            "현금구분(승인/취소)",
            "할인액",
        ]
    )

    for row in parsed.rows:
        worksheet.append(
            [
                row.account_date,
                row.source_name,
                row.partner_code,
                "매출양식카드합계",
                row.card_amount,
                "승인",
                row.cash_amount,
                "승인",
                row.discount_amount,
            ]
        )

    workbook.save(output_path)


def _build_sales_input_filename(
    business: BusinessConfig,
    date_min: str | None,
    date_max: str | None,
) -> str:
    if date_min and date_max:
        if date_min == date_max:
            date_part = date_min
        else:
            date_part = f"{date_min}_{date_max}"
    else:
        date_part = "date_unknown"
    return f"{business.display_name}업체제출자동입력_{date_part}.xlsx"


def generate_voucher(
    *,
    business: BusinessConfig,
    card_file_path: Path,
    daily_file_path: Path,
    template_file_path: Path,
    config: AppConfig,
    output_dir: Path,
    manual_account_date: str | None = None,
    settlement_order_file_path: Path | None = None,
    settlement_store_aliases: dict[str, tuple[str, ...]] | None = None,
) -> GenerationArtifact:
    notes: list[str] = []
    allow_missing_date = manual_account_date is not None
    parsed_card = parse_card_sales(card_file_path, config, allow_missing_date=allow_missing_date)
    parsed_daily = parse_daily_sales(daily_file_path, allow_missing_date=allow_missing_date)

    if manual_account_date is not None:
        account_date = manual_account_date
        if parsed_card.account_date and parsed_card.account_date != manual_account_date:
            notes.append(
                f"카드 파일 회계일({parsed_card.account_date}) 대신 입력값({manual_account_date})을 사용했습니다."
            )
        if parsed_daily.account_date and parsed_daily.account_date != manual_account_date:
            notes.append(
                f"당일 파일 회계일({parsed_daily.account_date}) 대신 입력값({manual_account_date})을 사용했습니다."
            )
    else:
        if parsed_card.account_date != parsed_daily.account_date:
            raise ParsingError(
                "카드매출/당일매출 파일의 회계일이 다릅니다. "
                f"카드={parsed_card.account_date}, 당일={parsed_daily.account_date}"
            )

        if parsed_card.account_date is None:
            raise ParsingError("회계일자를 확인할 수 없습니다. 회계일자를 직접 입력해 주세요.")
        account_date = parsed_card.account_date

    parsed_settlement: ParsedSettlementOrders | None = None
    settlement_payment = config.payment_by_key.get(config.settlement_partner_payment_key)
    if settlement_order_file_path is not None:
        if settlement_payment is None:
            raise ParsingError(
                f"정산주문목록용 결제수단 키({config.settlement_partner_payment_key})를 payment_methods에서 찾지 못했습니다."
            )
        parsed_settlement = parse_settlement_orders(
            settlement_order_file_path,
            source_names=business.source_store_names,
            store_aliases=settlement_store_aliases,
            expected_account_date=account_date,
            allow_missing_date=True,
        )
        notes.append(
            "정산주문목록 매칭: "
            f"{len(parsed_settlement.matched_merchants)}개 가맹점, "
            f"{len(parsed_settlement.unmatched_merchants)}개 미매칭"
        )
        if parsed_settlement.unmatched_merchants:
            notes.append("정산주문목록 미매칭 가맹점: " + ", ".join(parsed_settlement.unmatched_merchants))

    generated_store_names: list[str] = []
    lines: list[VoucherLine] = []
    next_seq = 1
    first_line = True

    for store in business.active_stores:
        daily_row = parsed_daily.by_store.get(store.source_name)
        if not daily_row:
            continue

        amounts = _payment_amounts_for_store(
            store=store,
            card_sales=parsed_card,
            daily_sales=parsed_daily,
            payment_methods=config.payment_methods_in_order,
        )
        settlement_amount = parsed_settlement.by_store.get(store.source_name, 0.0) if parsed_settlement else 0.0

        has_amount = any(
            [
                daily_row.gross_sales,
                daily_row.discount_amount,
                daily_row.cash_sales,
                daily_row.electronic_money_sales,
                sum(amounts.values()),
                settlement_amount,
            ]
        )
        if not has_amount or daily_row.gross_sales <= 0:
            continue

        store_lines = _build_store_lines(
            business=business,
            store=store,
            daily_row=daily_row,
            payment_amounts=amounts,
            payment_methods=config.payment_methods_in_order,
            start_seq=next_seq,
            account_date=account_date,
            include_consul_dc=first_line,
            settlement_amount=settlement_amount,
            settlement_payment=settlement_payment,
            settlement_management_data=config.settlement_management_data,
        )
        lines.extend(store_lines)
        next_seq += len(store_lines)
        first_line = False
        generated_store_names.append(store.output_name)

        if store.manual_review_reason:
            notes.append(f"{store.output_name}: {store.manual_review_reason}")

    input_store_names = {
        name
        for name, row in parsed_daily.by_store.items()
        if row.gross_sales or row.discount_amount or row.cash_sales or row.electronic_money_sales
    }
    active_source_names = business.source_store_names
    skipped = sorted(input_store_names - active_source_names)
    if skipped:
        notes.append("자동 생성 제외 매장(입력에는 존재): " + ", ".join(skipped))

    output_filename = f"{business.display_name}POS매장임대을매출_{account_date}.xlsx"
    output_path = output_dir / output_filename
    _write_output(template_file_path, output_path, lines)

    return GenerationArtifact(
        output_path=output_path,
        output_filename=output_filename,
        generated_store_names=tuple(generated_store_names),
        notes=tuple(notes),
    )


def generate_sales_input_from_sales_format(
    *,
    business: BusinessConfig,
    sales_format_file_path: Path,
    output_dir: Path,
    date_from: str | None = None,
    date_to: str | None = None,
) -> SalesFormatInputArtifact:
    output_dir.mkdir(parents=True, exist_ok=True)
    parsed = parse_sales_format_input(
        sales_format_file_path,
        business=business,
        date_from=date_from,
        date_to=date_to,
    )
    if not parsed.rows:
        date_range_text = ""
        if date_from or date_to:
            date_range_text = f" (요청 범위: {date_from or '시작제한없음'} ~ {date_to or '종료제한없음'})"
        raise ParsingError(f"조건에 맞는 자동 입력 데이터가 없습니다{date_range_text}.")

    output_filename = _build_sales_input_filename(
        business=business,
        date_min=parsed.date_min,
        date_max=parsed.date_max,
    )
    output_path = output_dir / output_filename
    _write_sales_input_output(output_path, parsed)

    notes = list(parsed.notes)
    if parsed.skipped_store_names:
        notes.append("자동 입력 제외 매장(시트 미존재): " + ", ".join(parsed.skipped_store_names))
    notes.append("입력 대상 매장 외 시트는 자동 무시했습니다.")

    return SalesFormatInputArtifact(
        output_path=output_path,
        output_filename=output_filename,
        row_count=len(parsed.rows),
        store_count=len({row.source_name for row in parsed.rows}),
        date_min=parsed.date_min,
        date_max=parsed.date_max,
        notes=tuple(notes),
    )


def verify_settlement_against_voucher(
    *,
    business: BusinessConfig,
    voucher_file_path: Path,
    settlement_order_file_path: Path,
    settlement_partner_code: str,
    settlement_management_data: str | None = None,
    settlement_store_aliases: dict[str, tuple[str, ...]] | None = None,
    manual_account_date: str | None = None,
) -> SettlementVerificationArtifact:
    parsed_voucher: ParsedVoucherSettlement = parse_generated_voucher_settlement(
        voucher_file_path,
        settlement_partner_code=settlement_partner_code,
        settlement_management_data=settlement_management_data,
    )
    expected_account_date = manual_account_date or parsed_voucher.account_date
    if expected_account_date is None:
        raise ParsingError("검증 기준 회계일자를 찾지 못했습니다. 회계일자를 직접 입력해 주세요.")

    parsed_settlement: ParsedSettlementOrders = parse_settlement_orders(
        settlement_order_file_path,
        source_names=business.source_store_names,
        store_aliases=settlement_store_aliases,
        expected_account_date=expected_account_date,
        allow_missing_date=True,
    )

    source_to_output = {store.source_name: store.output_name for store in business.active_stores}
    expected_by_output: dict[str, float] = {}
    for source_name, amount in parsed_settlement.by_store.items():
        output_name = source_to_output.get(source_name)
        if output_name is None:
            continue
        expected_by_output[output_name] = expected_by_output.get(output_name, 0.0) + amount

    actual_by_output = parsed_voucher.by_output_store
    all_output_names = sorted(set(expected_by_output) | set(actual_by_output))
    differences: list[str] = []
    for output_name in all_output_names:
        expected_amount = expected_by_output.get(output_name, 0.0)
        actual_amount = actual_by_output.get(output_name, 0.0)
        diff = actual_amount - expected_amount
        if abs(diff) > 0.1:
            differences.append(
                f"{output_name}: 전표={_format_amount_text(actual_amount)}, "
                f"정산주문목록={_format_amount_text(expected_amount)}, 차이={_format_amount_text(diff)}"
            )

    return SettlementVerificationArtifact(
        account_date=expected_account_date,
        settlement_store_count=len(expected_by_output),
        voucher_store_count=len(actual_by_output),
        matched_merchant_count=len(parsed_settlement.matched_merchants),
        unmatched_merchants=parsed_settlement.unmatched_merchants,
        differences=tuple(differences),
    )
