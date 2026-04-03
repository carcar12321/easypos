from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any

from openpyxl import load_workbook

from app.config import AppConfig, BusinessConfig, PaymentMethodConfig, StoreMapping
from app.parsers import DailySalesRow, ParsedCardSales, ParsedDailySales, ParsingError, parse_card_sales, parse_daily_sales


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
    amounts = dict(card_sales.by_store.get(store.source_name, {}))
    daily_row = daily_sales.by_store.get(store.source_name)
    if not daily_row:
        return amounts

    for payment in payment_methods:
        if payment.daily_field == "전자화폐" and daily_row.electronic_money_sales:
            amounts[payment.key] = amounts.get(payment.key, 0.0) + daily_row.electronic_money_sales

    return amounts


def _build_store_lines(
    business: BusinessConfig,
    store: StoreMapping,
    daily_row: DailySalesRow,
    payment_amounts: dict[str, float],
    payment_methods: tuple[PaymentMethodConfig, ...],
    start_seq: int,
    account_date: str,
    include_consul_dc: bool,
) -> list[VoucherLine]:
    common = _line_common(business, account_date)
    note_dc = _build_note_dc(business, account_date, store.output_name)
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
        amount=daily_row.gross_sales,
        partner_cd=store.partner_code,
        partner_nm=store.partner_name,
        cc_cd="1300",
        management_data=None,
        consul_dc=_build_consul_dc(business, account_date) if include_consul_dc else None,
    )

    if daily_row.discount_amount:
        add_line(
            drcrfg_cd="2",
            acct_cd=business.account_codes["discount"],
            amount=(-1) * daily_row.discount_amount,
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


def generate_voucher(
    *,
    business: BusinessConfig,
    card_file_path: Path,
    daily_file_path: Path,
    template_file_path: Path,
    config: AppConfig,
    output_dir: Path,
) -> GenerationArtifact:
    parsed_card = parse_card_sales(card_file_path, config)
    parsed_daily = parse_daily_sales(daily_file_path)

    if parsed_card.account_date != parsed_daily.account_date:
        raise ParsingError(
            "카드매출/당일매출 파일의 회계일이 다릅니다. "
            f"카드={parsed_card.account_date}, 당일={parsed_daily.account_date}"
        )

    account_date = parsed_card.account_date
    notes: list[str] = []
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

        has_amount = any(
            [
                daily_row.gross_sales,
                daily_row.discount_amount,
                daily_row.cash_sales,
                daily_row.electronic_money_sales,
                sum(amounts.values()),
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

