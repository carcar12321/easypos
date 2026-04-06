from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from app.config import AppConfig


class ParsingError(ValueError):
    pass


@dataclass(frozen=True)
class DailySalesRow:
    source_name: str
    gross_sales: float
    discount_amount: float
    cash_sales: float
    electronic_money_sales: float


@dataclass(frozen=True)
class ParsedCardSales:
    account_date: str | None
    by_store: dict[str, dict[str, float]]


@dataclass(frozen=True)
class ParsedDailySales:
    account_date: str | None
    by_store: dict[str, DailySalesRow]


@dataclass(frozen=True)
class ParsedSettlementOrders:
    account_date: str | None
    by_store: dict[str, float]
    by_merchant: dict[str, float]
    matched_merchants: dict[str, str]
    unmatched_merchants: tuple[str, ...]


@dataclass(frozen=True)
class ParsedVoucherSettlement:
    account_date: str | None
    by_output_store: dict[str, float]


def _normalize_date_digits(digits: str) -> str | None:
    normalized = digits
    if len(normalized) == 6:
        normalized = f"20{normalized}"
    if len(normalized) != 8:
        return None
    try:
        datetime.strptime(normalized, "%Y%m%d")
    except ValueError:
        return None
    return normalized


def _as_yyyymmdd(value: Any) -> str | None:
    if isinstance(value, datetime):
        return value.strftime("%Y%m%d")
    if isinstance(value, date):
        return value.strftime("%Y%m%d")
    if isinstance(value, str):
        digits = re.sub(r"[^0-9]", "", value)
        return _normalize_date_digits(digits)
    return None


def _extract_account_date_from_filename(filename: str) -> str | None:
    match_with_separator = re.search(r"(20\d{2})[-_.](\d{2})[-_.](\d{2})", filename)
    if match_with_separator:
        return _normalize_date_digits("".join(match_with_separator.groups()))

    match_8 = re.search(r"(20\d{6})", filename)
    if match_8:
        return _normalize_date_digits(match_8.group(1))

    match_6 = re.search(r"(?<!\d)(\d{6})(?!\d)", filename)
    if match_6:
        return _normalize_date_digits(match_6.group(1))

    return None


def _extract_header_row(worksheet, required_headers: set[str], search_rows: int = 5) -> tuple[int, dict[str, int]] | None:
    for row_index in range(1, min(worksheet.max_row, search_rows) + 1):
        headers: dict[str, int] = {}
        for column_index in range(1, worksheet.max_column + 1):
            value = worksheet.cell(row_index, column_index).value
            if isinstance(value, str) and value.strip():
                headers[value.strip()] = column_index
        if required_headers.issubset(headers):
            return row_index, headers
    return None


def _normalize_store_key(value: str) -> str:
    normalized = re.sub(r"[^0-9A-Za-z가-힣]", "", value).lower()
    return normalized.strip()


def _find_card_sheet(path: Path):
    workbook = load_workbook(path, data_only=True)
    required_headers = {"가맹점명", "구분", "승인일자", "카드사명", "승인금액"}

    for worksheet in workbook.worksheets:
        extracted = _extract_header_row(worksheet, required_headers, search_rows=5)
        if extracted is not None:
            row_index, headers = extracted
            return worksheet, headers, row_index

    raise ParsingError(f"카드매출 파일에서 필수 헤더를 찾지 못했습니다: {path.name}")


def parse_card_sales(path: Path, config: AppConfig, allow_missing_date: bool = False) -> ParsedCardSales:
    worksheet, headers, header_row_index = _find_card_sheet(path)
    payment_by_source = config.payment_by_source_card_name
    by_store: dict[str, dict[str, float]] = {}
    account_date: str | None = None

    for row_index in range(header_row_index + 1, worksheet.max_row + 1):
        store_name = worksheet.cell(row_index, headers["가맹점명"]).value
        status = worksheet.cell(row_index, headers["구분"]).value
        card_name = worksheet.cell(row_index, headers["카드사명"]).value
        amount = worksheet.cell(row_index, headers["승인금액"]).value
        approval_date = worksheet.cell(row_index, headers["승인일자"]).value

        if not store_name or not card_name or not isinstance(amount, (int, float)):
            continue

        if account_date is None:
            account_date = _as_yyyymmdd(approval_date)

        payment = payment_by_source.get(str(card_name))
        if payment is None:
            continue

        # 카드매출은 승인금액만 사용한다.
        if status != "승인":
            continue

        store_bucket = by_store.setdefault(str(store_name), {})
        store_bucket[payment.key] = store_bucket.get(payment.key, 0.0) + float(amount)

    if account_date is None and not allow_missing_date:
        raise ParsingError(f"카드매출 파일에서 회계일을 찾지 못했습니다: {path.name}")

    return ParsedCardSales(account_date=account_date, by_store=by_store)


def _parse_daily_sales_new_format(path: Path, worksheet, allow_missing_date: bool) -> ParsedDailySales:
    required_headers = {"매장명", "영업일자", "총매출", "할인", "현금매출"}
    extracted = _extract_header_row(worksheet, required_headers, search_rows=5)
    if extracted is None:
        raise ParsingError(f"새양식 매출내역 헤더를 찾지 못했습니다: {path.name}")
    header_row_index, headers = extracted

    account_date = _extract_account_date_from_filename(path.name)
    if account_date is None:
        for row_index in range(header_row_index + 1, worksheet.max_row + 1):
            row_account_date = _as_yyyymmdd(worksheet.cell(row_index, headers["영업일자"]).value)
            if row_account_date:
                account_date = row_account_date
                break

    if account_date is None and not allow_missing_date:
        raise ParsingError(f"새양식 파일에서 회계일을 찾지 못했습니다: {path.name}")

    by_store: dict[str, DailySalesRow] = {}
    for row_index in range(header_row_index + 1, worksheet.max_row + 1):
        source_name = worksheet.cell(row_index, headers["매장명"]).value
        gross_sales = worksheet.cell(row_index, headers["총매출"]).value
        discount_amount = worksheet.cell(row_index, headers["할인"]).value
        cash_sales = worksheet.cell(row_index, headers["현금매출"]).value
        row_account_date = _as_yyyymmdd(worksheet.cell(row_index, headers["영업일자"]).value)

        if not isinstance(source_name, str) or not source_name.strip():
            continue
        if not isinstance(gross_sales, (int, float)):
            continue
        if account_date is not None and row_account_date is not None and row_account_date != account_date:
            continue

        key = source_name.strip()
        current = by_store.get(key)
        if current is None:
            by_store[key] = DailySalesRow(
                source_name=key,
                gross_sales=float(gross_sales or 0.0),
                discount_amount=float(discount_amount or 0.0),
                cash_sales=float(cash_sales or 0.0),
                electronic_money_sales=0.0,
            )
            continue

        by_store[key] = DailySalesRow(
            source_name=key,
            gross_sales=current.gross_sales + float(gross_sales or 0.0),
            discount_amount=current.discount_amount + float(discount_amount or 0.0),
            cash_sales=current.cash_sales + float(cash_sales or 0.0),
            electronic_money_sales=current.electronic_money_sales,
        )

    return ParsedDailySales(account_date=account_date, by_store=by_store)


def _parse_daily_sales_legacy_format(path: Path, worksheet, allow_missing_date: bool) -> ParsedDailySales:
    account_date = _extract_account_date_from_filename(path.name)
    if account_date is None and not allow_missing_date:
        raise ParsingError(f"당일 매출내역 파일명에서 회계일을 찾지 못했습니다: {path.name}")

    by_store: dict[str, DailySalesRow] = {}
    for row_index in range(3, worksheet.max_row + 1):
        source_name = worksheet.cell(row_index, 3).value
        gross_sales = worksheet.cell(row_index, 5).value
        discount_amount = worksheet.cell(row_index, 8).value
        cash_sales = worksheet.cell(row_index, 11).value
        electronic_money_sales = worksheet.cell(row_index, 20).value

        if not source_name or not isinstance(gross_sales, (int, float)):
            continue

        by_store[str(source_name)] = DailySalesRow(
            source_name=str(source_name),
            gross_sales=float(gross_sales or 0.0),
            discount_amount=float(discount_amount or 0.0),
            cash_sales=float(cash_sales or 0.0),
            electronic_money_sales=float(electronic_money_sales or 0.0),
        )

    return ParsedDailySales(account_date=account_date, by_store=by_store)


def parse_daily_sales(path: Path, allow_missing_date: bool = False) -> ParsedDailySales:
    workbook = load_workbook(path, data_only=True)
    worksheet = workbook.worksheets[0]

    new_format_headers = {"매장명", "영업일자", "총매출", "할인", "현금매출"}
    if _extract_header_row(worksheet, new_format_headers, search_rows=5) is not None:
        return _parse_daily_sales_new_format(path, worksheet, allow_missing_date)
    return _parse_daily_sales_legacy_format(path, worksheet, allow_missing_date)


def parse_settlement_orders(
    path: Path,
    *,
    source_names: set[str],
    store_aliases: dict[str, tuple[str, ...]] | None = None,
    expected_account_date: str | None = None,
    allow_missing_date: bool = False,
) -> ParsedSettlementOrders:
    workbook = load_workbook(path, data_only=True)
    worksheet = workbook.worksheets[0]
    required_headers = {"가맹점", "결제금액"}
    extracted = _extract_header_row(worksheet, required_headers, search_rows=10)
    if extracted is None:
        raise ParsingError(f"정산주문목록 파일에서 필수 헤더를 찾지 못했습니다: {path.name}")
    header_row_index, headers = extracted

    alias_index: dict[str, str] = {}
    for source_name in source_names:
        source_norm = _normalize_store_key(source_name)
        alias_index[source_norm] = source_name
        if store_aliases and source_name in store_aliases:
            for alias in store_aliases[source_name]:
                alias_norm = _normalize_store_key(alias)
                if alias_norm:
                    alias_index[alias_norm] = source_name

    account_date = expected_account_date or _extract_account_date_from_filename(path.name)
    merchant_totals: dict[str, float] = {}
    matched_merchants: dict[str, str] = {}
    by_store: dict[str, float] = {}
    unmatched_merchants: set[str] = set()

    for row_index in range(header_row_index + 1, worksheet.max_row + 1):
        merchant = worksheet.cell(row_index, headers["가맹점"]).value
        amount = worksheet.cell(row_index, headers["결제금액"]).value
        approval_date = (
            worksheet.cell(row_index, headers["결제 승인일"]).value
            if "결제 승인일" in headers
            else None
        )
        row_account_date = _as_yyyymmdd(approval_date)

        if account_date is None and row_account_date:
            account_date = row_account_date
        if account_date is not None and row_account_date is not None and row_account_date != account_date:
            continue

        if not isinstance(merchant, str) or not merchant.strip():
            continue
        if not isinstance(amount, (int, float)):
            continue

        merchant_name = merchant.strip()
        merchant_totals[merchant_name] = merchant_totals.get(merchant_name, 0.0) + float(amount)

    if account_date is None and not allow_missing_date:
        raise ParsingError(f"정산주문목록 파일에서 회계일을 찾지 못했습니다: {path.name}")

    for merchant_name, total_amount in merchant_totals.items():
        merchant_norm = _normalize_store_key(merchant_name)
        source_name = alias_index.get(merchant_norm)

        if source_name is None:
            # 보조 매칭: source_name 일부 문자열이 merchant 이름에 포함되거나 그 반대
            candidates: list[tuple[int, str]] = []
            for source in source_names:
                source_norm = _normalize_store_key(source)
                if not source_norm:
                    continue
                if source_norm in merchant_norm or merchant_norm in source_norm:
                    candidates.append((len(source_norm), source))
            if candidates:
                source_name = sorted(candidates, reverse=True)[0][1]

        if source_name is None:
            unmatched_merchants.add(merchant_name)
            continue

        matched_merchants[merchant_name] = source_name
        by_store[source_name] = by_store.get(source_name, 0.0) + total_amount

    return ParsedSettlementOrders(
        account_date=account_date,
        by_store=by_store,
        by_merchant=merchant_totals,
        matched_merchants=matched_merchants,
        unmatched_merchants=tuple(sorted(unmatched_merchants)),
    )


def parse_generated_voucher_settlement(
    path: Path,
    *,
    settlement_partner_code: str,
    settlement_management_data: str | None = None,
) -> ParsedVoucherSettlement:
    workbook = load_workbook(path, data_only=True)
    worksheet = workbook.worksheets[0]

    by_output_store: dict[str, float] = {}
    account_date: str | None = None
    row_index = 4
    note_pattern = re.compile(r"POS매출\((.+)\)")

    while True:
        docu_no = worksheet.cell(row_index, 1).value
        if docu_no is None:
            break

        actg_dt = worksheet.cell(row_index, 4).value
        if account_date is None and actg_dt:
            account_date = _as_yyyymmdd(str(actg_dt))

        account_code = worksheet.cell(row_index, 8).value
        amount = worksheet.cell(row_index, 9).value
        note_dc = worksheet.cell(row_index, 11).value
        partner_code = worksheet.cell(row_index, 12).value
        management_data = worksheet.cell(row_index, 15).value

        if str(account_code) != "102700":
            row_index += 1
            continue
        if str(partner_code) != settlement_partner_code:
            row_index += 1
            continue
        if settlement_management_data is not None and str(management_data) != settlement_management_data:
            row_index += 1
            continue
        if not isinstance(amount, (int, float)):
            row_index += 1
            continue
        if not isinstance(note_dc, str):
            row_index += 1
            continue

        matched = note_pattern.search(note_dc)
        if not matched:
            row_index += 1
            continue

        output_name = matched.group(1).strip()
        by_output_store[output_name] = by_output_store.get(output_name, 0.0) + float(amount)
        row_index += 1

    return ParsedVoucherSettlement(account_date=account_date, by_output_store=by_output_store)
