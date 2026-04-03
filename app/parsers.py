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
    account_date: str
    by_store: dict[str, dict[str, float]]


@dataclass(frozen=True)
class ParsedDailySales:
    account_date: str
    by_store: dict[str, DailySalesRow]


def _as_yyyymmdd(value: Any) -> str | None:
    if isinstance(value, datetime):
        return value.strftime("%Y%m%d")
    if isinstance(value, date):
        return value.strftime("%Y%m%d")
    if isinstance(value, str):
        digits = re.sub(r"[^0-9]", "", value)
        if len(digits) == 8:
            return digits
    return None


def _find_card_sheet(path: Path):
    workbook = load_workbook(path, data_only=True)
    required_headers = {"가맹점명", "구분", "승인일자", "카드사명", "승인금액"}

    for worksheet in workbook.worksheets:
        for row_index in range(1, min(worksheet.max_row, 5) + 1):
            headers = {}
            for column_index in range(1, worksheet.max_column + 1):
                value = worksheet.cell(row_index, column_index).value
                if isinstance(value, str) and value.strip():
                    headers[value.strip()] = column_index
            if required_headers.issubset(headers):
                return worksheet, headers, row_index

    raise ParsingError(f"카드매출 파일에서 필수 헤더를 찾지 못했습니다: {path.name}")


def parse_card_sales(path: Path, config: AppConfig) -> ParsedCardSales:
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

        direction = 0
        if status == "승인":
            direction = 1
        elif status == "취소":
            direction = -1
        else:
            continue

        store_bucket = by_store.setdefault(str(store_name), {})
        store_bucket[payment.key] = store_bucket.get(payment.key, 0.0) + (float(amount) * direction)

    if account_date is None:
        raise ParsingError(f"카드매출 파일에서 회계일을 찾지 못했습니다: {path.name}")

    return ParsedCardSales(account_date=account_date, by_store=by_store)


def parse_daily_sales(path: Path) -> ParsedDailySales:
    workbook = load_workbook(path, data_only=True)
    worksheet = workbook.worksheets[0]
    match = re.search(r"(20\d{6})", path.name)
    if not match:
        raise ParsingError(f"당일 매출내역 파일명에서 회계일을 찾지 못했습니다: {path.name}")
    account_date = match.group(1)

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

