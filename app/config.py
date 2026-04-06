from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class PaymentMethodConfig:
    key: str
    source_card_name: str | None
    include_from_card_file: bool
    daily_field: str | None
    partner_code: str
    partner_name: str
    management_data: str | None
    order: int


@dataclass(frozen=True)
class StoreMapping:
    output_name: str
    source_name: str
    partner_code: str
    partner_name: str
    enabled: bool = True
    manual_review_reason: str | None = None


@dataclass(frozen=True)
class BusinessConfig:
    key: str
    display_name: str
    doc_prefix: str
    consul_dc_format: str
    note_dc_format: str
    company_code: str
    pc_code: str
    gaap_code: str
    docu_code: str
    wrt_dept_code: str
    wrt_emp_no: str
    insr_fg_cd: str
    account_codes: dict[str, str]
    stores: tuple[StoreMapping, ...]

    @property
    def active_stores(self) -> tuple[StoreMapping, ...]:
        return tuple(store for store in self.stores if store.enabled)

    @property
    def source_store_names(self) -> set[str]:
        return {store.source_name for store in self.active_stores}


@dataclass(frozen=True)
class AppConfig:
    businesses: dict[str, BusinessConfig]
    payment_methods: tuple[PaymentMethodConfig, ...]
    settlement_partner_payment_key: str = "common"
    settlement_management_data: str = "11/당근정산"

    @property
    def payment_methods_in_order(self) -> tuple[PaymentMethodConfig, ...]:
        return tuple(sorted(self.payment_methods, key=lambda item: item.order))

    @property
    def payment_by_key(self) -> dict[str, PaymentMethodConfig]:
        return {payment.key: payment for payment in self.payment_methods}

    @property
    def payment_by_source_card_name(self) -> dict[str, PaymentMethodConfig]:
        return {
            payment.source_card_name: payment
            for payment in self.payment_methods
            if payment.include_from_card_file and payment.source_card_name
        }


def _load_store_mapping(raw: dict[str, Any]) -> StoreMapping:
    return StoreMapping(
        output_name=raw["output_name"],
        source_name=raw["source_name"],
        partner_code=raw["partner_code"],
        partner_name=raw["partner_name"],
        enabled=raw.get("enabled", True),
        manual_review_reason=raw.get("manual_review_reason"),
    )


def _load_business(key: str, raw: dict[str, Any]) -> BusinessConfig:
    stores = tuple(_load_store_mapping(item) for item in raw["stores"])
    return BusinessConfig(
        key=key,
        display_name=raw["display_name"],
        doc_prefix=raw["doc_prefix"],
        consul_dc_format=raw["consul_dc_format"],
        note_dc_format=raw["note_dc_format"],
        company_code=raw["company_code"],
        pc_code=raw["pc_code"],
        gaap_code=raw["gaap_code"],
        docu_code=raw["docu_code"],
        wrt_dept_code=raw["wrt_dept_code"],
        wrt_emp_no=raw["wrt_emp_no"],
        insr_fg_cd=raw["insr_fg_cd"],
        account_codes=raw["account_codes"],
        stores=stores,
    )


def _load_payment_method(raw: dict[str, Any]) -> PaymentMethodConfig:
    return PaymentMethodConfig(
        key=raw["key"],
        source_card_name=raw.get("source_card_name"),
        include_from_card_file=raw.get("include_from_card_file", False),
        daily_field=raw.get("daily_field"),
        partner_code=raw["partner_code"],
        partner_name=raw["partner_name"],
        management_data=raw.get("management_data"),
        order=raw["order"],
    )


def load_config(path: Path) -> AppConfig:
    raw = json.loads(path.read_text(encoding="utf-8"))
    businesses = {
        key: _load_business(key, value) for key, value in raw["businesses"].items()
    }
    payment_methods = tuple(_load_payment_method(item) for item in raw["payment_methods"])
    settlement = raw.get("settlement_order", {})
    return AppConfig(
        businesses=businesses,
        payment_methods=payment_methods,
        settlement_partner_payment_key=settlement.get("partner_payment_key", "common"),
        settlement_management_data=settlement.get("management_data", "11/당근정산"),
    )


def load_store_name_mapping(path: Path) -> dict[str, dict[str, tuple[str, ...]]]:
    if not path.exists():
        return {}

    raw = json.loads(path.read_text(encoding="utf-8"))
    result: dict[str, dict[str, tuple[str, ...]]] = {}
    for business_key, mapping in raw.items():
        business_mapping: dict[str, tuple[str, ...]] = {}
        if not isinstance(mapping, dict):
            continue
        for source_name, aliases in mapping.items():
            if isinstance(source_name, str) and isinstance(aliases, list):
                alias_values = tuple(str(alias) for alias in aliases if isinstance(alias, str))
                business_mapping[source_name] = alias_values
        result[business_key] = business_mapping
    return result
