"""稳定数据模型 — Pydantic v2、Decimal 金额、来源引用。"""

from __future__ import annotations

from datetime import date
from decimal import Decimal
from enum import Enum
from typing import Literal

from pydantic import BaseModel, Field

# ──────────────────────────────────────────────
# 枚举
# ──────────────────────────────────────────────

class DecisionStatus(str, Enum):
    MISSING = "missing"
    INFERRED = "inferred_needs_confirmation"
    CONFIRMED = "user_confirmed"
    UNAVAILABLE = "unavailable"
    NOT_APPLICABLE = "not_applicable"


class NormalBalance(str, Enum):
    DEBIT = "debit"
    CREDIT = "credit"


# ──────────────────────────────────────────────
# 基础模型
# ──────────────────────────────────────────────

class SourceRef(BaseModel):
    """数据来源引用 — 文件/Sheet/行/列。"""

    model_config = {"extra": "forbid"}

    file_id: str
    sha256: str
    sheet: str | None = None
    row: int
    columns: dict[str, str] = Field(default_factory=dict)


class ConfirmedValue(BaseModel):
    """用户确认的值 — 字段级追溯。"""

    model_config = {"extra": "forbid"}

    value: str | None = None
    status: DecisionStatus = DecisionStatus.MISSING
    confirmed_by: str | None = None
    decision_id: str | None = None


# ──────────────────────────────────────────────
# 业务模型
# ──────────────────────────────────────────────

class CounterpartyBalance(BaseModel):
    """供应商/往来余额。"""

    model_config = {"extra": "forbid", "ser_json_bytes": "utf8"}

    entity_id: str
    counterparty_id: str
    raw_name: str
    account_code: str | None = None
    account_name: str
    normal_balance: NormalBalance
    opening_balance: Decimal | None = None
    ending_balance: Decimal
    source: SourceRef


class JournalLine(BaseModel):
    """序时账行 — 保留正负号。"""

    model_config = {"extra": "forbid", "ser_json_bytes": "utf8"}

    entity_id: str
    ledger_id: str
    posting_date: date
    voucher_type: str
    voucher_number: str
    line_number: int
    account_code: str | None = None
    account_name: str
    counterparty_raw: str | None = None
    summary: str
    debit: Decimal
    credit: Decimal
    source: SourceRef

    @property
    def voucher_key(self) -> tuple[str, str, date, str, str]:
        """稳定凭证键：entity、账套、日期、凭证类型、凭证号。"""
        return (
            self.entity_id,
            self.ledger_id,
            self.posting_date,
            self.voucher_type,
            self.voucher_number,
        )


class MatchDecision(BaseModel):
    """对手方匹配决定。"""

    model_config = {"extra": "forbid"}

    journal_line_id: str
    counterparty_id: str | None = None
    status: Literal[
        "exact_unique", "needs_confirmation", "confirmed", "unmatched", "rejected"
    ]
    candidate_ids: list[str] = Field(default_factory=list)
    reason: str
    decision_id: str | None = None


class ProcedureResult(BaseModel):
    """检查程序结果 — 未确认字段必须为 null/空白。"""

    model_config = {"extra": "forbid"}

    candidate_id: str
    attachment_status: ConfirmedValue = Field(default_factory=ConfirmedValue)
    procedure_result: ConfirmedValue = Field(default_factory=ConfirmedValue)
    exception_description: ConfirmedValue = Field(default_factory=ConfirmedValue)
    conclusion: ConfirmedValue = Field(default_factory=ConfirmedValue)


# ──────────────────────────────────────────────
# JSON 序列化钩子
# ──────────────────────────────────────────────

class _DecimalEncoder:
    """Pydantic 自定义序列化：Decimal 写为十进制字符串。"""

    @staticmethod
    def _serialize_decimal(v: object) -> str:
        if isinstance(v, Decimal):
            return str(v)
        raise TypeError(type(v))
