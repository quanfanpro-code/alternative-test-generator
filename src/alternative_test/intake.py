"""输入探测与角色识别 → 列映射确认 → 逐项追问。"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

from .adapters import ColumnMapping
from .fileio import TableCandidate
from .models import DecisionStatus

# ──────────────────────────────────────────────
# 角色评分关键词
# ──────────────────────────────────────────────

_BALANCE_KEYWORDS = [
    "余额", "期初", "期末", "年初", "年末",
    "供应商", "客户名称", "往来单位", "往来方", "被询证单位",
    "科目名称", "科目代码", "科目",
    "应付账款", "预付账款", "应收账款", "预收账款",
    "其他应收款", "其他应付款", "应付票据", "应收票据",
    "公司名称", "公司", "单位名称", "单位",
    "审计项目", "索引号",
]

_JOURNAL_KEYWORDS = [
    "凭证", "记账", "序时账", "日记账",
    "借方", "贷方", "借方金额", "贷方金额",
    "凭证号", "凭证类型", "凭证种类",
    "摘要", "业务说明", "业务内容",
    "对方科目", "辅助核算",
    "合计科目名称",
    "日期", "记账日期",
]


# ──────────────────────────────────────────────
# 数据类
# ──────────────────────────────────────────────

@dataclass
class RoleScore:
    """表格角色评分。"""
    candidate: TableCandidate
    role: Literal["balance", "journal", "unknown"]
    confidence: float  # 0.0 ~ 1.0


@dataclass
class ConfirmedMapping:
    """已确认的列映射。"""
    mapping: ColumnMapping
    status: DecisionStatus
    sheet_name: str = ""
    issues: list[str] = field(default_factory=list)


@dataclass
class Answers:
    """用户提供的回答（可积累）。"""
    period_start: str | None = None
    period_end: str | None = None
    entity_ids: list[str] = field(default_factory=list)
    normal_balance: str | None = None  # "credit" or "debit"


@dataclass
class Question:
    """单个待确认问题。"""
    question_id: str
    text: str
    category: str  # "period", "scope", "role", "mapping", "direction", "matching", "rule"
    options: list[str] = field(default_factory=list)


@dataclass
class QuestionQueue:
    """问题队列。"""
    questions: list[Question] = field(default_factory=list)
    _cursor: int = 0

    def has_questions(self) -> bool:
        return self._cursor < len(self.questions)

    def next(self) -> Question | None:
        if self.has_questions():
            q = self.questions[self._cursor]
            self._cursor += 1
            return q
        return None

    def peek(self) -> Question | None:
        if self.has_questions():
            return self.questions[self._cursor]
        return None


# ──────────────────────────────────────────────
# 角色评分
# ──────────────────────────────────────────────

def _score_headers(headers: list[str], keywords: list[str]) -> float:
    """计算表头中匹配关键词的比例（每个表头只计一次）。"""
    if not headers:
        return 0.0
    matched = 0
    for header in headers:
        for kw in keywords:
            if kw in header:
                matched += 1
                break
    return matched / len(headers)


def classify_tables(candidates: list[TableCandidate]) -> list[RoleScore]:
    """对候选表格进行角色评分。"""
    result: list[RoleScore] = []
    for cand in candidates:
        balance_score = _score_headers(cand.headers, _BALANCE_KEYWORDS)
        journal_score = _score_headers(cand.headers, _JOURNAL_KEYWORDS)

        if balance_score > journal_score and balance_score >= 0.30:
            confidence = balance_score
            role = "balance"
        elif journal_score > balance_score and journal_score >= 0.30:
            confidence = journal_score
            role = "journal"
        elif balance_score == journal_score and balance_score >= 0.30:
            # 同分时偏向 journal（更常见）
            role = "journal"
            confidence = balance_score
        else:
            role = "unknown"
            confidence = max(balance_score, journal_score)

        result.append(RoleScore(candidate=cand, role=role, confidence=confidence))

    return result


# ──────────────────────────────────────────────
# 列映射确认
# ──────────────────────────────────────────────

# 余额表必需字段
_BALANCE_REQUIRED_FIELDS = ["ending_balance_col", "counterparty_col"]

# 序时账必需字段
_JOURNAL_REQUIRED_FIELDS = ["date_col", "voucher_no_col", "debit_col", "credit_col"]

# 字段名 → CSV 常见别名
_FIELD_ALIASES: dict[str, list[str]] = {
    "date_col": ["日期", "记账日期", "凭证日期", "过账日期", "posting_date", "date"],
    "voucher_type_col": ["凭证类型", "凭证种类", "voucher_type"],
    "voucher_no_col": ["凭证号", "凭证编号", "序号", "voucher_no"],
    "summary_col": ["摘要", "业务说明", "业务内容", "description", "summary"],
    "account_col": ["科目", "科目名称", "合计科目名称", "会计科目", "account"],
    "debit_col": ["借方", "借方金额", "借", "debit"],
    "credit_col": ["贷方", "贷方金额", "贷", "credit"],
    "counterparty_aux_col": ["辅助核算", "供应商辅助核算", "客户辅助核算", "counterparty"],
    "entity_col": ["公司", "公司名称", "被审计单位", "核算单位", "单位名称", "entity"],
    "counterparty_col": ["供应商", "供应商名称", "客户名称", "往来单位", "被询证单位全称", "counterparty"],
    "opening_balance_col": ["期初余额", "年初余额", "期初", "opening"],
    "ending_balance_col": ["期末余额", "年末余额", "期末", "ending"],
    "account_code_col": ["科目代码", "科目编码", "account_code"],
}


def _auto_map_columns(
    candidate: TableCandidate,
    role: str,
) -> ColumnMapping:
    """基于表头自动匹配列名。"""
    mapping = ColumnMapping()
    fields = list(ColumnMapping.__dataclass_fields__.keys())

    if role == "balance":
        # 余额表优先匹配余额相关字段
        priority_fields = [
            "entity_col", "counterparty_col", "account_code_col",
            "opening_balance_col", "ending_balance_col",
        ]
        fields = [f for f in priority_fields if f in fields] + [
            f for f in fields if f not in priority_fields
        ]
    elif role == "journal":
        priority_fields = [
            "date_col", "voucher_type_col", "voucher_no_col",
            "debit_col", "credit_col", "account_col",
            "summary_col", "counterparty_aux_col",
        ]
        fields = [f for f in priority_fields if f in fields] + [
            f for f in fields if f not in priority_fields
        ]

    for field_name in fields:
        aliases = _FIELD_ALIASES.get(field_name, [])
        for alias in aliases:
            if alias in candidate.headers:
                setattr(mapping, field_name, alias)
                break

    return mapping


def confirm_mapping(
    candidate: TableCandidate,
    mapping: ColumnMapping,
    role: str = "journal",
) -> ConfirmedMapping:
    """确认列映射：自动匹配 + 检查必需字段。"""
    issues: list[str] = []

    # 自动补全未设置的字段
    if mapping.date_col is None:
        auto = _auto_map_columns(candidate, role)
        for field_name in ColumnMapping.__dataclass_fields__:
            if getattr(mapping, field_name) is None:
                auto_val = getattr(auto, field_name)
                if auto_val is not None:
                    setattr(mapping, field_name, auto_val)

    # 检查必需字段
    if role == "journal":
        required = _JOURNAL_REQUIRED_FIELDS
    else:
        required = _BALANCE_REQUIRED_FIELDS

    for field_name in required:
        if getattr(mapping, field_name) is None:
            issues.append(f"缺少必需字段映射: {field_name}")

    status = DecisionStatus.CONFIRMED if not issues else DecisionStatus.INFERRED
    return ConfirmedMapping(
        mapping=mapping,
        status=status,
        sheet_name=candidate.name,
        issues=issues,
    )


# ──────────────────────────────────────────────
# 问题队列
# ──────────────────────────────────────────────

def build_question_queue(
    scores: list[RoleScore],
    answers: Answers,
) -> QuestionQueue:
    """根据角色评分和已有答案构建逐项追问队列。"""
    questions: list[Question] = []

    # 1. 期间问题（最高优先级）
    if not answers.period_start or not answers.period_end:
        questions.append(Question(
            question_id="period",
            text="请提供审计期间（起止日期）",
            category="period",
        ))

    # 2. 公司/账套范围
    if not answers.entity_ids:
        questions.append(Question(
            question_id="scope",
            text="请确认公司/账套范围",
            category="scope",
        ))

    # 3. 角色确认（如有歧义）
    unknowns = [s for s in scores if s.role == "unknown"]
    if unknowns:
        sheet_names = [s.candidate.name for s in unknowns]
        questions.append(Question(
            question_id="role",
            text=f"以下工作表无法自动识别角色，请确认: {', '.join(sheet_names)}",
            category="role",
            options=["余额表", "序时账", "忽略"],
        ))

    # 4. 列映射确认（如有 INFERRED 字段）
    # 此步骤在 agent 调用 confirm_mapping 后由外部推动

    # 5. 方向确认
    if not answers.normal_balance:
        questions.append(Question(
            question_id="direction",
            text="请确认科目正常余额方向（应付/预收/其他应付款通常为贷方，应收/预付/其他应收款通常为借方）",
            category="direction",
            options=["credit（贷方）", "debit（借方）"],
        ))

    # 6. 匹配确认（在 matching step 中处理）

    # 7. 规则确认
    questions.append(Question(
        question_id="rule",
        text="请提供候选选择规则，若无规则将输出完整候选总体",
        category="rule",
    ))

    return QuestionQueue(questions=questions)
