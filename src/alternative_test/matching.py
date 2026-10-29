"""供应商唯一匹配 — 只做全字符串精确且唯一时自动确认，其余均需用户确认。"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field

from .models import CounterpartyBalance, JournalLine, MatchDecision

# ──────────────────────────────────────────────
# exact_key 规范化
# ──────────────────────────────────────────────

def exact_key(value: str) -> str:
    """名称规范化：NFKC + 全角/半角 + 折叠空白 + 大小写统一。"""
    # Unicode NFKC
    s = unicodedata.normalize("NFKC", value)
    # 全角空格 → 半角空格
    s = s.replace("\u3000", " ")
    # 全角字母/数字 → 半角
    result: list[str] = []
    for ch in s:
        code = ord(ch)
        if 0xFF01 <= code <= 0xFF5E:
            result.append(chr(code - 0xFEE0))
        else:
            result.append(ch)
    s = "".join(result)
    # 删除 CJK 字符之间的空白（中日韩文本中空格不具语义）
    s = re.sub(r"([\u4e00-\u9fff\u3400-\u4dbf\uff00-\uffef])\s+([\u4e00-\u9fff\u3400-\u4dbf\uff00-\uffef])", r"\1\2", s)
    # 首尾空白删除
    s = s.strip()
    # 连续空白折叠为单个空格
    s = re.sub(r"\s+", " ", s)
    # ASCII 大小写统一
    s = s.lower()
    return s


# ──────────────────────────────────────────────
# 辅助：去掉公司后缀和括号
# ──────────────────────────────────────────────

_SUFFIX_PATTERN = re.compile(
    r"(股份)?有限公司|有限责任公司|分公司|（普通合伙）|\(普通合伙\)|厂$|集团"
)


def _strip_suffix(name: str) -> str:
    """去掉常见公司后缀。"""
    return _SUFFIX_PATTERN.sub("", name).strip()


def _strip_parens(name: str) -> str:
    """去掉中英文括号及其中内容。"""
    return re.sub(r"[（(][^）)]*[）)]", "", name).strip()


def _strip_punctuation(name: str) -> str:
    """去掉标点符号。"""
    return re.sub(r"[^\w\u4e00-\u9fff\s]", "", name).strip()


# ──────────────────────────────────────────────
# 匹配计划
# ──────────────────────────────────────────────

@dataclass
class MatchPlan:
    """待确认匹配计划。"""
    decisions: list[MatchDecision] = field(default_factory=list)
    summary: str = ""


@dataclass
class UserDecision:
    """用户对匹配候选的确认决定。"""
    decision_id: str
    journal_line_id: str
    counterparty_id: str | None  # None = 拒绝/不匹配
    original_text: str = ""


# ──────────────────────────────────────────────
# propose_matches
# ──────────────────────────────────────────────

def propose_matches(
    lines: list[JournalLine],
    balances: list[CounterpartyBalance],
) -> MatchPlan:
    """为序时账行生成匹配计划。

    分层：
    1. exact_unique：规范化后全字符串精确且唯一
    2. needs_confirmation：去后缀、去括号、简称、摘要匹配
    3. unmatched：无法匹配
    """
    decisions: list[MatchDecision] = []

    # 构建余额表规范化名称索引
    balance_keys: dict[str, list[str]] = {}
    for b in balances:
        ek = exact_key(b.raw_name)
        if ek not in balance_keys:
            balance_keys[ek] = []
        balance_keys[ek].append(b.counterparty_id)

    for line in lines:
        scoped = [b for b in balances if b.entity_id == line.entity_id and b.account_name == line.account_name]
        scoped_keys: dict[str, list[str]] = {}
        for balance in scoped:
            scoped_keys.setdefault(exact_key(balance.raw_name), []).append(balance.counterparty_id)
        decisions.append(_match_one_line(line, scoped, scoped_keys))

    return MatchPlan(decisions=decisions)


def _match_one_line(
    line: JournalLine,
    balances: list[CounterpartyBalance],
    balance_keys: dict[str, list[str]],
) -> MatchDecision:
    """匹配单行序时账。

    分层：
    1. exact_unique：exact_key 相等且唯一
    2. needs_confirmation：去后缀、去括号、子串、摘要
    3. unmatched
    """
    journal_line_id = _make_line_id(line)

    # 优先用辅助核算中的名称
    search_text = line.counterparty_raw or ""
    search_key = exact_key(search_text) if search_text else ""

    # ── 第一层：exact_unique ──
    if search_key and search_key in balance_keys and len(balance_keys[search_key]) == 1:
        return MatchDecision(
            journal_line_id=journal_line_id,
            counterparty_id=balance_keys[search_key][0],
            status="exact_unique",
            candidate_ids=list(balance_keys[search_key]),
            reason=f"全字符串精确匹配: {search_text}",
        )

    # ── 第二层：候选生成 ──
    candidate_ids: list[str] = []
    reason_parts: list[str] = []

    # 对各余额表供应商产生候选
    for b in balances:
        b_name = b.raw_name

        # A. 去后缀后相等（双方都去后缀）
        line_no_suffix = _strip_suffix(search_text)
        bal_no_suffix = _strip_suffix(b_name)
        if (
            len(line_no_suffix) >= 1
            and len(bal_no_suffix) >= 1
            and exact_key(line_no_suffix) == exact_key(bal_no_suffix)
        ):
            if b.counterparty_id not in candidate_ids:
                candidate_ids.append(b.counterparty_id)
                reason_parts.append(f"去后缀: {line_no_suffix} ↔ {bal_no_suffix}")

        # B. 去括号后相等
        line_no_paren = exact_key(_strip_parens(search_text))
        bal_no_paren = exact_key(_strip_parens(b_name))
        if line_no_paren and bal_no_paren and line_no_paren == bal_no_paren:
            if b.counterparty_id not in candidate_ids:
                candidate_ids.append(b.counterparty_id)
                reason_parts.append(f"去括号: {_strip_parens(search_text)} ↔ {_strip_parens(b_name)}")

        # C. 子串包含（≥4 字符片段）
        line_ek = exact_key(search_text) if search_text else ""
        bal_ek = exact_key(b_name)
        if len(line_ek) >= 4 and len(bal_ek) >= 4:
            if line_ek in bal_ek or bal_ek in line_ek:
                if b.counterparty_id not in candidate_ids:
                    candidate_ids.append(b.counterparty_id)
                    reason_parts.append(f"子串包含: {search_text} ↔ {b_name}")

        # D. 摘要匹配
        if line.summary:
            summary_ek = exact_key(line.summary)
            if len(bal_ek) >= 4 and bal_ek in summary_ek:
                if b.counterparty_id not in candidate_ids:
                    candidate_ids.append(b.counterparty_id)
                    reason_parts.append(f"摘要含: {b_name}")

            # 摘要子串匹配（4字符片段）
            if not candidate_ids or b.counterparty_id not in candidate_ids:
                for i in range(len(bal_ek) - 3):
                    chunk = bal_ek[i:i + 4]
                    if chunk in summary_ek and b.counterparty_id not in candidate_ids:
                        candidate_ids.append(b.counterparty_id)
                        reason_parts.append(f"摘要子串({chunk}): {b_name}")
                        break

    # ── 第三层：判定 ──
    if candidate_ids:
        return MatchDecision(
            journal_line_id=journal_line_id,
            counterparty_id=None,
            status="needs_confirmation",
            candidate_ids=candidate_ids,
            reason="; ".join(reason_parts) if reason_parts else "候选生成",
        )
    else:
        return MatchDecision(
            journal_line_id=journal_line_id,
            counterparty_id=None,
            status="unmatched",
            candidate_ids=[],
            reason=f"未匹配到供应商: '{search_text}'",
        )


def _make_line_id(line: JournalLine) -> str:
    """生成唯一行 ID。"""
    return (
        f"{line.entity_id}_{line.ledger_id}_"
        f"{line.posting_date.isoformat()}_"
        f"{line.voucher_type}_{line.voucher_number}_{line.line_number}"
    )


# ──────────────────────────────────────────────
# apply_confirmed_decisions
# ──────────────────────────────────────────────

def apply_confirmed_decisions(
    plan: MatchPlan,
    decisions: list[UserDecision],
) -> MatchPlan:
    """应用用户确认的决定到匹配计划。"""
    decision_map = {d.journal_line_id: d for d in decisions}
    updated: list[MatchDecision] = []

    for md in plan.decisions:
        if md.journal_line_id in decision_map:
            ud = decision_map[md.journal_line_id]
            updated.append(MatchDecision(
                journal_line_id=md.journal_line_id,
                counterparty_id=ud.counterparty_id,
                status="confirmed" if ud.counterparty_id else "rejected",
                candidate_ids=md.candidate_ids,
                reason=f"用户确认: {ud.original_text}",
                decision_id=ud.decision_id,
            ))
        else:
            updated.append(md)

    return MatchPlan(
        decisions=updated,
        summary=f"应用了 {len(decisions)} 条用户确认决定",
    )
