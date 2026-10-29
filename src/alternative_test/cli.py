"""实际运行探测、准备、生成和验证，不使用占位成功状态。"""

from __future__ import annotations

import json
import re
import sys
import uuid
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from .adapters import ColumnMapping, read_rows
from .amounts import parse_amount
from .controls import reconcile_counterparty
from .errors import OutputConflictError
from .fileio import calc_sha256, probe_file
from .intake import classify_tables
from .matching import UserDecision, _make_line_id, apply_confirmed_decisions, propose_matches
from .models import CounterpartyBalance, JournalLine, NormalBalance, SourceRef
from .population import ConfirmedScope, PopulationResult, build_population
from .renderer import RenderContext, post_payment_evidence, render_workbook, validate_workbook
from .selection import SelectionRule, select_candidates

EXIT_SUCCESS = 0
EXIT_NEEDS_ANSWER = 2
EXIT_INPUT_UNUSABLE = 3
EXIT_VALIDATION_ERROR = 4
EXIT_OUTPUT_CONFLICT = 5
EXIT_INTERNAL = 6


def _json_out(data, code=0):
    print(json.dumps(data, ensure_ascii=False, default=str))
    raise SystemExit(code)


def _write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(data, stream, ensure_ascii=False, indent=2, default=str)


def _read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def _date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"日期无法读取：{value!r}")


def cmd_inspect(files, output=None):
    results, failures = {}, []
    for file in files:
        try:
            probe = probe_file(Path(file))
            scores = classify_tables(probe.candidates)
            results[file] = {
                "path": str(probe.path.resolve()), "format": probe.format,
                "sha256": probe.sha256, "size": probe.size,
                "candidates": [{"name": c.name, "headers": c.headers, "row_count": c.row_count,
                    "has_formulas": c.has_formulas, "has_merged_cells": c.has_merged_cells,
                    "role": s.role, "confidence": s.confidence} for c, s in zip(probe.candidates, scores)],
                "warnings": probe.warnings,
            }
        except Exception as exc:
            failures.append(f"{file}：{exc}")
            results[file] = {"error": str(exc)}
    code = EXIT_INPUT_UNUSABLE if failures else 0
    data = {"files": results, "exit_code": code}
    if output:
        try:
            _write_json(Path(output), data)
        except FileExistsError:
            _json_out({"error": "探测输出已经存在，未覆盖"}, EXIT_OUTPUT_CONFLICT)
        _json_out({"status": "error" if failures else "ok", "output": output, "errors": failures}, code)
    _json_out(data, code)


def _entity(row, spec, answers):
    raw = row.get("公司")
    explicit = spec.get("entity_id")
    entities = answers["entity_ids"]
    names = answers["entity_names"]
    inverse = {v: k for k, v in names.items()}
    if raw not in (None, ""):
        actual = str(raw).strip()
        actual = inverse.get(actual, actual)
        if explicit and actual != explicit:
            raise ValueError(f"资料单位{raw}与配置单位{explicit}不一致")
        entity = actual
    else:
        entity = explicit or (entities[0] if len(entities) == 1 else None)
    if entity not in entities:
        raise ValueError(f"无法唯一确定资料所属单位：{raw!r}")
    return entity


def _prepare_inputs(inspect, answers):
    balances, lines, hashes = [], [], {}
    inputs = answers.get("inputs", [])
    if not inputs or not {"balance", "journal"}.issubset({s.get("role") for s in inputs}):
        raise ValueError("缺少余额表或序时账的输入配置")
    known = {str(Path(k).resolve()): v for k, v in inspect["files"].items()}
    seen_balances, seen_lines = set(), set()
    for spec in inputs:
        path = Path(spec["path"]).resolve()
        info = known.get(str(path))
        if not info or info.get("error"):
            raise ValueError(f"输入未经成功探测：{path}")
        actual_hash = calc_sha256(path)
        if actual_hash != info["sha256"]:
            raise ValueError(f"探测后输入发生变化：{path}，重新探测即可")
        hashes[str(path)] = actual_hash
        mapping = ColumnMapping(**spec["columns"])
        role = spec["role"]
        required = ("counterparty_col", "ending_balance_col") if role == "balance" else ("date_col", "voucher_no_col", "debit_col", "credit_col")
        if any(getattr(mapping, key) is None for key in required):
            raise ValueError(f"{path.name}缺少必需列映射：{required}")
        candidates = info["candidates"]
        sheet = spec.get("sheet")
        if not sheet:
            if len(candidates) != 1:
                raise ValueError(f"{path.name}有多个工作表，需要确定本次使用的表")
            sheet = candidates[0]["name"]
        if sheet not in [c["name"] for c in candidates]:
            raise ValueError(f"工作表不存在：{path.name}/{sheet}")
        actual_sheet = None if info["format"] == "csv" else sheet
        header_row = int(spec.get("header_row", 1))
        if header_row < 1:
            raise ValueError("表头行号必须大于0")
        for row in read_rows(path, mapping, actual_sheet, header_row):
            if not any(v not in (None, "") for k, v in row.items() if k != "_source_row"):
                continue
            source = SourceRef(file_id=str(path), sha256=actual_hash, sheet=sheet,
                row=row["_source_row"], columns={k: v for k, v in spec["columns"].items() if v})
            entity = _entity(row, spec, answers)
            account = spec.get("account_name") or row.get("科目")
            if not account:
                raise ValueError(f"{path.name}第{source.row}行缺少测试科目")
            def amount(key):
                value = parse_amount(row.get(key), source=source, field=key)
                if value is not None and not value.is_finite():
                    raise ValueError(f"金额必须是有限数值：{path.name}第{source.row}行/{key}")
                return value
            if role == "balance":
                name, ending = row.get("供应商"), amount("期末余额")
                if not name or ending is None:
                    raise ValueError(f"余额表第{source.row}行缺少往来单位或期末余额")
                key = (entity, str(account), str(name))
                if key in seen_balances:
                    raise ValueError(f"余额清单重复：{key}，需先明确汇总口径")
                seen_balances.add(key)
                balances.append(CounterpartyBalance(entity_id=entity, counterparty_id=f"CP{len(balances)+1:04d}",
                    raw_name=str(name), account_code=str(row["科目代码"]) if row.get("科目代码") else None,
                    account_name=str(account), normal_balance=NormalBalance(answers["normal_balance"]),
                    opening_balance=amount("期初余额"), ending_balance=ending, source=source))
            elif role == "journal":
                ledger = row.get("账套") or spec.get("ledger_id")
                if not ledger or row.get("日期") in (None, "") or row.get("凭证号") in (None, ""):
                    raise ValueError(f"序时账第{source.row}行缺少账套、日期或凭证号")
                debit, credit = amount("借方"), amount("贷方")
                if debit is None and credit is None:
                    raise ValueError(f"序时账第{source.row}行借贷金额均为空")
                line = JournalLine(entity_id=entity, ledger_id=str(ledger), posting_date=_date(row["日期"]),
                    voucher_type=str(row.get("凭证类型") or ""), voucher_number=str(row["凭证号"]),
                    line_number=int(spec.get("line_offset", 0))+source.row,
                    account_code=str(row["科目代码"]) if row.get("科目代码") else None,
                    account_name=str(account), counterparty_raw=str(row["辅助核算"]) if row.get("辅助核算") else None,
                    summary=str(row.get("摘要") or ""), opposing_account_name=str(row["对方科目"]) if row.get("对方科目") else None,
                    debit=debit or Decimal(0), credit=credit or Decimal(0), source=source)
                identity = _make_line_id(line)
                if identity in seen_lines:
                    raise ValueError(f"序时账行标识重复：{identity}；重叠账簿不能重复计入，分表行号可设置line_offset")
                seen_lines.add(identity)
                lines.append(line)
            else:
                raise ValueError(f"不支持的输入角色：{role}")
    if not balances or not lines:
        raise ValueError("余额表或序时账没有有效数据")
    return balances, lines, hashes


def _build_context(prepared):
    answers = prepared["answers"]
    balances = [CounterpartyBalance.model_validate(x) for x in prepared["balances"]]
    lines = [JournalLine.model_validate(x) for x in prepared["journal_lines"]]
    direction = NormalBalance(answers["normal_balance"])
    scope = ConfirmedScope(start_date=_date(answers["period_start"]), end_date=_date(answers["period_end"]),
        entity_ids=answers["entity_ids"], ledger_ids=answers.get("ledger_ids", []))
    scoped_lines = [l for l in lines if l.entity_id in scope.entity_ids
                    and (not scope.ledger_ids or l.ledger_id in scope.ledger_ids)]
    # 从同一整张凭证的反方向分录补出对方科目，有多个科目时如实并列，保留来源。
    voucher_groups = {}
    for line in scoped_lines:
        voucher_groups.setdefault(line.voucher_key, []).append(line)
    for line in scoped_lines:
        if line.opposing_account_name:
            continue
        other = [l for l in voucher_groups[line.voucher_key] if l.account_name != line.account_name
                 and ((line.debit != 0 and l.credit != 0) or (line.credit != 0 and l.debit != 0))]
        if other:
            line.opposing_account_name = "、".join(dict.fromkeys(l.account_name for l in other))
            line.opposing_account_sources = [l.source for l in other]
    target_accounts = {(b.entity_id, b.account_name) for b in balances}
    target_lines = [l for l in scoped_lines if (l.entity_id, l.account_name) in target_accounts]
    plan = propose_matches(target_lines, balances)
    if answers.get("confirmed_decisions"):
        decisions = [UserDecision(**d) for d in answers["confirmed_decisions"]]
        valid_ids = {_make_line_id(l) for l in scoped_lines}
        for decision in decisions:
            line = next((l for l in scoped_lines if _make_line_id(l) == decision.journal_line_id), None)
            balance = next((b for b in balances if b.counterparty_id == decision.counterparty_id), None)
            if decision.journal_line_id not in valid_ids or (decision.counterparty_id and
                    (not balance or balance.entity_id != line.entity_id or balance.account_name != line.account_name)):
                raise ValueError("确认映射不能跨单位、跨科目或指向不存在的凭证")
        plan = apply_confirmed_decisions(plan, decisions)
    population = build_population(target_lines, scope, plan, direction)
    approved = {d.journal_line_id for d in plan.decisions if d.status in ("exact_unique", "confirmed")}
    eligible = [l for l in population.included if _make_line_id(l) in approved]
    controls = []
    for balance in balances:
        cpids = {d.journal_line_id for d in plan.decisions if d.counterparty_id == balance.counterparty_id and d.status in ("exact_unique", "confirmed")}
        cp = [l for l in eligible if _make_line_id(l) in cpids]
        cp_population = PopulationResult(included=cp, total_debit=sum((l.debit for l in cp), Decimal(0)),
                                       total_credit=sum((l.credit for l in cp), Decimal(0)))
        controls.append(reconcile_counterparty(balance, cp_population))
    post_end = _date(answers["post_period_end"]) if answers.get("post_period_end") else max((l.posting_date for l in scoped_lines), default=scope.end_date)
    side = "liability" if direction == NormalBalance.CREDIT else "asset"
    post = [l for l in target_lines if scope.end_date < l.posting_date <= post_end and _make_line_id(l) in approved
            and post_payment_evidence(l, side)[0]]
    rules = [SelectionRule.model_validate(r) for r in answers.get("rules", [])]
    selection = select_candidates(PopulationResult(included=eligible), rules, direction)
    mode = "selected" if any(r.user_confirmed for r in rules) else "population"
    if "selected_line_ids" in answers:
        ids = set(answers["selected_line_ids"])
        if not ids.issubset({_make_line_id(l) for l in eligible}):
            raise ValueError("指定抽取凭证中含不存在、未归户或非当期的行")
        selection.selected = [(l, "明确指定") for l in eligible if _make_line_id(l) in ids]
        mode = "selected"
    # 期后单独检索，不被当期抽样规则误删。未给期后规则时保留全部期后候选。
    post_rules = [SelectionRule.model_validate(r) for r in answers.get("post_rules", [])]
    post_selection = select_candidates(PopulationResult(included=post), post_rules, direction)
    post_selected = post_selection.selected if any(r.user_confirmed for r in post_rules) else [(l, "期后候选") for l in post]
    if mode == "selected":
        selection.selected.extend(post_selected)
    ctx = RenderContext(project_name=answers["project_name"],
        period=f"{scope.start_date.isoformat()} ~ {scope.end_date.isoformat()}",
        entity_ids=scope.entity_ids, side=side,
        balances=balances, journal_lines=population.included, match_decisions=plan.decisions,
        controls=controls, selection=selection, population=population, selection_mode=mode,
        excluded_lines=[(l, "期后") for l in post], run_id=prepared["run_id"], input_hashes=prepared["input_hashes"],
        entity_names=answers["entity_names"], period_start=scope.start_date, period_end=scope.end_date,
        post_period_end=max(scope.end_date, post_end), theme=answers.get("theme", "deep-navy"),
        prepared_by=answers.get("prepared_by", ""), reviewed_by=answers.get("reviewed_by", ""),
        prepared_date=_date(answers["prepared_date"]) if answers.get("prepared_date") else None,
        reviewed_date=_date(answers["reviewed_date"]) if answers.get("reviewed_date") else None,
        index_number=answers.get("index_number", ""), index_numbers=answers.get("index_numbers", {}),
        currency=answers.get("currency", "人民币"))
    return ctx, plan


def cmd_prepare(inspect_file, answers_file, output_dir):
    try:
        inspect, answers = _read_json(inspect_file), _read_json(answers_file)
        if "period_end" not in answers and answers.get("test_date"):
            answers["period_end"] = answers["test_date"]
        if answers.get("period_end") and not answers.get("period_start"):
            answers["period_start"] = date(_date(answers["period_end"]).year, 1, 1).isoformat()
            answers.setdefault("assumptions", []).append("未指定期间起点，按测试时点所在自然年度年初累计；有财年口径时应按资料修改")
        for key in ("project_name", "period_start", "period_end", "entity_ids", "entity_names", "normal_balance"):
            if not answers.get(key):
                raise ValueError(f"缺少必要信息：{key}；先从资料和上下文自动补齐，确实不能确定再询问")
        if _date(answers["period_start"]) > _date(answers["period_end"]):
            raise ValueError("审计期间起始日晚于截止日")
        if answers.get("post_period_end") and _date(answers["post_period_end"]) < _date(answers["period_end"]):
            raise ValueError("期后查询截止日早于测试时点")
        if any(not answers["entity_names"].get(e) for e in answers["entity_ids"]):
            raise ValueError("单位名称缺失，不能用内部编号冒充实际名称")
        balances, lines, hashes = _prepare_inputs(inspect, answers)
        prepared = {"schema": "alternative-test-prepared-v2", "run_id": uuid.uuid4().hex,
            "answers": answers, "balances": [b.model_dump(mode="json") for b in balances],
            "journal_lines": [l.model_dump(mode="json") for l in lines], "input_hashes": hashes}
        ctx, plan = _build_context(prepared)
        output = Path(output_dir)
        if output.exists() and any(output.iterdir()):
            raise OutputConflictError("准备目录已有文件，未覆盖")
        _write_json(output / "prepared.json", prepared)
        _write_json(output / "matching.json", [d.model_dump(mode="json") for d in plan.decisions])
        _write_json(output / "controls.json", [c.__dict__ for c in ctx.controls])
        _write_json(output / "excluded.json", [{"line": l.model_dump(mode="json"), "reason": reason} for l, reason in ctx.population.excluded])
        _write_json(output / "post-query.json", [{"line": l.model_dump(mode="json"),
            "eligible": post_payment_evidence(l, ctx.side)[0], "reason": post_payment_evidence(l, ctx.side)[1]}
            for l, _ in ctx.population.excluded if ctx.period_end < l.posting_date <= ctx.post_period_end])
        unresolved = sum(d.status == "needs_confirmation" for d in plan.decisions)
        _json_out({"status": "ok", "prepared_dir": str(output), "population_rows": len(ctx.population.included),
            "selected_rows": len(ctx.selection.selected), "post_query_end": ctx.post_period_end,
            "unresolved_matches": unresolved, "unmatched_rows": sum(d.status == "unmatched" for d in plan.decisions),
            "assumptions": answers.get("assumptions", [])})
    except OutputConflictError as exc:
        _json_out({"error": str(exc)}, EXIT_OUTPUT_CONFLICT)
    except Exception as exc:
        _json_out({"error": str(exc)}, EXIT_INPUT_UNUSABLE)


def cmd_generate(prepared_dir, output_parent, template=None):
    try:
        prepared_path = Path(prepared_dir) / "prepared.json"
        prepared = _read_json(prepared_path)
        for file, digest in prepared["input_hashes"].items():
            if calc_sha256(Path(file)) != digest:
                raise ValueError(f"准备后输入发生变化：{file}，重新准备即可")
        ctx, plan = _build_context(prepared)
        unresolved = [d.model_dump(mode="json") for d in plan.decisions if d.status == "needs_confirmation"]
        if unresolved:
            _json_out({"error": "存在无法唯一归户的候选；AI应先结合辅助核算和上下文自行核实",
                       "unresolved_matches": unresolved}, EXIT_NEEDS_ANSWER)
        root = Path(__file__).resolve().parents[2]
        filename = "替代测试-负债方-v1.xlsx" if ctx.side == "liability" else "替代测试-资产方-v1.xlsx"
        template_path = Path(template).resolve() if template else root / "assets/templates" / filename
        output_dir = Path(output_parent)
        manifest_path = output_dir / "run-manifest.json"
        report_path = output_dir / "validation.json"
        if manifest_path.exists() or report_path.exists():
            raise OutputConflictError("生成目录已有运行记录，未覆盖")
        name = re.sub(r'[\\/:*?"<>|]', "_", ctx.project_name).strip().rstrip(".") or "往来"
        output = output_dir / f"{name[:80]}_{ctx.period_end.isoformat()}_往来替代测试候选底稿.xlsx"
        render_workbook(ctx, template_path, output)
        report = validate_workbook(output)
        _write_json(report_path, report)
        _write_json(manifest_path, {"schema": "alternative-test-run-v2", "run_id": ctx.run_id,
            "workbook": output.name, "sha256": calc_sha256(output), "prepared_sha256": calc_sha256(prepared_path),
            "template": str(template_path), "template_sha256": calc_sha256(template_path),
            "input_hashes": ctx.input_hashes, "post_query_end": ctx.post_period_end})
        _json_out({"status": "ok", "output": str(output), "validation": report,
                   "post_query_end": ctx.post_period_end}, 0 if report["valid"] else EXIT_VALIDATION_ERROR)
    except (OutputConflictError, FileExistsError) as exc:
        _json_out({"error": str(exc)}, EXIT_OUTPUT_CONFLICT)
    except Exception as exc:
        _json_out({"error": str(exc)}, EXIT_INPUT_UNUSABLE)


def cmd_validate(run_dir):
    try:
        path = Path(run_dir)
        if path.is_file():
            report = validate_workbook(path)
        else:
            manifest = _read_json(path / "run-manifest.json")
            workbook = (path / manifest["workbook"]).resolve()
            if workbook.parent != path.resolve():
                raise ValueError("运行记录中的工作簿路径超出本目录")
            report = validate_workbook(workbook)
            if calc_sha256(workbook) != manifest["sha256"]:
                report["errors"].append("成品在生成后发生变化，文件哈希不一致")
                report["valid"] = False
        _json_out({"status": "ok" if report["valid"] else "error", **report}, 0 if report["valid"] else EXIT_VALIDATION_ERROR)
    except Exception as exc:
        _json_out({"status": "error", "valid": False, "errors": [str(exc)]}, EXIT_VALIDATION_ERROR)
