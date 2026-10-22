# -*- coding: utf-8 -*-
"""校验并追加招生计划记录 / 来源日志（append-only，写入前自动备份）。

输入：一个 JSON 文件，可为两种形式
  A) 记录数组：            [ {...record...}, {...} ]
  B) 记录 + 来源：          {"records": [...], "sources": [...]}

用法：
    python append_records.py --run ./data/考研招生计划_2027 --input batch_01.json
    python append_records.py --run <dir> --input batch.json --mark-conflict
    python append_records.py --run <dir> --input -            # 从 stdin 读取

退出码：0 正常；1 存在被拒绝的记录；2 参数/环境错误。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from urllib.parse import urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import (  # noqa: E402
    CONFLICT_FIELDS, RECORD_FIELDS, RECORD_HEADER_MAP, SCHOOL_FIELDS,
    SCHOOL_HEADER_MAP, SOURCE_FIELDS, SOURCE_HEADER_MAP, VERIFY_STATUS, VERSIONS,
    append_log, backup_file, now_str, read_rows, to_int, today_str, write_rows,
)

DEDUPE_KEYS = ("year", "school_code", "major_code", "direction", "study_mode",
               "exam_mode", "source_url")
CONFLICT_KEYS = ("year", "school_code", "major_code", "direction", "study_mode",
                 "exam_mode")
QUOTA_FIELDS = ("planned_total", "recommend_quota", "unified_quota")

# 允许作为人数依据的官方域名后缀（学校 .edu.cn / 中科院系统 .ac.cn / 政府 .gov.cn / 研招网）
OFFICIAL_SUFFIXES = (".edu.cn", ".ac.cn", ".gov.cn", "chsi.com.cn", ".edu", ".mil.cn")

FIELD_LABEL = {k: h for k, h, _w in RECORD_FIELDS}


def is_official(url: str) -> bool:
    host = (urlparse(url).hostname or "").lower()
    if not host:
        return False
    return any(host == s.lstrip(".") or host.endswith(s) for s in OFFICIAL_SUFFIXES)


def norm(value) -> str:
    """文本字段：去首尾空白，内部连续空白折叠为一个空格（保留可读性）。"""
    return re.sub(r"\s+", " ", str(value or "")).strip()


def normc(value) -> str:
    """代码/枚举字段：去掉全部空白，避免“085400 01”与“08540001”被判为不同专业。"""
    return re.sub(r"\s+", "", str(value or ""))


CODE_KEYS = {"year", "school_code", "college_code", "major_code", "study_mode",
             "exam_mode", "verify_status", "is_current", "source_url", "publish_date",
             "planned_total", "recommend_quota", "unified_quota"}


def validate_record(rec: dict, year_default: int, allow_nonofficial: bool = False):
    """返回 (清洗后记录, 错误列表, 警告列表)。"""
    errors, warnings = [], []
    out = {}

    for k, _h, _w in RECORD_FIELDS:
        out[k] = normc(rec.get(k, "")) if k in CODE_KEYS else norm(rec.get(k, ""))

    # 年份
    if not out["year"]:
        out["year"] = str(year_default)
        warnings.append("缺少招生年份，已按任务年份补齐")
    if not re.fullmatch(r"20\d{2}", out["year"]):
        errors.append(f"招生年份不合法：{out['year']}")

    # 必填项
    for k in ("school", "major", "source_title", "source_url"):
        if not out[k]:
            errors.append(f"缺少必填字段：{FIELD_LABEL[k]}")

    # 来源必须是 http(s) 原始地址
    if out["source_url"] and not re.match(r"^https?://", out["source_url"]):
        errors.append(f"来源网址不是可追溯的 http(s) 原始链接：{out['source_url']}")
    nonofficial = False
    if out["source_url"] and re.match(r"^https?://", out["source_url"]):
        if not is_official(out["source_url"]):
            nonofficial = True
            if not allow_nonofficial:
                errors.append(
                    f"来源域名非官方（需 .edu.cn / .ac.cn / .gov.cn / chsi.com.cn）："
                    f"{out['source_url']} —— 此类材料只能用于发现线索，不得作为人数依据；"
                    f"确需保留请加 --allow-nonofficial（将强制标为待核验）"
                )
            else:
                warnings.append("来源域名非官方，仅作线索保留，已强制标为“待核验”")

    # 代码格式
    if out["school_code"] and not re.fullmatch(r"\d{4,5}", out["school_code"]):
        warnings.append(f"学校代码非 4-5 位数字：{out['school_code']}（需回原文核对）")
    if out["major_code"] and not re.fullmatch(r"\d{6}(\d{0,2})?", out["major_code"]):
        warnings.append(f"专业代码非 6 位（或 6+2 方向码）：{out['major_code']}（需回原文核对）")

    # 人数解析
    for k in QUOTA_FIELDS:
        raw = out[k]
        if raw == "":
            continue
        val, text = to_int(raw)
        if val is None:
            warnings.append(f"{FIELD_LABEL[k]}无法解析为数字（原文：{text}），已置空并标待核验")
            out[k] = ""
        else:
            if not re.fullmatch(r"\d+", text):
                warnings.append(f"{FIELD_LABEL[k]}原文含附加说明（{text}），已取数值 {val}，请在“人数口径说明”写清口径")
            out[k] = str(val)

    # 口径与推算检查（核心红线）
    p, r, u = (out["planned_total"], out["recommend_quota"], out["unified_quota"])
    if not out["quota_basis"] and (p or r or u):
        warnings.append("缺少“人数口径说明”：必须抄录原文对人数范围/口径的表述，否则不得认定人数")

    if u and p and r and int(u) != int(p) - int(r):
        warnings.append(
            f"数值关系不自洽：统考 {u} ≠ 拟招生 {p} − 推免 {r}"
            f"（{FIELD_LABEL['unified_quota']}与其余两项可能不是同一口径，禁止自行修正，需回原文核对）"
        )
    if u and not p:
        warnings.append("只填了统考计划人数而未填拟招生人数：请确认原文是否单列统考名额")
    if p and not u and r:
        warnings.append("仅有拟招生与推免人数：统考名额留空，禁止用减法推算（除非原文明确写“统考拟招 X 人”）")

    # 核验状态与版本
    if out["verify_status"] not in VERIFY_STATUS:
        out["verify_status"] = "待核验"
        warnings.append("核验状态缺失或不在枚举内，已置为“待核验”")
    if out["is_current"] not in VERSIONS:
        out["is_current"] = "当前采用"
    if out["verify_status"] == "已核验" and not out["quota_basis"] and (p or r or u):
        out["verify_status"] = "待核验"
        warnings.append("人数缺口径说明，不能标为“已核验”，已置为“待核验”")
    if nonofficial:
        out["verify_status"] = "待核验"
    if not out["publish_date"]:
        if out["verify_status"] == "已核验":
            out["verify_status"] = "待核验"
        warnings.append("缺少发布日期，已降级为“待核验”")

    out["collected_at"] = out["collected_at"] or now_str()
    return out, errors, warnings


def detect_conflicts(new_rows, old_rows):
    """同校同专业同方向同学制下，不同来源给出的数值不一致 → 冲突条目。"""
    def key(rec):
        return tuple(norm(rec.get(k)) for k in CONFLICT_KEYS)

    pool = {}
    for rec in list(old_rows) + list(new_rows):
        if rec.get("is_current") == "历史版本":
            continue
        pool.setdefault(key(rec), []).append(rec)

    conflicts = []
    for _k, group in pool.items():
        urls = {}
        for rec in group:
            urls.setdefault(rec.get("source_url", ""), rec)
        if len(urls) < 2:
            continue
        url_items = list(urls.items())
        for i in range(len(url_items)):
            for j in range(i + 1, len(url_items)):
                (ua, ra), (ub, rb) = url_items[i], url_items[j]
                for f in QUOTA_FIELDS:
                    va, vb = ra.get(f, ""), rb.get(f, "")
                    if va != "" and vb != "" and va != vb:
                        conflicts.append({
                            "detected_at": now_str(),
                            "year": ra.get("year", ""),
                            "school": ra.get("school", ""),
                            "school_code": ra.get("school_code", ""),
                            "major": ra.get("major", ""),
                            "major_code": ra.get("major_code", ""),
                            "direction": ra.get("direction", ""),
                            "study_mode": ra.get("study_mode", ""),
                            "field": FIELD_LABEL.get(f, f),
                            "value_a": va, "source_a": ra.get("source_title", "") + "｜" + ua,
                            "value_b": vb, "source_b": rb.get("source_title", "") + "｜" + ub,
                            "resolution": "保留两条原始记录并标注差异，待人工确认采用版本",
                        })
    return conflicts


def main():
    ap = argparse.ArgumentParser(description="校验并追加招生计划记录")
    ap.add_argument("--run", required=True, help="任务根目录（含 data/records.csv）")
    ap.add_argument("--input", required=True, help="JSON 文件路径，或 - 表示 stdin")
    ap.add_argument("--mark-conflict", action="store_true",
                    help="冲突时自动把相关记录核验状态改为“口径冲突”")
    ap.add_argument("--batch", default="", help="批次名，写入日志，如 trial5 / 江苏批次")
    ap.add_argument("--allow-nonofficial", action="store_true",
                    help="允许非官方域名的记录入库，但强制标为“待核验”（仅用于保留线索）")
    args = ap.parse_args()

    run = Path(args.run).expanduser().resolve()
    records_path = run / "data" / "records.csv"
    if not records_path.exists():
        print(f"[错误] 未找到 {records_path}，请先运行 init_run.py", file=sys.stderr)
        sys.exit(2)

    if args.input == "-":
        payload = json.load(sys.stdin)
    else:
        p = Path(args.input).expanduser()
        if not p.is_absolute():
            p = (Path.cwd() / p).resolve()
        payload = json.loads(p.read_text(encoding="utf-8-sig"))

    if isinstance(payload, dict):
        in_records = payload.get("records", [])
        in_sources = payload.get("sources", [])
    else:
        in_records, in_sources = payload, []

    cfg = {}
    cfg_path = run / "run.json"
    if cfg_path.exists():
        cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
    year_default = int(cfg.get("year") or today_str()[:4])

    old = read_rows(records_path, RECORD_HEADER_MAP)
    next_seq = 1
    for rec in old:
        rid = rec.get("record_id", "")
        m = re.fullmatch(r"R(\d+)", rid or "")
        if m:
            next_seq = max(next_seq, int(m.group(1)) + 1)

    seen = {tuple(norm(r.get(k)) for k in DEDUPE_KEYS) for r in old}
    accepted, rejected, warnings, duplicates = [], [], [], []

    for idx, raw in enumerate(in_records, 1):
        rec, errs, warns = validate_record(raw, year_default, args.allow_nonofficial)
        tag = f"第{idx}条（{rec.get('school') or '?'} / {rec.get('major_code') or rec.get('major') or '?'}）"
        if errs:
            rejected.append({"item": tag, "errors": errs})
            continue
        dkey = tuple(norm(rec.get(k)) for k in DEDUPE_KEYS)
        if dkey in seen:
            duplicates.append(tag)
            continue
        seen.add(dkey)
        rec["record_id"] = f"R{next_seq:05d}"
        next_seq += 1
        accepted.append(rec)
        if warns:
            warnings.append({"item": tag, "warnings": warns})

    # 冲突检测
    conflicts = detect_conflicts(accepted, old) if accepted else []
    if conflicts and args.mark_conflict:
        hit = {(c["school_code"], c["major_code"], c["direction"], c["study_mode"]) for c in conflicts}
        for rec in accepted:
            k = (norm(rec.get("school_code")), norm(rec.get("major_code")),
                 norm(rec.get("direction")), norm(rec.get("study_mode")))
            if k in hit and rec["verify_status"] == "已核验":
                rec["verify_status"] = "口径冲突"

    # 写入（先备份）
    if accepted:
        backup_file(records_path, run)
        write_rows(records_path, RECORD_FIELDS, old + accepted)

    sources_path = run / "data" / "sources.csv"
    if in_sources:
        old_src = read_rows(sources_path, SOURCE_HEADER_MAP)
        backup_file(sources_path, run)
        new_src = []
        for s in in_sources:
            row = {k: norm(s.get(k, "")) for k, _h, _w in SOURCE_FIELDS}
            row["collected_at"] = row["collected_at"] or now_str()
            if not row["url"]:
                continue
            new_src.append(row)
        write_rows(sources_path, SOURCE_FIELDS, old_src + new_src)

    if conflicts:
        cpath = run / "data" / "conflicts.csv"
        old_c = read_rows(cpath, {h: k for k, h, _w in CONFLICT_FIELDS})
        backup_file(cpath, run)
        write_rows(cpath, CONFLICT_FIELDS, old_c + conflicts)

    # 同步学校清单的记录数与状态
    schools_path = run / "data" / "schools.csv"
    if schools_path.exists() and accepted:
        schools = read_rows(schools_path, SCHOOL_HEADER_MAP)
        all_recs = old + accepted
        for s in schools:
            if not s.get("school"):
                continue
            n = sum(1 for r in all_recs if r.get("school") == s["school"])
            if n:
                s["records_count"] = str(n)
                # 有记录 ≠ 已完成：仅置为“部分完成”，由人工确认全量覆盖后再改“已完成”
                if s.get("status") in ("待检索", "检索中", ""):
                    s["status"] = "部分完成"
                s["updated_at"] = now_str()
        write_rows(schools_path, SCHOOL_FIELDS, schools)

    append_log(
        run,
        f"批次「{args.batch or '未命名'}」追加记录 {len(accepted)} 条、来源 {len(in_sources)} 条、"
        f"冲突 {len(conflicts)} 条；拒绝 {len(rejected)} 条；重复跳过 {len(duplicates)} 条",
    )

    summary = {
        "accepted": len(accepted),
        "rejected": len(rejected),
        "duplicated": len(duplicates),
        "conflicts": len(conflicts),
        "warnings": warnings,
        "reject_detail": rejected,
        "duplicate_detail": duplicates,
        "total_records": len(old) + len(accepted),
        "records_file": str(records_path),
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    sys.exit(1 if rejected else 0)


if __name__ == "__main__":
    main()
