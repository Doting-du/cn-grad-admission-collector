# -*- coding: utf-8 -*-
"""初始化一次采集任务的目录骨架与学校清单。

用法示例：
    python init_run.py --year 2027 --provinces 北京,上海 --out ./data/考研招生计划_2027
    python init_run.py --year 2027 --schools schools.txt --out ./data/run_2027
    python init_run.py --year 2027 --provinces 江苏 --schools extra.txt --slug js

schools.txt 每行：学校名称[,学校代码[-省份]]
    # 以 # 开头的行视为注释
    南京大学,10284
    东南大学,10286-江苏
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import (  # noqa: E402
    CONFLICT_FIELDS, MAINLAND_PROVINCES, RECORD_FIELDS, SCHOOL_FIELDS,
    SCHOOL_STATUS, SOURCE_FIELDS, now_str, read_rows, write_rows,
    SCHOOL_HEADER_MAP, today_str,
)


def parse_school_lines(path: Path, default_province: str):
    items = []
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        parts = [p.strip() for p in line.replace("\t", ",").split(",")]
        name = parts[0]
        code = parts[1] if len(parts) > 1 else ""
        prov = parts[2] if len(parts) > 2 else default_province
        if name:
            items.append({"school": name, "school_code": code, "province": prov})
    return items


def main():
    ap = argparse.ArgumentParser(description="初始化考研招生计划采集任务")
    ap.add_argument("--year", type=int, required=True, help="招生年份，如 2027")
    ap.add_argument("--out", required=True, help="任务根目录")
    ap.add_argument("--provinces", default="", help="省份，逗号分隔，如 北京,上海；留空表示全部31个省级行政区")
    ap.add_argument("--schools", default="", help="学校清单 txt 路径（每行 学校名,学校代码,省份）")
    ap.add_argument("--slug", default="", help="目录后缀标记，如 trial5")
    ap.add_argument("--force", action="store_true", help="目标目录已存在时允许复用（不删除任何已有数据）")
    args = ap.parse_args()

    out = Path(args.out).expanduser().resolve()
    if out.exists() and any(out.iterdir()) and not args.force:
        print(f"[中止] 目录已存在且非空：{out}\n       默认不覆盖已有采集数据；如需复用请加 --force（仅补齐缺失文件，不删除数据）")
        sys.exit(2)

    data_dir = out / "data"
    for sub in (data_dir, data_dir / "history", out / "sources", out / "logs"):
        sub.mkdir(parents=True, exist_ok=True)

    provinces = [p.strip() for p in args.provinces.replace("，", ",").split(",") if p.strip()]
    if not provinces:
        provinces = list(MAINLAND_PROVINCES)
    unknown = [p for p in provinces if p not in MAINLAND_PROVINCES]
    if unknown:
        print(f"[警告] 非大陆省级行政区或名称不规范：{unknown}（请使用标准简称，如 北京/内蒙古/广西）")

    # 1) 主数据表、冲突表、来源日志：仅在缺失时创建空表
    for rel, fields in (
        ("data/records.csv", RECORD_FIELDS),
        ("data/conflicts.csv", CONFLICT_FIELDS),
        ("data/sources.csv", SOURCE_FIELDS),
    ):
        p = out / rel
        if not p.exists():
            write_rows(p, fields, [])
            print(f"[创建] {rel}")

    # 2) 学校清单
    schools_path = data_dir / "schools.csv"
    existing = read_rows(schools_path, SCHOOL_HEADER_MAP)
    seen = {(r.get("school", ""), r.get("province", "")) for r in existing}
    added = 0

    if args.schools:
        sp = Path(args.schools).expanduser()
        if not sp.is_absolute():
            sp = (Path.cwd() / sp).resolve()
        if not sp.exists():
            print(f"[错误] 找不到学校清单文件：{sp}")
            sys.exit(2)
        for item in parse_school_lines(sp, provinces[0] if len(provinces) == 1 else ""):
            key = (item["school"], item["province"])
            if key in seen:
                continue
            existing.append({
                "province": item["province"], "school": item["school"],
                "school_code": item["school_code"], "grad_school_url": "",
                "admissions_url": "", "colleges_found": "", "status": "待检索",
                "records_count": "0", "notes": "", "updated_at": now_str(),
            })
            seen.add(key)
            added += 1
    else:
        # 未提供学校清单时，仅登记省份占位，由后续检索阶段补充学校
        for prov in provinces:
            key = ("", prov)
            if key in seen:
                continue
            existing.append({
                "province": prov, "school": "", "school_code": "",
                "grad_school_url": "", "admissions_url": "", "colleges_found": "",
                "status": "待检索", "records_count": "0",
                "notes": "待按省份补充学校清单", "updated_at": now_str(),
            })
            seen.add(key)
            added += 1

    write_rows(schools_path, SCHOOL_FIELDS, existing)

    # 3) 运行配置
    cfg_path = out / "run.json"
    cfg = {}
    if cfg_path.exists():
        try:
            cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
        except Exception:
            cfg = {}
    cfg.update({
        "year": args.year,
        "provinces": provinces,
        "slug": args.slug,
        "created_at": cfg.get("created_at") or now_str(),
        "updated_at": now_str(),
        "schools_total": len([r for r in existing if r.get("school")]),
        "data_files": {
            "records": "data/records.csv",
            "sources": "data/sources.csv",
            "conflicts": "data/conflicts.csv",
            "schools": "data/schools.csv",
        },
    })
    cfg_path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

    # 4) 采集日志首条
    log = out / "logs" / "collection_log.md"
    if not log.exists():
        log.write_text(
            f"# 采集日志 · {args.year} 年硕士研究生招生计划\n\n"
            f"- 任务创建：{now_str()}\n"
            f"- 年份：{args.year}\n"
            f"- 省份范围：{'、'.join(provinces)}\n\n## 操作记录\n\n",
            encoding="utf-8",
        )
    with log.open("a", encoding="utf-8") as fh:
        fh.write(f"- `{now_str()}` 初始化任务，新增学校 {added} 所，状态枚举：{'/'.join(SCHOOL_STATUS)}\n")

    print(f"[完成] 任务目录：{out}")
    print(f"       学校清单：{schools_path}（本次新增 {added} 行，状态=待检索）")
    print(f"       年份 {args.year}｜日期 {today_str()}")
    print("\n下一步：填写 schools.csv 的 grad_school_url / admissions_url，然后逐校检索原始材料。")


if __name__ == "__main__":
    main()
