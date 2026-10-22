# -*- coding: utf-8 -*-
"""共享常量与工具：字段定义、CSV 读写、备份、日志。

所有脚本共用本模块，保证 records.csv / sources.csv / conflicts.csv / schools.csv
的列顺序、中文表头与读写口径完全一致（用户可直接用 Excel 打开编辑后再被脚本读回）。
"""
from __future__ import annotations

import csv
import shutil
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------- 字段定义
# (key, 中文表头, Excel 列宽)
RECORD_FIELDS = [
    ("year", "招生年份", 10),
    ("province", "省份", 9),
    ("school", "学校名称", 22),
    ("school_code", "学校代码", 10),
    ("college", "学院名称", 24),
    ("college_code", "学院代码", 10),
    ("major", "专业名称", 24),
    ("major_code", "专业代码", 11),
    ("direction", "研究方向", 30),
    ("study_mode", "学习方式", 11),
    ("exam_mode", "考试方式", 12),
    ("planned_total", "拟招生人数", 11),
    ("recommend_quota", "推免人数", 10),
    ("unified_quota", "统考计划人数", 12),
    ("quota_basis", "人数口径说明", 34),
    ("source_title", "原文标题", 40),
    ("source_type", "材料类型", 12),
    ("source_url", "来源网址", 46),
    ("publish_date", "发布日期", 12),
    ("verify_status", "核验状态", 11),
    ("is_current", "版本", 11),
    ("collected_at", "采集时间", 18),
    ("notes", "备注", 30),
    ("record_id", "记录ID", 10),
]

SCHOOL_FIELDS = [
    ("province", "省份", 9),
    ("school", "学校名称", 24),
    ("school_code", "学校代码", 10),
    ("grad_school_url", "研究生院网址", 38),
    ("admissions_url", "研招/专业目录入口", 42),
    ("colleges_found", "已定位学院数", 11),
    ("status", "采集状态", 12),
    ("records_count", "记录条数", 10),
    ("notes", "说明", 34),
    ("updated_at", "更新时间", 18),
]

SOURCE_FIELDS = [
    ("collected_at", "采集时间", 18),
    ("school", "学校名称", 22),
    ("school_code", "学校代码", 10),
    ("source_title", "材料标题", 40),
    ("source_type", "材料类型", 12),
    ("url", "网址", 46),
    ("publish_date", "发布日期", 12),
    ("access_result", "访问结果", 12),
    ("http_status", "HTTP状态", 10),
    ("is_pdf", "是否PDF", 9),
    ("pages_read", "读取页码/篇", 14),
    ("extracted_fields", "已提取字段", 26),
    ("notes", "问题记录", 34),
]

CONFLICT_FIELDS = [
    ("detected_at", "发现时间", 18),
    ("year", "招生年份", 10),
    ("school", "学校名称", 22),
    ("school_code", "学校代码", 10),
    ("major", "专业名称", 22),
    ("major_code", "专业代码", 11),
    ("direction", "研究方向", 24),
    ("study_mode", "学习方式", 11),
    ("field", "冲突字段", 12),
    ("value_a", "来源A数值", 12),
    ("source_a", "来源A", 40),
    ("value_b", "来源B数值", 12),
    ("source_b", "来源B", 40),
    ("resolution", "处理说明", 34),
]

# ---------------------------------------------------------------- 枚举值
STUDY_MODES = ["全日制", "非全日制", "全日制/非全日制", "不区分"]
EXAM_MODES = ["全国统考", "推荐免试", "单独考试", "管理类联考", "法律硕士联考", "不区分"]
SOURCE_TYPES = ["招生简章", "专业目录", "招生计划", "调整公告", "复试录取办法", "其他"]
VERIFY_STATUS = ["已核验", "待核验", "口径冲突", "无数据"]
VERSIONS = ["当前采用", "历史版本"]
SCHOOL_STATUS = ["待检索", "检索中", "已完成", "部分完成", "未完成", "访问受限"]
ACCESS_RESULTS = ["成功", "404/已下线", "需登录", "验证码拦截", "权限限制", "PDF无法解析", "超时/网络失败"]

MAINLAND_PROVINCES = [
    "北京", "天津", "河北", "山西", "内蒙古", "辽宁", "吉林", "黑龙江",
    "上海", "江苏", "浙江", "安徽", "福建", "江西", "山东", "河南",
    "湖北", "湖南", "广东", "广西", "海南", "重庆", "四川", "贵州",
    "云南", "西藏", "陕西", "甘肃", "青海", "宁夏", "新疆",
]

# ---------------------------------------------------------------- 头部映射
def _header_map(fields):
    m = {h: k for k, h, _w in fields}
    for k, h, _w in fields:
        m.setdefault(k, k)
    return m


RECORD_HEADER_MAP = _header_map(RECORD_FIELDS)
SCHOOL_HEADER_MAP = _header_map(SCHOOL_FIELDS)
SOURCE_HEADER_MAP = _header_map(SOURCE_FIELDS)
CONFLICT_HEADER_MAP = _header_map(CONFLICT_FIELDS)


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today_str() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


# ---------------------------------------------------------------- CSV 读写
def read_rows(path, header_map):
    """读取 CSV，按中文表头（或英文 key）映射为 dict。文件不存在返回 []。"""
    path = Path(path)
    if not path.exists():
        return []
    rows = []
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh)
        try:
            raw_header = next(reader)
        except StopIteration:
            return []
        keys = [header_map.get(h.strip(), h.strip()) for h in raw_header]
        for raw in reader:
            if not any((c or "").strip() for c in raw):
                continue
            row = {}
            for i, key in enumerate(keys):
                row[key] = (raw[i] if i < len(raw) else "").strip()
            rows.append(row)
    return rows


def write_rows(path, fields, rows):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    headers = [h for _k, h, _w in fields]
    keys = [k for k, _h, _w in fields]
    tmp = path.with_suffix(path.suffix + ".tmp")
    with tmp.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(headers)
        for row in rows:
            writer.writerow(["" if row.get(k) is None else row.get(k, "") for k in keys])
    tmp.replace(path)


def backup_file(path, run_dir):
    """写入前把旧文件备份到 data/history/，保留修改日志。"""
    path = Path(path)
    if not path.exists():
        return None
    hist = Path(run_dir) / "data" / "history"
    hist.mkdir(parents=True, exist_ok=True)
    dst = hist / f"{path.stem}_{stamp()}{path.suffix}"
    shutil.copy2(path, dst)
    return dst


def append_log(run_dir, message):
    log_dir = Path(run_dir) / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    with (log_dir / "collection_log.md").open("a", encoding="utf-8") as fh:
        fh.write(f"- `{now_str()}` {message}\n")


def to_int(value):
    """把 '30'、'30人'、'30（含推免）' 等解析为 int；无法确定返回 (None, 原文)。"""
    if value is None:
        return None, ""
    if isinstance(value, int):
        return value, str(value)
    text = str(value).strip()
    if text == "" or text in {"-", "—", "/", "无", "不区分", "待核验"}:
        return None, text
    digits = "".join(ch for ch in text if ch.isdigit())
    if digits == "":
        return None, text
    # 只接受纯数字或数字+明确单位，避免把“30-40”误读成 3040
    core = text.replace("人", "").replace("名", "").strip()
    if core.isdigit():
        return int(core), text
    return int(digits), text
