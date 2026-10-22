---
name: cn-grad-admission-collector
description: 采集中国大陆高校硕士研究生招生计划与专业目录中的官方招生人数，逐条回溯到研究生院、研招网、学院官网的原始网页或 PDF，提取学校、学院、专业、研究方向、学习方式、拟招生人数、推免人数与统考计划人数，标注人数口径、核验状态与版本差异，输出可筛选的 Excel 与采集报告，遇验证码或权限限制即停止并登记。当用户要求采集、整理、核对某年考研招生计划、招生人数、专业目录，或核实某校某专业是否招生、招多少人时使用。
version: 0.1.0
category: 数据采集
platforms: [workbuddy]
agent_created: true
---

# 全国高校考研官方招生计划采集

## 用途

按用户指定的招生年份与学校范围，逐校检索官方发布的硕士研究生招生简章、专业目录、招生计划及调整公告，
打开原始网页/PDF 读取内容后提取分专业招生人数，标注口径与核验状态，输出可筛选 Excel 与采集报告。
核心约束：**每一条人数都必须能回溯到官方原文**，无法确认的一律标「待核验」，绝不猜测或推算。

## 三条红线（任何情况下不得突破）

1. **人数只来自官方原文**。搜索引擎摘要、论坛、考研机构、公众号、往年数据仅用于发现线索，
   不得作为当年招生人数的依据。非官方域名的记录会被脚本直接拒收。
2. **禁止推算**。不得用「拟招生 − 推免」得出统考名额，不得按比例或往年数据倒推；
   原文没写的字段留空并标「无数据」或「待核验」。
3. **不绕过访问限制**。遇到验证码、登录墙、付费墙、校内 IP 限制，立即停止，
   在 `sources.csv` 登记 `access_result` 并说明，不得改 UA、伪造 Cookie、调用打码服务或构造接口参数。

## 环境与路径

技能目录即本 `SKILL.md` 所在目录，默认安装位置：

- 用户级：`~/.workbuddy/skills/cn-grad-admission-collector`（Windows 为 `%USERPROFILE%\.workbuddy\skills\cn-grad-admission-collector`）
- 项目级：`<项目>/.workbuddy/skills/cn-grad-admission-collector`

执行前先定位技能目录并安装依赖（仅首次）：

```bash
SKILL="$HOME/.workbuddy/skills/cn-grad-admission-collector"   # 按实际安装位置替换
python -m pip install -r "$SKILL/requirements.txt"
```

- 需要 Python 3.9+；Windows 上 `python` 不可用时改用 `py -3`。
- 唯一第三方依赖是 `openpyxl`（导出 Excel 用），不联网、不写注册表、不申请额外权限。
- 下文命令写作 `python "$SKILL/scripts/xxx.py"`，把 `$SKILL` 换成实际目录即可；也可先 `cd "$SKILL"` 再用相对路径 `python scripts/xxx.py`。

## 标准流程

### Stage 0 · 确认参数

必须明确四项，缺项时按括号内默认值补齐，并在回复中说明所做假设：

| 参数 | 说明 |
| --- | --- |
| 招生年份 | 入学年份。未指定时：9 月及以后取「当年+1」，否则取当年 |
| 学校范围 | 省份 / 学校名单 / 数量上限。清单须有依据（研招网招生单位列表、省考试院公告），不得编造学校代码 |
| 输出目录 | 默认 `./考研招生计划_<年份>` |
| 是否试运行 | **默认先跑 5 所学校**，验收通过再扩量；用户明确要求直接全量时才跳过 |

### Stage 1 · 建任务与清单

把学校清单写成 txt（每行 `学校名称,学校代码,省份`，`#` 开头为注释），然后初始化：

```bash
python "$SKILL/scripts/init_run.py" --year 2027 --schools schools.txt --out ./考研招生计划_2027
```

只给省份未给名单时，先查该省招生单位列表填入 txt，再初始化。脚本不覆盖已有数据；
目标目录非空须显式加 `--force`，且只补齐缺失文件、不删除任何记录。

### Stage 2 · 逐校来源发现（核心工作量）

按 `references/collection-sop.md` 的单校九步执行，按 `references/source-rules.md` 判定来源级别与人数口径。要点：

- 先找 9 月的招生简章与专业目录（PDF 优先），再查 10 月调整公告、次年 3 月复试细则；
- 必须打开原文读取：网页直接抓取，PDF 下载后用 `Read` 逐页读取并记录页码，**不得只看搜索摘要填人数**；
- 每个访问（含失败）都登记到 `sources.csv`。

### Stage 3 · 校验入库

把本批记录写成 JSON（结构见 `references/field-schema.md`，可同时携带 `sources`），然后：

```bash
python "$SKILL/scripts/append_records.py" --run ./考研招生计划_2027 --input batch_01.json --batch trial5
```

脚本执行：必填项与代码格式校验、官方域名校验、人数解析与口径检查、重复去重、
跨来源差异检测、写入前自动备份到 `data/history/`。必须逐条处理返回的 `warnings`
（尤其「数值关系不自洽」「缺人数口径说明」），被拒记录修正后重新提交。
需要自动把冲突记录标为「口径冲突」时加 `--mark-conflict`。

### Stage 4 · 试运行验收

导出（见 Stage 6）后请用户核对 5 项：来源可点开、数字对得上、口径说得清、缺口标得出、字段够用。
**用户确认字段与准确性后**再进入 Stage 5。字段需增删时改 `scripts/common.py` 的 `RECORD_FIELDS`，
已采集数据按字段 key 自动对齐，无须重录。

### Stage 5 · 按省扩量

一省一批，每批入库一次（`--batch 江苏批次`），阶段性导出便于抽查。
单校完成判据见 `references/collection-sop.md`；`schools.csv` 的「已完成」须人工确认后填写。

### Stage 6 · 导出 Excel 与采集报告

```bash
python "$SKILL/scripts/export_excel.py" --run ./考研招生计划_2027
```

生成 `考研招生计划_<年份>.xlsx` 与 `采集报告_<年份>.md`（同名已存在时自动加时间戳，不覆盖）。
用 `present_files` 同时交付 xlsx 与 md，并在回复中给出关键统计（记录数、待核验数、冲突数、
未完成学校）与下一步建议。

Excel 六个 Sheet：招生计划明细（自动筛选、冻结、状态高亮、人数为数值型可做数值筛选）、
采集报告、口径冲突、来源日志、学校清单、字段说明。

## 记录模板（最小可用）

```json
{
  "year": "2027", "province": "江苏", "school": "东南大学", "school_code": "10286",
  "college": "计算机科学与工程学院", "major": "电子信息", "major_code": "085400",
  "direction": "01 计算机技术", "study_mode": "全日制", "exam_mode": "全国统考",
  "planned_total": "120", "recommend_quota": "50", "unified_quota": "",
  "quota_basis": "专业目录拟招生人数为该专业全日制总数，含推免",
  "source_title": "东南大学2027年硕士研究生招生专业目录",
  "source_type": "专业目录",
  "source_url": "https://yzb.seu.edu.cn/2027/0915/c6676a123456/page.htm",
  "publish_date": "2026-09-15", "verify_status": "已核验", "is_current": "当前采用",
  "notes": "统考名额原文未单列"
}
```

字段全集、枚举值与状态流转规则见 `references/field-schema.md`。

## 迭代与复用

- 已有任务继续采集：直接运行 Stage 2→3，脚本按「学校+专业+方向+学习方式+考试方式+来源」去重。
- 官方更新或用户要求修正：只能**新增**记录（新 `source_url` 或 `is_current=历史版本`），
  不得直接改写旧行；`data/history/` 保留每次写入前快照。
- 换年份：新建 run 目录（`init_run.py --year <新年份>`），不要在同一目录混入多年数据。

## 合规与免责

- 仅采集高校与研招网**公开发布**的招生信息，不采集个人信息，不提交表单，不进行批量高频请求；
  同一站点请求间隔建议 ≥3 秒，并遵守目标站点 `robots.txt` 与使用条款。
- 遇到验证码、登录墙、付费墙、校内 IP 限制即停止，不采用任何规避手段（见「三条红线」第 3 条）。
- 采集结果中的学校名称、专业代码、招生人数等内容的著作权归属各高校及中国研究生招生信息网，
  本技能仅做结构化整理，不主张权利。
- 输出的「拟招生人数」为各校**计划口径**数据，可能与实际录取人数不同；本技能不构成报考建议，
  使用者须自行复核原文，一切决策后果由使用者承担。
- 数据结构与字段口径随年份与政策变化，不保证长期适用；重要用途请在采集后人工抽检。

## 资源

- `requirements.txt`：唯一第三方依赖 `openpyxl`。
- `scripts/common.py`：字段定义、枚举、CSV 读写、备份、日志（改字段只改这里）。
- `scripts/init_run.py`：建目录骨架与学校清单。
- `scripts/append_records.py`：校验、去重、冲突检测、增量入库。
- `scripts/export_excel.py`：导出多 Sheet Excel 与 Markdown 采集报告。
- `references/source-rules.md`：来源分级、域名模式、口径陷阱、访问受限处理。
- `references/field-schema.md`：字段/枚举/状态流转/工作簿结构。
- `references/collection-sop.md`：单校九步 SOP、关键词模板、试运行验收清单。
