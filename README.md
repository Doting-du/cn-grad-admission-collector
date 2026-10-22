# 全国高校考研官方招生计划采集 Skill

一个用于 [WorkBuddy](https://www.workbuddy.cn) 的自定义技能：按年份与学校范围，逐校检索中国大陆高校
**官方发布**的硕士研究生招生简章、专业目录、招生计划与调整公告，提取分专业招生人数，
标注口径与核验状态，输出可筛选的 Excel 与 Markdown 采集报告。

核心设计原则只有一句：**每一条人数都必须能回溯到官方原文；无法确认的一律标「待核验」，绝不猜测或推算。**

## 它解决什么问题

考研招生人数散落在各校研究生院、研招网、学院官网，格式各异、口径混乱（含不含推免？学院级还是专业级？），
且 9 月目录与 10 月调整公告常有出入。手工整理一省数据通常要几天，且极易把「学院总数」误当「专业人数」。

本技能把这件事变成可复现的流水线，并把三个最容易出错的地方写进代码里强制执行：

| 风险 | 强制手段 |
| --- | --- |
| 用机构数据、搜索摘要充当年份人数 | 来源域名白名单（`.edu.cn` / `.ac.cn` / `.gov.cn` / `chsi.com.cn`），其余**直接拒收** |
| 用「拟招生 − 推免」倒推统考名额 | 缺「人数口径说明」不允许标「已核验」；数值不自洽只告警、不修正 |
| 多版本公告相互覆盖，历史丢失 | 数据 append-only，写入前快照到 `data/history/`，版本字段区分「当前采用 / 历史版本」 |

## 安装

```bash
git clone https://github.com/<你的用户名>/cn-grad-admission-collector.git
cd cn-grad-admission-collector

# 用户级安装（所有项目可用）
cp -r skills/cn-grad-admission-collector ~/.workbuddy/skills/

# 或项目级安装（仅当前项目）
mkdir -p <项目>/.workbuddy/skills && cp -r skills/cn-grad-admission-collector <项目>/.workbuddy/skills/
```

Windows（PowerShell）：

```powershell
Copy-Item -Recurse .\skills\cn-grad-admission-collector "$env:USERPROFILE\.workbuddy\skills\"
```

依赖：Python 3.9+ 与 `openpyxl`（仅用于导出 Excel）。

```bash
python -m pip install -r ~/.workbuddy/skills/cn-grad-admission-collector/requirements.txt
```

## 使用

安装后在 WorkBuddy 中直接说：

```
用 cn-grad-admission-collector 采集 2027 年江苏省 5 所高校的硕士招生计划，
输出到当前目录，先做试运行等我确认。
```

也可显式调用 `/cn-grad-admission-collector`。

### 命令行流程（不经对话也能用）

```bash
SKILL="$HOME/.workbuddy/skills/cn-grad-admission-collector"

# 1) 建任务：schools.txt 每行「学校名称,学校代码,省份」
python "$SKILL/scripts/init_run.py" --year 2027 --schools schools.txt --out ./考研招生计划_2027

# 2) 校验入库：batch_01.json 为本校提取结果
python "$SKILL/scripts/append_records.py" --run ./考研招生计划_2027 --input batch_01.json --batch trial5

# 3) 导出 Excel + 采集报告
python "$SKILL/scripts/export_excel.py" --run ./考研招生计划_2027
```

### 输出

- `考研招生计划_<年份>.xlsx`：六个工作表 —— 招生计划明细（自动筛选、冻结窗格、核验状态色标、
  人数为数值型可做数值筛选）、采集报告、口径冲突、来源日志、学校清单、字段说明
- `采集报告_<年份>.md`：已完成 / 未完成学校、缺失数据、异常数据、待人工核验项、来源统计
- `data/records.csv` 等四张表：可被 Excel 直接打开编辑，脚本能读回，写入前自动备份

## 字段

24 个字段，核心为：招生年份、省份、学校名称与代码、学院、专业名称与代码、研究方向、学习方式、
考试方式、拟招生人数、推免人数、统考计划人数、人数口径说明、原文标题、材料类型、来源网址、
发布日期、核验状态（已核验 / 待核验 / 口径冲突 / 无数据）、版本。

完整定义见 [`skills/cn-grad-admission-collector/references/field-schema.md`](skills/cn-grad-admission-collector/references/field-schema.md)。

## 工作流

1. **确认参数** → 2. **建任务与学校清单** → 3. **逐校来源发现**（打开原始网页/PDF 读取，PDF 记页码）
→ 4. **校验入库**（脚本返回 warnings 必须逐条处理）→ 5. **5 校试运行验收** →
6. **按省扩量** → 7. **导出 Excel 与报告**

细节见 [`SKILL.md`](skills/cn-grad-admission-collector/SKILL.md)、
[`collection-sop.md`](skills/cn-grad-admission-collector/references/collection-sop.md)、
[`source-rules.md`](skills/cn-grad-admission-collector/references/source-rules.md)。

## 已知限制

- **不绕过任何访问限制**：遇验证码、登录墙、付费墙即停止并登记，因此部分学校可能只能取到院校级信息。
- **不做 OCR 兜底**：扫描版 PDF 标为「PDF无法解析」，不自动识别，避免引入不可核验的数字。
- **不对接学校私有接口**：只读取公开页面与公开 PDF。
- **年份强相关**：招生目录结构逐年调整，跨年份复用前需重新核对字段口径。
- **未做全量实测**：初始版本以 2 校构造数据跑通全链路，真实采集的首批结果仍建议人工抽检。

## 合规与免责

仅采集公开发布的招生信息，不采集个人信息、不提交表单、不进行高频请求，遵守目标站点 `robots.txt` 与使用条款。
采集内容的著作权归属各高校与中国研究生招生信息网，本项目仅做结构化整理。
输出数据为各校**计划口径**，可能与实际录取人数不同，**不构成报考建议**，使用者须自行复核原文。

## 许可

[MIT](LICENSE)
