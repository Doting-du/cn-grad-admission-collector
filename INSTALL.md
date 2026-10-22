# 安装与使用指南

面向第一次使用本技能的人。全程约 3 分钟，不需要 git 以外的任何工具。

## 一、前置条件

| 项目 | 要求 | 检查方式 |
| --- | --- | --- |
| WorkBuddy | 支持自定义技能（技能管理内可见） | 侧边栏「技能」页存在 |
| Python | 3.9 或更高 | `python --version` 或 `py -3 --version` |
| openpyxl | 3.0+，仅导出 Excel 用 | 由下方 pip 命令安装 |
| 网络 | 采集时需要访问高校官网与研招网 | 浏览器能打开 `*.edu.cn` |

说明：脚本本身**不发起任何网络请求**，只读写本地文件；抓取由 WorkBuddy 的联网能力完成。

## 二、安装（三种方式，任选一种）

### 方式 A：git clone（推荐，便于更新）

```bash
git clone https://github.com/Doting-du/cn-grad-admission-collector.git
mkdir -p ~/.workbuddy/skills
cp -r cn-grad-admission-collector/skills/cn-grad-admission-collector ~/.workbuddy/skills/
```

### 方式 B：下载 Release 压缩包

1. 打开 <https://github.com/Doting-du/cn-grad-admission-collector/releases>
2. 下载最新版 `cn-grad-admission-collector.zip`
3. 解压到 `~/.workbuddy/skills/`，确认最终路径为：

```
~/.workbuddy/skills/cn-grad-admission-collector/SKILL.md
```

> 常见错误：解压后多套了一层目录（如 `.../cn-grad-admission-collector/cn-grad-admission-collector/SKILL.md`），技能不会被识别。

### 方式 C：仅当前项目使用

```bash
mkdir -p <你的项目>/.workbuddy/skills
cp -r cn-grad-admission-collector/skills/cn-grad-admission-collector <你的项目>/.workbuddy/skills/
```

项目级安装只在该项目内可见，适合只给单个课题使用。

### Windows（PowerShell）

```powershell
git clone https://github.com/Doting-du/cn-grad-admission-collector.git
New-Item -ItemType Directory -Force "$env:USERPROFILE\.workbuddy\skills" | Out-Null
Copy-Item -Recurse .\cn-grad-admission-collector\skills\cn-grad-admission-collector "$env:USERPROFILE\.workbuddy\skills\"
```

### 安装依赖

```bash
# macOS / Linux / Git Bash
python -m pip install -r ~/.workbuddy/skills/cn-grad-admission-collector/requirements.txt

# Windows PowerShell
python -m pip install -r "$env:USERPROFILE\.workbuddy\skills\cn-grad-admission-collector\requirements.txt"
```

若提示 `pip` 不在 PATH，改用 `python3 -m pip`（macOS/Linux）或 `py -3 -m pip`（Windows）。

## 三、验证安装

```bash
SKILL="$HOME/.workbuddy/skills/cn-grad-admission-collector"

python "$SKILL/scripts/init_run.py" --help
python "$SKILL/scripts/append_records.py" --help
python "$SKILL/scripts/export_excel.py" --help
```

三条命令都能打印参数说明，即脚本可用。再执行一次最小闭环（会在临时目录生成一份示例，可随时删除）：

```bash
cd /tmp && python "$SKILL/scripts/init_run.py" --year 2027 --provinces 北京 --out ./demo_run
python "$SKILL/scripts/export_excel.py" --run ./demo_run   # 无记录时应提示「无数据」并退出，属正常
```

最后重启 WorkBuddy，在输入框输入 `/`，列表中应出现 `cn-grad-admission-collector`。

## 四、第一次使用（务必先试运行 5 校）

```
用 cn-grad-admission-collector 采集 2027 年江苏省 5 所高校的硕士招生计划，
输出到当前目录，先做试运行等我确认。
```

技能会先确认四项参数（年份 / 学校范围 / 输出目录 / 是否试运行），然后：

1. 建立任务目录与 `schools.csv` 清单
2. 逐校打开官方招生简章与专业目录（PDF 记录页码）
3. 校验入库，返回 warnings（数值不自洽、缺口径说明等）
4. 导出 Excel 与采集报告，交你抽查

**抽查这 5 点**，确认无误后再扩量：

- 任取 3 条，点链接能打开原页面/PDF，标题一致
- 表中人数与原文数字一致（不是机构数据）
- `人数口径说明` 能让人判断「含不含推免、专业级还是学院级」
- 未公布人数的项确实标了「无数据 / 待核验」，没有凭空填数
- 字段是否够用（是否要加学制、学费、复试科目等）

## 五、命令行用法（不经对话）

```bash
SKILL="$HOME/.workbuddy/skills/cn-grad-admission-collector"

# schools.txt 每行：学校名称,学校代码,省份
python "$SKILL/scripts/init_run.py" --year 2027 --schools schools.txt --out ./考研招生计划_2027
python "$SKILL/scripts/append_records.py" --run ./考研招生计划_2027 --input batch_01.json --batch trial5
python "$SKILL/scripts/export_excel.py" --run ./考研招生计划_2027
```

产物：

- `考研招生计划_<年份>.xlsx`（六工作表，已开自动筛选）
- `采集报告_<年份>.md`（已完成/未完成/缺失/异常/待核验/来源统计）
- `data/*.csv`（可编辑后由脚本读回；写入前自动快照到 `data/history/`）

## 六、更新与卸载

```bash
# 更新（方式 A 安装的）
cd cn-grad-admission-collector && git pull
cp -r skills/cn-grad-admission-collector ~/.workbuddy/skills/

# 卸载
rm -rf ~/.workbuddy/skills/cn-grad-admission-collector
```

已采集的数据在**你的任务目录**（如 `考研招生计划_2027/`），不在技能目录内，卸载不会删除数据。

Windows：

```powershell
Remove-Item -Recurse "$env:USERPROFILE\.workbuddy\skills\cn-grad-admission-collector"
```

## 七、常见问题

**Q1：技能列表里看不到。**
检查路径是否正好是 `.../skills/cn-grad-admission-collector/SKILL.md`（只套一层），然后重启 WorkBuddy。

**Q2：说了「采集考研招生计划」却没触发。**
描述匹配偏弱。改用 `/cn-grad-admission-collector`，或说「用 cn-grad-admission-collector 采集…」。

**Q3：导出时报 `ModuleNotFoundError: openpyxl`。**
依赖装到了另一个 Python。用与运行 WorkBuddy 相同的解释器执行 pip，或直接在命令里写全路径的 python。

**Q4：脚本报 `No such file or directory`。**
多半是相对路径问题。用绝对路径 `python "$SKILL/scripts/xxx.py"`，或先 `cd "$SKILL"`。

**Q5：某学校目录打不开。**
这是预期行为。技能遇验证码/登录墙会停止并在 `sources.csv` 记 `access_result`，报告里单列。请改用同一公告的公开页面，或到研招网专业目录查询核对。

**Q6：为什么「统考计划人数」总是空的？**
多数学校只公布「拟招生总人数」与「推免数」，未单列统考名额。技能禁止用减法推算，因此留空。若原文明确写了「统考拟招 X 人」，会自动填入。

**Q7：能用于往年（2025、2026）吗？**
可以，`init_run.py --year 2025` 新建任务即可。但往年旧公告可能已下线，成功率低于当年。

## 八、用于其他 AI 工具

`SKILL.md` 是标准 Markdown + YAML frontmatter，其它支持 Agent Skills 的工具也能读取。但需注意：

- 脚本依赖约定目录结构（`scripts/`、`references/`），迁移时保持相对位置
- 触发词与工具名（如 `present_files`）绑定 WorkBuddy，换到别的工具需按该工具的规则调整
- 文件读写与联网能力由宿主提供，脚本本身不联网

## 九、安全与隐私

- 脚本仅读写你指定的本地目录，不访问网络、不上传数据、不收集任何信息
- 不含埋点、不含远程配置、不修改系统设置
- 采集内容为高校公开发布的招生信息，不含个人信息
- 所有 Excel/CSV 都在你本地，仓库不包含任何采集结果

有问题请到 <https://github.com/Doting-du/cn-grad-admission-collector/issues> 反馈。
