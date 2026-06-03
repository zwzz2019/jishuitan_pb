# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

> **职责边界**：本文件只描述**流水线、代码架构、命令**。**所有领域规则**——每日人数、配额解析、辅助制度、个人特殊规则、周度变更等——一律以 `rules.md` 为唯一源。CLAUDE.md 内**不得**出现具体姓名、具体数字、白/黑名单成员、本周变更说明等任何会随 `rules.md` 变化而需要同步的内容；只能以 §章节号引用。

## Project Overview

Hospital radiology weekly scheduler (医院医生排班程序). A single Python script reads a weekly Excel roster, fills in `X` / `CT` / `MR` shifts and `辅助` (assist) entries Mon–Fri according to `rules.md`, and writes a numbered `_result.xlsx`.

- 领域规则源：`rules.md`（中文）。改任何规则、添加任何个人特殊规则、调整周度变更，**先改 `rules.md`，再改 `scheduler.py` 中对应的常量/分支**。
- 周度数据：`*.xlsx`（输入）、`*_result.xlsx`（输出）、`*.docx`（打印件）。当作数据处理，不要把规则知识刻进这些文件名。

## 排班工作流（必须遵循）

每次执行排班任务时，**必须**严格按照以下四步进行，不得跳过或合并：

### Step 1 — 大模型预推理（人工固定班次）

读取本周输入 Excel + `rules.md`，**纯推理**识别可以先行固定的班次，覆盖范围必须同时包含两类：

**(A) 单条规则直接固定**：`rules.md` §九 的硬性指定、列3 中的 `在周X`、`original[]` 中已填的格子。

**(B) 多条规则联合约束派生（重点 — 程序未必能枚举完）**。需要把以下因素**两两/多多组合**后判断哪些班次的人/天/班种已唯一确定：

- §二 每日人数硬上下限 + 已固定占位 → 反推剩余空位的人/班种。
- §5.3 一周最多 1 MR + §5.3 多 MR 例外名单 + `original[]` 中的 `MR盯机` → 反推个人 `max_mr` 与剩余 MR 配额必须落在哪几天。
- §5.4 回急白 + 个人配额 + 列3 `在周X` → 班种与日期联合锁死。
- §九 被辅助人辅助白/黑名单 + §6.5 `不能做MR辅助` 名单 + §6.3 半效辅助成对 → 唯一可配辅助人推导。
- §5.3 多 MR 例外人员的配额 + §二 每日 MR 上限 + §九 周内已固定的 MR 占位 → 反推多 MR 配额的人 MR 必须分布在哪几天。
- §5.1 追加班限制 + `original[]` 中的可追加格子 + §二 当天 X/CT 缺口 → 班种联合锁死。
- §二 余量分配优先级 + 已固定占位 → 推算各天 X/CT 配额上限。
- §九 覆盖式配额 + §5.2 班次间隔软约束 → 当配额数 ≥ 4 时分布几乎锁死。
- §十 介入处理 → 标注纯介入"可追加辅助"、复合介入"禁追加"。

> 上述列出的是**联合约束类型**，不是规则本身——具体姓名、具体数字、白名单成员请始终回 `rules.md` 查阅。

将 (A) + (B) 推出的预固定班次以**以下任一方式**落地，使 Step 2 不能再覆盖它们：
1. 写回输入 Excel 的对应格子（最稳，phase1–phase4 读 `original[]`）；或
2. 在 `scheduler.py` 的 `parse_quota()` / `phase1` 中以硬编码 / `constraints` 字段表达；或
3. 至少以注释形式列在交付物中，供 Step 4 终检对照。

**目的**：把规则联合推理出的 *deterministic* 部分从程序的搜索空间中拿掉，避免 phase1–phase4 因启发式排序在联合约束上踩坑。

### Step 2 — 程序执行排班

```bash
python3 scheduler.py
```

由 `main()` 中的 7 个 phase 完成剩余排班并写出 `*_result.xlsx`。**只跑程序，不做人工干预**。

### Step 3 — 程序自检

依赖 `validate(counts)` 打印的 `=== 验证 ===` 块以及程序内部硬约束（每日人数、上限、个人 `max_mr`、回急白禁 MR、辅助配对完整性）。

只有程序输出 `✅ 排班完成!` 才能进入 Step 4；若出现 `⚠️ 有约束未满足`，回到 Step 1 调整预固定或修正 `scheduler.py`，**不要**直接跳到 LLM 终检。

### Step 4 — 大模型纯推理终检（不依赖程序）

打开 `*_result.xlsx`，由大模型**逐人、逐天、逐条**对照 `rules.md` **全文**复核，**不得调用 `scheduler.py` 或任何脚本辅助判断**。重点覆盖程序未必能完整捕获的软约束 / 隐性规则，按 `rules.md` §三、§四、§5.2、§5.5、§六、§七、§八、§九、§十 顺序通览。

若发现任何不符项，回到 Step 1 修正预固定或改 `scheduler.py`，**重新跑全流水线**——禁止只手工改 `_result.xlsx` 的单元格。

## Commands

```bash
pip install pandas openpyxl   # 一次性
python3 scheduler.py          # 跑当前周
```

无 tests / linter / build。验证靠 `validate()` 的 `=== 验证 ===` 输出和肉眼对照 `*_result.xlsx`。

### Switching to a new week

`main()` 里的 `filepath` 是硬编码（scheduler.py 末尾 `main()` 中），换周需要编辑这一行。输出路径由 `.xlsx` 替换成 `_result.xlsx` 派生。

## Architecture

`scheduler.py` 是单文件流水线（~1370 行）。从上往下阅读；section banner（`# ===…`）是结构骨架。

### Data model

- `Doctor` dataclass 持有人级状态：`quota`、`schedule[5]`（Mon–Fri 决策）、`original[5]`（输入未改值）、`constraints`、`mr_count_this_week`、`max_mr`。
- 三类驱动几乎每一处分支：`senior`（被辅助人 = `SENIORS`）、`assistant`（纯辅助人 = `ASSISTANTS`）、`normal`（其余）。`马毅民` 在 `load_all` 中被硬跳过。
- `ASSISTANTS` 进一步切成 `HALF_EFFICIENCY` 与 `FULL_EFFICIENCY = ASSISTANTS - HALF_EFFICIENCY`；`NO_MR_ASSIST` 是 MR 辅助的 deny-list；`MULTI_MR_ALLOWED` 是周 MR 上限例外名单；`HUI_JI_BAI_PEOPLE` 是回急白名单。**这些常量的成员归属来自 `rules.md`，改成员名单时必须先改 `rules.md` 再同步常量；CLAUDE.md 不存任何成员名。**
- `parse_quota()` 是把 Excel 列3 转成 `(quota_dict, needs_scheduling, constraints, max_mr)` 的唯一入口，内部含**大量 per-person hardcoded 分支**对应 `rules.md` §九。`rules.md` §九 任何变化都需要在此函数同步。
- `Doctor.constraints` 是细粒度约束的承载字典（如 `pin_day`、`fixed_days`、`hui_ji_bai`、`mr_blocked_days`、`append_day`、`append_shift`、`is_pinned_assist` 等）。新增个人约束时**优先**走这个字典而不是再加全局名单。

### Pipeline (`main()`)

7 个 phase 顺序执行，**就地修改** `counts` 与每个 `Doctor.schedule`。**顺序敏感**，不要在不读完函数体的情况下重排。

1. `phase1` — 固定排班 + 给有 MR 配额的 senior 调用 `assign_quota`，按可用 MR 辅助候选数升序（最受约束者优先）。
2. `phase2` — 把 `APPENDABLE_SHIFTS` 类格子（追加班）落到当天，禁 MR；处理 `is_pinned_assist` 占位。
3. `phase3` — 列3 有显式 `XCTMR` 配额的普通医生。
4. `phase4` — 列3 仅是数字（灵活）的医生：按当日缺口选班种。
5. `validate(counts)` — 打印每日 X/CT/MR 与 `DAILY_TARGETS` 的差距并返回 ok 标记。
6. `phase6_5_optimize` — 局部 swap 优化连续日同班种的软违例（受 `can_swap` 约束保护，包括 `mr_blocked_days`）。
7. `phase5` — Stage-1 辅助：senior 的 X/CT/MR（非主班）必须配辅助。半效辅助走成对分支（若 `HALF_EFFICIENCY` 非空）。
8. `phase6` — Stage-2 辅助：剩余辅助配额按 `rules.md` §6.7 优先级派给普通医生。
9. `phase7` — 编号：MR1/MR2/MR3 给盯机配对（`asst_dinji` + `can_pair_dinji`），其余 MR 从 4 起、X/CT 从 1 起，渲染 `<senior_num>辅助` 并写盘。

### Cross-cutting invariants（修改前必读）

- `DAILY_TARGETS` 与 `get_upper_limit()` 的硬上限解释见 `rules.md` §二；CLAUDE.md 不复述具体数字。
- `count_existing()` 故意**不**把 `assistant` 行的 `MR盯机` 计入当日 MR 人数——那部分由 `phase7` 通过 `asst_dinji` / `can_pair_dinji()` 配对回填。改其中一个必须改另一个，否则 MR 人数会重复或缺失。
- `original[]` 是输入未改值，`schedule[]` 是程序决策。`get_shift_on_day()` 与 `phase7` 都同时读取两者。新增班种时两个读取点都要更新。
- 个人级日期/班种禁排约束统一走 `Doctor.constraints['mr_blocked_days']` 之类字段，并在 `find_best_day()` 与 `can_swap()` 两处过滤——这是一个"加在源头、检查在两处"的不变量，新增同类约束时务必同时更新两处。

## Repo configuration

- `.claude/settings.local.json` 预放行 `pip install *`、`python3 *`、以及把 `scheduler.py`/`rules.md` 软链到 `/root/` 的两条 `ln -s`。
- 无 `.gitignore`、CI、linter、虚拟环境。git 历史目前只有一个提交。
