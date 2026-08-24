# AGENTS.md

给在**本仓库上工作**的 agent 看的。使用技能本身不用读这份 — 使用说明在 `skills/invoice-processing/SKILL.md`, 由 Claude Code 自动加载。

## 仓库是什么

一个 Claude Code 技能仓库。目前只有一个技能: `skills/invoice-processing`, 处理发票识别、分类、报销单与月度汇总。没有构建、没有 CI、没有测试套件。

```txt
README.md                              # 人类使用文档
AGENTS.md                              # 本文件
skills/invoice-processing/
├── SKILL.md                           # 技能主体: 通用原则 + 步骤①~⑥ + 完成判据
├── scripts/gen_receipt.py             # 报销单 .docx 生成 (含金额数字→大写转换)
├── scripts/gen_summary.py             # 月度汇总 .xlsx 生成
├── references/环境准备.md              # 按需加载: Python 与依赖安装
└── assets/                            # 二进制模版与已完成示例, 脚本按结构定位填充
```

`skills/invoice-processing/` 是交付物, 会被整个复制到 `~/.claude/skills/`。只放运行技能真正需要的东西 —— 测试文件、开发笔记、临时数据一律不进这个目录 (唯一的自检是 `to_cn_amount` 的 doctest, 它同时是那个函数的文档)。

## 改动前必须知道的

**SKILL.md 是提示词, 不是文档。** 每一句都会进上下文。加内容前先问是否真的改变执行结果, 不改变就别加。冗长的 SKILL.md 会稀释关键约束。

**脚本按结构定位模版, 不按下标。** 改 `assets/` 里的 .docx/.xlsx 会直接打破脚本:

- `gen_receipt.py` 依赖备注单元格内**恰好 5 个嵌套表** (`NESTED_INDEX`: 餐饮/办公/差旅/交通/快递), 依赖类别标题文本形如 `差旅费用（　）` (`CATEGORY_NAMES` + `TITLE_TEMPLATE`), 依赖存在 `共计费用：` 段落, 依赖 body 末尾存在含 "填写说明" 的段落作为截断锚点。表头位置写死在 `t0.rows[0..4]`。类别在文档里的排版顺序也由模版固定, **序号按这个顺序生成** —— 调整模版里类别块的先后, 序号跟着变。
- `gen_summary.py` 从模版第 1/3/4/5/6 行分别取标题、人员、类别、表头、数据行的样式, 然后清空整表重建。行的**内容**无所谓, **样式所在行号**不能动。

改模版 = 同时改脚本里的定位逻辑。

**填入内容的字体字号写死在代码里, 不在模版。** `HEADER_RPR` (表头填充: 中文宋体 + 数字 Calibri, 小三) 和 `ROW_RPR` (嵌套表数据行: 全宋体 11) 由 `make_rpr` 构造后 deepcopy 到每个填入的 run 上。在 Word 里改模版字体只影响标签、类别标题、`共计费用：` 和嵌套表表头行 —— 那些走 `fill_para_text` 或根本不被脚本触碰; 脚本填进去的值一律覆盖成上面两个常量。要改填充字体, 改常量的参数, 别改模版。

**脚本一律硬失败, 不静默降级。** 已有的自检 (类别重复、rows 非 5 元素、类别合计 ≠ 总金额、月度记录缺字段/类别非法/日期无法解析、模版定位失败、`--source-dir` 不存在) 都是 `sys.exit`。新增逻辑保持这个风格, 别加兜底默认值或 "找不到就跳过"。

**金额大写由脚本算, 不接受外部传入。** `to_cn_amount` 从 `total` 转换, data.json 里的 `total_cn` 会被忽略。模版 row2 的 `金额(大写)` 是表单必填格, 不能留空。改了转换逻辑跑它的 doctest, 其中 `1857.02` 一例锁的是 `assets/` 里人工填好的那份示例。

**`--output` 只收文件名, 输出落在 `--source-dir` 下, 该目录必须已存在。** 两个脚本都校验了这点, 别放宽。

**金额一律 `Decimal`。** 两个脚本各有一个 `money`/`amount` 辅助函数负责解析并在失败时报错, 用它, 不要引入 float 比较或求和。

## 怎么验证改动

没有测试框架。改金额转换跑 doctest, 改别的手动跑一遍:

```bash
cd skills/invoice-processing
python3 -m doctest scripts/gen_receipt.py    # 大写金额, 无输出即通过

mkdir -p /tmp/inv && cat > /tmp/d.json <<'EOF'
{"name":"测试","date":"2026-8-14","total":"940.00",
 "items":[{"category":"差旅费","total":"940.00",
 "rows":[["1","6月19日","940.00","元","机票 北京-厦门 SC2130"]]}]}
EOF
python3 scripts/gen_receipt.py --data /tmp/d.json --source-dir /tmp/inv --output t.docx
```

汇总同理, 用 `gen_summary.py` 顶部 docstring 里的结构。打开产物肉眼看版式 —— 脚本能跑通不代表版式没崩, 列宽、合并单元格、字体都得对照 `assets/` 里的模版和已完成示例看。

改了自检或金额逻辑, 顺手补一个反例 (故意让类别合计对不上) 确认它确实报错。

## 约定

- 中文文档用中文半角标点 + 空格分隔, 与现有文件保持一致 (`, ` 而非 `，`)。
- 脚本只依赖 `python-docx` / `openpyxl`, 不加新依赖。
- commit 用 conventional commits, scope 是技能名: `feat(invoice-processing): ...`。
- `assets/` 里的示例文件含真实人名与金额, 新增示例前先脱敏。
