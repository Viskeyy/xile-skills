# Skills for xile team

## invoice-processing — 发票识别与报销整理

把一堆发票丢进文件夹, 让 Claude 识别、分类、排序, 生成每人每购买方的报销申请单 (.docx) 和月度汇总表 (.xlsx)。

### 1. 安装

把技能目录复制到 Claude Code 的 skills 目录:

```bash
# 全局可用
cp -r skills/invoice-processing ~/.claude/skills/

# 或只在某个项目里可用
cp -r skills/invoice-processing /path/to/project/.claude/skills/
```

复制后新开一个会话, 输入 `/` 能看到 `invoice-processing` 即安装成功。

### 2. 环境准备

需要 Python 3.12+, 以及 `python-docx`、`openpyxl` 两个依赖:

```bash
# macOS
python3 -m pip install python-docx openpyxl   # 报 externally-managed 错误时加 --break-system-packages
# Windows
python -m pip install python-docx openpyxl
```

没装 Python 的话不用自己折腾, 直接开始用, 技能会停下来问你用哪种方式装。完整说明见 `skills/invoice-processing/references/环境准备.md`。

### 3. 准备发票目录

建一个当月目录, 下面按**人员**分文件夹, 发票原件 (图片/PDF) 直接放进去。类别子文件夹不用自己建, 技能会补:

```txt
2026年7月/
├── 黄卉/
│   ├── IMG_2381.jpg
│   ├── 滴滴行程单.pdf
│   └── ...
└── 曹英辰/
    └── ...
```

两点约定:

- **人员归属看文件夹名**, 不看发票上的姓名 — 放错文件夹就会归错人。
- **识别与分类只看发票内容**, 文件名随便叫, 不影响结果。

### 4. 使用

在会话里说清楚**目录**和**要做什么**即可:

```txt
帮我处理 ~/Desktop/报销/2026年7月 里的发票
```

技能会自动触发, 依次跑完识别 → 分类 → 排序 → 报销单 → 月度汇总 → 交叉核对。

只做其中几步也可以:

```txt
把 2026年7月 里的发票识别并分类归档, 先不用生成报销单
只重新生成 2026年7月 的月度汇总
```

过程中遇到任何存疑的地方 (金额大小写对不上、发票看不清、不知道归哪类), 技能会停下来问你, 不会自己猜。跑完会逐条列出所有没能处理的发票及原因。

### 5. 产出

全部直接写在源发票目录下, 发票原件仍按 `人员/类别` 归档:

```txt
2026年7月/
├── 黄卉/
│   ├── 差旅费/
│   ├── 交通费/
│   └── 办公费/
├── 曹英辰/
│   └── ...
├── 黄卉报销申请单中轻（雄安）2026-8-14.docx    # 每个 "姓名+购买方" 一份
└── 2026年7月发票月度汇总.xlsx
```

分类规则:

| 类别 | 包含 |
| --- | --- |
| 餐饮费 | 餐费 |
| 差旅费 | 机票、行程单、高铁票、火车票、住宿费、退票费 |
| 交通费 | 打车费 |
| 快递费 | 快递费 |
| 办公费 | 以上之外的一切 (咖啡、水果、办公用品、名片/标书制作费等) |

每人每类内部按 日期从远到近 → 金额从小到大 → 编号从小到大 排序。

### 6. 单独跑脚本 (可选)

一般不需要, 技能会自己调。要手动重跑某一份时:

```bash
# 报销单
python3 skills/invoice-processing/scripts/gen_receipt.py \
  --data data.json --source-dir "2026年7月" --output "黄卉报销申请单XX公司2026-8-14.docx"

# 月度汇总
python3 skills/invoice-processing/scripts/gen_summary.py \
  --data data.json --source-dir "2026年7月" --output "2026年7月发票月度汇总.xlsx"
```

`--output` 只能是文件名 (不含子目录), 文件写到 `--source-dir` 下。`--template` 可选, 默认用 `assets/` 里的模版。

`data.json` 结构:

<details>
<summary>gen_receipt.py</summary>

```json
{
  "name": "黄卉",
  "date": "2026-8-14",
  "total": "4130.50",
  "total_cn": "肆仟壹佰叁拾元伍角",
  "items": [
    {
      "category": "差旅费",
      "total": "3711.00",
      "rows": [["1", "6月19日", "940.00", "元", "机票 北京-厦门 SC2130"]]
    }
  ]
}
```

`items` 按报销单上的显示顺序排列, 每个类别只能出现一次, 空类别要整项删掉。各类别 `total` 之和必须等于 `total`, 对不上脚本会直接报错。

</details>

<details>
<summary>gen_summary.py</summary>

```json
{
  "month": "2026年7月",
  "records": [
    {
      "name": "黄卉",
      "category": "差旅费",
      "number": "26378324211049524479",
      "date": "7月6日",
      "amount": "940.00",
      "description": "机票 北京-厦门 SC2130"
    }
  ]
}
```

`records` 平铺即可, 分组、排序、合计由脚本处理。

</details>

### 常见问题

- **生成的报销单少了几条记录** — 有多少发票就填多少条, 脚本不合并。少了说明识别阶段就漏了, 看跑完列出的问题清单。
- **金额不对** — 一律以发票识别到的价税合计为准。大小写不一致时技能会问你以哪个为准, 别让它自己算。
- **快递费** — 有运单记录的话需要逐笔一一对应, 并写明始发地与送达地, 把运单文件一起放进目录。
- **改了模版之后脚本报错** — `assets/` 里的两个模版是脚本按结构定位填充的, 改动版式可能导致定位失败。改模版前先看脚本里的 `NESTED_INDEX` 和 "填写说明" 定位逻辑。
