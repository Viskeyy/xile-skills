# invoice-processing

发票识别与报销整理技能。把发票图片或 PDF 按人员放进文件夹，交给 Agent，即可自动完成识别、分类、排序，并生成报销申请单（.docx）和月度汇总表（.xlsx），替代逐张人工整理。

## 效果预览

处理前
<img width="1784" height="1196" alt="img_v3_02150_804500e8-965a-4d5f-84b0-16a660cda6dg" src="https://github.com/user-attachments/assets/56727f19-919b-41ea-a5af-8b28fe19890a" />

处理后
<img width="1786" height="1198" alt="img_v3_02150_2456137f-df2f-4ef6-994a-c5cbdeb120eg" src="https://github.com/user-attachments/assets/919974ce-f6f5-4175-a037-959504ed0d29" />

## 准备发票目录

按月份建一个总文件夹，每位员工一个子文件夹，发票原件直接放进去：

```txt
2026年7月/
├── 张三/
│   ├── IMG_2381.jpg
│   ├── 滴滴行程单.pdf
│   └── ...
└── 李四/
    └── ...
```

两条规则：

- 子文件夹名称就是报销单上的姓名，请用真实姓名。人员归属以文件夹名称为准，不读取发票上的姓名信息
- 分类只看发票内容，与文件名无关，发票不需要改名

PDF、PNG、JPG 等常见格式都可以。

## 安装

复制技能目录到 Agent skills 目录：

```bash
# 全局安装（所有项目可用）
cp -r skills/invoice-processing ~/.agent/skills/

# 项目级安装（仅当前项目可用）
cp -r skills/invoice-processing /path/to/project/.agent/skills/
```

安装完成后重启会话，输入 `/` 即可在技能列表中看到 invoice-processing。

## 环境要求

技能脚本需要 Python 3.9 或更高版本，以及 `python-docx`、`openpyxl` 两个依赖包：

```bash
# macOS / Linux
python3 -m pip install python-docx openpyxl

# Windows
python -m pip install python-docx openpyxl
```

机器上没有 Python 时，技能会先询问安装方式再引导安装，详见 `skills/invoice-processing/references/环境准备.md`。

## 使用

在会话中提供发票目录路径和任务描述：

```txt
帮我处理 ~/Desktop/报销/2026年7月 里的发票
```

技能按 识别 → 分类 → 排序 → 生成报销单 → 生成月度汇总 → 交叉核对 的顺序执行，识别发票内容需要联网。

支持只执行部分步骤：

```txt
识别并分类归档发票，暂不生成报销单
仅重新生成 2026年7月 的月度汇总
仅生成「张三」的报销申请单
```

**存疑即停**：遇到金额大小写不一致、发票内容模糊、分类归属不明等情况，技能会暂停并等待确认，不会自行推断。处理结束后会列出所有未能处理的发票及原因，确认没有遗漏后再提交报销。

## 输出结果

生成的文件直接写入发票目录，发票原件移动至 `人员/类别` 子目录归档：

```txt
2026年7月/
├── 张三/
│   ├── 差旅费/
│   ├── 交通费/
│   └── 办公费/
├── 李四/
│   └── ...
├── 张三报销申请单XX公司2026-8-14.docx    # 每位员工每家购买方各一份
└── 2026年7月发票月度汇总.xlsx
```

- **报销申请单（.docx）**：各类金额与合计都已填好，拿来就能用
- **月度汇总表（.xlsx）**：当月全部发票按人员和类别分组，含组内合计与总计

### 分类

| 类别   | 适用范围                                     |
| ------ | -------------------------------------------- |
| 餐饮费 | 餐费                                         |
| 差旅费 | 机票、行程单、高铁票、火车票、住宿费、退票费 |
| 交通费 | 打车费                                       |
| 办公费 | 饮料、办公用品、名片/logo/标书制作费等       |

每张发票只归入一个类别，移动（而非复制）到对应人员的类别子文件夹。

### 排序

每位员工的每类发票按 日期（从远到近）→ 金额（从小到大）→ 编号（从小到大）排列。

## 脚本接口（可选）

正常使用无需手动调用，技能会自动执行。需要重跑某个产物时：

```bash
# 生成报销单
python3 skills/invoice-processing/scripts/gen_receipt.py \
  --data data.json --source-dir "2026年7月" --output "张三报销申请单XX公司2026-8-14.docx"

# 生成月度汇总表
python3 skills/invoice-processing/scripts/gen_summary.py \
  --data data.json --source-dir "2026年7月" --output "2026年7月发票月度汇总.xlsx"
```

- `--output` 只接受文件名，不能包含子目录路径，输出文件保存到 `--source-dir` 指定的目录
- `--template` 可选，指定自定义模版路径，默认使用 `assets/` 目录下的标准模版

### 数据格式

<details>
<summary>报销单（gen_receipt.py）</summary>

```json
{
  "name": "张三",
  "date": "2026-8-14",
  "total": "4130.50",
  "items": [
    {
      "category": "差旅费",
      "total": "3711.00",
      "rows": [["1", "6月19日", "940.00", "元", "机票 北京-厦门 SC2130"]]
    }
  ]
}
```

- 金额大写由脚本根据 `total` 自动转换，无需填写
- 每个类别在 `items` 中只能出现一次，空类别整项删除
- 报销单的排版顺序由模版决定（固定为餐饮/办公/差旅/交通），`items` 的顺序不影响输出
- 各类别的 `total` 之和必须等于顶层 `total`，否则脚本报错

</details>

<details>
<summary>月度汇总（gen_summary.py）</summary>

```json
{
  "month": "2026年7月",
  "records": [
    {
      "name": "张三",
      "category": "差旅费",
      "number": "26378324211049524479",
      "date": "7月6日",
      "amount": "940.00",
      "description": "机票 北京-厦门 SC2130"
    }
  ]
}
```

`records` 按自然顺序填写即可，分组、排序及合计由脚本自动处理。

</details>

## 常见问题

**报销单记录缺失**

脚本按发票张数生成记录，不合并。出现缺失时，先检查识别阶段输出的问题清单。

**金额有误**

一律以发票上识别的价税合计为准。大小写不一致时技能会暂停确认，不自行推断；报销单上的大写金额由脚本从小写金额转换生成，不依赖模型输出。

**修改模版后脚本报错**

`assets/` 里的模版依赖脚本的结构定位逻辑填充，改动版式会导致定位失败并报错（不会生成不完整的文件）。修改前先确认脚本中的 `NESTED_INDEX` 配置和「填写说明」定位逻辑。
