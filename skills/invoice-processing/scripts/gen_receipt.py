# -*- coding: utf-8 -*-
"""按模版生成报销申请单 .docx。

用法:
    python gen_receipt.py --data data.json --source-dir 源发票目录 --output out.docx
    python gen_receipt.py --data data.json --source-dir 源发票目录 --output out.docx --template 其他模版.docx

data.json 结构:
{
  "name": "黄卉",                 // 姓名
  "date": "2026-8-14",           // 生成日期
  "total": "4130.50",            // 总金额; 大写由脚本转换, 无需提供
  "items": [                     // 类别; 顺序任意, 排版顺序与序号由模版决定
    {
      "category": "差旅费",      // 餐饮费 | 办公费 | 差旅费 | 交通费 | 快递费
      "total": "3711.00",        // 该类合计
      "rows": [                  // 数据行: [序号, 日期, 金额, 单位, 说明]
        ["1", "6月19日", "940.00", "元", "机票 北京-厦门 SC2130"]
      ]
    }
  ]
}
"""
import argparse
import copy
import json
import os
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import sys

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.join(SCRIPT_DIR, '..', 'assets', '报销申请单模版.docx')

CATEGORY_NAMES = {
    '餐饮费': '餐饮费用',
    '办公费': '办公费用',
    '差旅费': '差旅费用',
    '交通费': '交通费用',
    '快递费': '快递费用',
}
NESTED_INDEX = {'餐饮费': 0, '办公费': 1, '差旅费': 2, '交通费': 3, '快递费': 4}

CN_DIGITS = '零壹贰叁肆伍陆柒捌玖'
CN_UNITS = ('', '拾', '佰', '仟')


def money(value, field):
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError):
        sys.exit('%s 不是有效金额: %r' % (field, value))


def cn_section(num):
    """0..9999 -> 汉字, 内部零合并 (4007 -> 肆仟零柒)。"""
    text, pending_zero = '', False
    for i, ch in enumerate(str(num).zfill(4)):
        digit = int(ch)
        if digit == 0:
            pending_zero = True
            continue
        if pending_zero and text:
            text += CN_DIGITS[0]
        pending_zero = False
        text += CN_DIGITS[digit] + CN_UNITS[3 - i]
    return text


def cn_integer(num):
    if num == 0:
        return CN_DIGITS[0]
    if num >= 10 ** 8:
        sys.exit('总金额超出支持范围 (需小于 1 亿元): %d' % num)
    high, low = divmod(num, 10000)
    if not high:
        return cn_section(low)
    if low == 0:
        return cn_section(high) + '万'
    return cn_section(high) + '万' + (CN_DIGITS[0] if low < 1000 else '') + cn_section(low)


def to_cn_amount(value):
    """金额数字转大写。自检: python3 -m doctest scripts/gen_receipt.py

    >>> to_cn_amount(Decimal('4130.50'))
    '肆仟壹佰叁拾元伍角'
    >>> to_cn_amount(Decimal('1857.02'))  # 与 assets/ 中已完成示例逐字一致
    '壹仟捌佰伍拾柒元零贰分'
    >>> to_cn_amount(Decimal('940.00'))
    '玖佰肆拾元整'
    >>> to_cn_amount(Decimal('4007.00'))
    '肆仟零柒元整'
    >>> to_cn_amount(Decimal('1000500.00'))
    '壹佰万零伍佰元整'
    >>> to_cn_amount(Decimal('1234567.89'))
    '壹佰贰拾叁万肆仟伍佰陆拾柒元捌角玖分'
    """
    cents = int((value * 100).to_integral_value(ROUND_HALF_UP))
    if cents < 0:
        sys.exit('总金额不能为负: %s' % value)
    yuan, rest = divmod(cents, 100)
    jiao, fen = divmod(rest, 10)
    text = cn_integer(yuan) + '元'
    if not rest:
        return text + '整'
    if jiao:
        text += CN_DIGITS[jiao] + '角'
    elif yuan:
        text += CN_DIGITS[0]
    if fen:
        text += CN_DIGITS[fen] + '分'
    return text

TITLE_TEMPLATE = '（\u3000）'


def make_rpr(latin, east_asia, half_points):
    """构造 run 格式。一个 rFonts 同时管中西文, Word 按字符所属文种自动分派。

    half_points 是半磅: 小三 = 15pt = 30, 11pt = 22。
    """
    return parse_xml(
        '<w:rPr %s><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:hint="eastAsia"/>'
        '<w:sz w:val="%s"/><w:szCs w:val="%s"/></w:rPr>'
        % (nsdecls('w'), latin, latin, east_asia, half_points, half_points))


HEADER_RPR = make_rpr('Calibri', '宋体', '30')  # 表头填充: 中文宋体 + 数字 Calibri, 小三
ROW_RPR = make_rpr('宋体', '宋体', '22')  # 嵌套表数据行: 全宋体 11


def fill_cell(el_cell, text, rpr=None):
    for p in el_cell.findall(qn('w:p'))[1:]:
        el_cell.remove(p)
    p = el_cell.findall(qn('w:p'))[0]
    for r in p.findall(qn('w:r')):
        p.remove(r)
    r = p.makeelement(qn('w:r'), {})
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = r.makeelement(qn('w:t'), {})
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t)
    p.append(r)


def fill_data(tbl_xml, items, rpr=None):
    rows = tbl_xml.findall(qn('w:tr'))
    n = len(rows) - 1
    if len(items) > n:
        for _ in range(len(items) - n):
            tbl_xml.append(copy.deepcopy(rows[n]))
    rows = tbl_xml.findall(qn('w:tr'))
    for i, item in enumerate(items):
        for ci, val in enumerate(item):
            fill_cell(rows[i + 1].findall(qn('w:tc'))[ci], val, rpr)
    for i in range(len(items) + 1, len(rows)):
        tbl_xml.remove(rows[i])



def fill_para_text(p, text, rpr=None):
    runs = p.findall(qn('w:r'))
    if runs:
        r = runs[0]
        for extra in runs[1:]:
            p.remove(extra)
        for t in r.findall(qn('w:t')):
            r.remove(t)
        t = r.makeelement(qn('w:t'), {})
    else:
        r = p.makeelement(qn('w:r'), {})
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        t = r.makeelement(qn('w:t'), {})
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t)
    p.append(r)


def main():
    parser = argparse.ArgumentParser(description='按模版生成报销申请单')
    parser.add_argument('--data', required=True, help='数据 JSON 文件路径')
    parser.add_argument('--source-dir', required=True, help='源发票目录; 生成文件直接写入该目录')
    parser.add_argument('--output', required=True, help='输出文件名, 只能是当前目录下的 .docx 文件名')
    parser.add_argument('--template', default=DEFAULT_TEMPLATE, help='模版 .docx 路径')
    args = parser.parse_args()
    if Path(args.output).name != args.output or Path(args.output).suffix.lower() != '.docx':
        sys.exit('--output 必须是当前目录下的 .docx 文件名, 不得包含子目录')
    if not os.path.isdir(args.source_dir):
        sys.exit('--source-dir 不存在或不是目录: %s' % args.source_dir)
    output = os.path.join(args.source_dir, args.output)

    with open(args.data, encoding='utf-8') as f:
        data = json.load(f)

    missing = [k for k in ('name', 'date', 'total', 'items') if k not in data]
    if missing:
        sys.exit('data.json 缺少字段: %s' % ', '.join(missing))
    bad = [i['category'] for i in data['items'] if i.get('category') not in CATEGORY_NAMES]
    if bad:
        sys.exit('未知类别: %s (仅支持 %s)' % (', '.join(bad), '/'.join(CATEGORY_NAMES)))

    # ---- 输入自检: 类别唯一、rows 结构 (不合格即报错, 不静默丢数据) ----
    if not data['items']:
        sys.exit('items 为空: 报销单至少要有一个类别')
    seen_cats = []
    for i, item in enumerate(data['items']):
        cat = item['category']
        if cat in seen_cats:
            sys.exit('items 中类别 "%s" 重复 (第 %d 项): 每个类别只能出现一次, 请先合并' % (cat, i + 1))
        seen_cats.append(cat)
        rows = item.get('rows')
        if not isinstance(rows, list):
            sys.exit('类别 "%s" 缺少 rows 列表' % cat)
        if not rows:
            sys.exit('类别 "%s" 的 rows 为空: 未使用的类别应从 items 中移除' % cat)
        for j, r in enumerate(rows):
            if not isinstance(r, list) or len(r) != 5:
                sys.exit('类别 "%s" 第 %d 行 rows 应为 5 个元素 [序号,日期,金额,单位,说明], 实际是 %r' % (cat, j + 1, r))

    # ---- 金额自检: 类别合计与总金额一致 ----
    item_sum = Decimal(0)
    for item in data['items']:
        item_sum += money(item['total'], '类别 "%s" 的合计' % item['category'])
    declared_total = money(data['total'], '总金额')
    if item_sum != declared_total:
        sys.exit('各类别合计 (%s) 与总金额 (%s) 不一致, 请核对' % (item_sum, declared_total))
    total_text = '%.2f' % declared_total
    total_cn = to_cn_amount(declared_total)

    doc = Document(args.template)
    t0 = doc.tables[0]
    cell = t0.rows[4].cells[1]
    tc = cell._tc
    nested = tc.findall(qn('w:tbl'))
    if len(nested) != len(NESTED_INDEX):
        sys.exit('模版 %s 备注单元格内有 %d 个类别表, 预期 %d (餐饮/办公/差旅/交通/快递); 若模版被改动请先核对脚本 NESTED_INDEX' % (args.template, len(nested), len(NESTED_INDEX)))
    titles = {}
    for el in list(tc):
        if el.tag == qn('w:p'):
            txt = ''.join(t.text or '' for t in el.iter(qn('w:t')))
            titles.setdefault(txt, el)

    # ---- 表头 ----
    fill_cell(t0.rows[0].cells[1]._tc, data['name'], HEADER_RPR)
    fill_cell(t0.rows[1].cells[1]._tc, data['date'], HEADER_RPR)
    fill_cell(t0.rows[1].cells[3]._tc, total_text + '元', HEADER_RPR)
    fill_cell(t0.rows[2].cells[1]._tc, total_cn, HEADER_RPR)
    fill_cell(t0.rows[3].cells[1]._tc, data['name'], HEADER_RPR)

    # ---- 删除未使用的类别块, 填充保留的类别 (序号按模版中的排版顺序从 1 递增) ----
    # 类别标题与共计费用沿用模版段落自身的格式, 故不传 rpr
    by_category = {item['category']: item for item in data['items']}
    order = 0
    for cat, idx in NESTED_INDEX.items():
        title_key = CATEGORY_NAMES[cat] + TITLE_TEMPLATE
        title_el = titles.get(title_key)
        if title_el is None:
            sys.exit('模版 %s 中找不到标题段落 "%s", 可能是模版被改动' % (args.template, title_key))
        item = by_category.get(cat)
        if item is None:
            title_el.getparent().remove(title_el)
            nested[idx].getparent().remove(nested[idx])
            continue
        order += 1
        fill_para_text(title_el, '%d、%s（%s元）' % (order, CATEGORY_NAMES[cat], '%.2f' % money(item['total'], '类别合计')))
        fill_data(nested[idx], item['rows'], ROW_RPR)

    # ---- 共计费用 ----
    if '共计费用：' not in titles:
        sys.exit('模版 %s 中找不到 "共计费用：" 段落, 可能是模版被改动' % args.template)
    fill_para_text(titles['共计费用：'], '共计费用：' + total_text + '元')

    # ---- 删除 body 末尾的填写说明块及重复的审查段落 (按内容定位, 避免硬编码下标) ----
    body = doc.element.body
    elements = list(body)
    start_idx = None
    for i, el in enumerate(elements):
        if el.tag == qn('w:p') and '填写说明' in ''.join(t.text or '' for t in el.iter(qn('w:t'))):
            start_idx = i
            break
    if start_idx is None:
        sys.exit('模版 %s 中找不到 "填写说明" 提示块, 可能是模版改版, 请人工核对' % args.template)
    review_count = 0
    for el in elements[:start_idx]:
        if el.tag == qn('w:p') and '审查' in ''.join(t.text or '' for t in el.iter(qn('w:t'))):
            review_count += 1
            if review_count > 1:
                body.remove(el)
    for el in elements[start_idx:]:
        body.remove(el)

    doc.save(output)
    print('已生成: %s | 合计: %s元 %s' % (output, total_text, total_cn))


if __name__ == '__main__':
    main()
