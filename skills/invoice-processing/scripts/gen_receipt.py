# -*- coding: utf-8 -*-
"""按模版生成报销申请单 .docx。

用法:
    python gen_receipt.py --data data.json --output out.docx
    python gen_receipt.py --data data.json --output out.docx --template 其他模版.docx

data.json 结构:
{
  "name": "黄卉",                 // 姓名
  "date": "2026-8-14",           // 生成日期
  "total": "4130.50",            // 总金额
  "total_cn": "肆仟壹佰叁拾元伍角",  // 总金额大写
  "items": [                     // 按显示顺序排列的类别
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
import sys

from docx import Document
from docx.oxml.ns import qn

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.join(SCRIPT_DIR, '..', 'templates', '报销申请单模版.docx')

CATEGORY_NAMES = {
    '餐饮费': '餐饮费用',
    '办公费': '办公费用',
    '差旅费': '差旅费用',
    '交通费': '交通费用',
    '快递费': '快递费用',
}
NESTED_INDEX = {'餐饮费': 0, '办公费': 1, '差旅费': 2, '交通费': 3, '快递费': 4}
TITLE_TEMPLATE = '（\u3000）'


def run_rpr(el):
    r = el.find(qn('w:r'))
    return r.find(qn('w:rPr')) if r is not None else None


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
    parser.add_argument('--output', required=True, help='输出 .docx 路径')
    parser.add_argument('--template', default=DEFAULT_TEMPLATE, help='模版 .docx 路径')
    args = parser.parse_args()

    with open(args.data, encoding='utf-8') as f:
        data = json.load(f)

    missing = [k for k in ('name', 'date', 'total', 'total_cn', 'items') if k not in data]
    if missing:
        sys.exit('data.json 缺少字段: %s' % ', '.join(missing))
    bad = [i['category'] for i in data['items'] if i.get('category') not in CATEGORY_NAMES]
    if bad:
        sys.exit('未知类别: %s (仅支持 %s)' % (', '.join(bad), '/'.join(CATEGORY_NAMES)))

    # ---- 输入自检: 类别唯一、rows 结构 (不合格即报错, 不静默丢数据) ----
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
    from decimal import Decimal
    item_sum = Decimal(0)
    for item in data['items']:
        item_sum += Decimal(str(item['total']))
    declared_total = Decimal(str(data['total']))
    if item_sum != declared_total:
        sys.exit('各类别合计 (%s) 与总金额 (%s) 不一致, 请核对' % (item_sum, declared_total))

    doc = Document(args.template)
    t0 = doc.tables[0]
    ref_rpr = run_rpr(t0.cell(0, 0)._tc)
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
    fill_cell(t0.rows[0].cells[1]._tc, data['name'], ref_rpr)
    fill_cell(t0.rows[1].cells[1]._tc, data['date'], ref_rpr)
    fill_cell(t0.rows[1].cells[3]._tc, data['total'] + '元', ref_rpr)
    fill_cell(t0.rows[2].cells[1]._tc, data['total_cn'], ref_rpr)
    fill_cell(t0.rows[3].cells[1]._tc, data['name'], ref_rpr)

    # ---- 删除未使用的类别块, 填充保留的类别 (序号按 items 顺序从 1 递增) ----
    for cat, idx in NESTED_INDEX.items():
        title_key = CATEGORY_NAMES[cat] + TITLE_TEMPLATE
        title_el = titles.get(title_key)
        order = [i for i, item in enumerate(data['items']) if item['category'] == cat]
        if not order:
            if title_el is not None:
                title_el.getparent().remove(title_el)
            nested[idx].getparent().remove(nested[idx])
            continue
        item = data['items'][order[0]]
        if title_el is not None:
            fill_para_text(title_el, '%d、%s（%s元）' % (order[0] + 1, CATEGORY_NAMES[cat], item['total']), ref_rpr)
        fill_data(nested[idx], item['rows'], ref_rpr)

    # ---- 共计费用 ----
    if '共计费用：' in titles:
        fill_para_text(titles['共计费用：'], '共计费用：' + data['total'] + '元', ref_rpr)

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

    os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
    doc.save(args.output)
    print('已生成:', args.output)


if __name__ == '__main__':
    main()