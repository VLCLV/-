# -*- coding: utf-8 -*-
"""从Excel提取所有数据，生成data.js"""
import openpyxl
import json

app_dir = r'D:\Desktop\外国新闻事业史1\news-history-app'

# ===== 1. 提取背诵资料 =====
wb1 = openpyxl.load_workbook(r'D:\Desktop\外国新闻事业史1\外国新闻事业史_国家时间轴背诵刷题表.xlsx')
materials = {}
for sheet_name in wb1.sheetnames:
    ws = wb1[sheet_name]
    items = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        items.append({
            'time': str(row[0]) if row[0] else '',
            'event': str(row[1]) if row[1] else '',
            'keypoint': str(row[2]) if row[2] else '',
            'significance': str(row[3]) if row[3] else ''
        })
    materials[sheet_name] = items

# ===== 2. 提取题库 =====
wb2 = openpyxl.load_workbook(r'D:\Desktop\外国新闻事业史1\外国新闻事业史题库.xlsx')
questions = {'single': [], 'multiple': [], 'term': [], 'short': [], 'essay': []}

# 单选题
ws = wb2['单选题']
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        continue
    q = {
        'id': int(row[0]),
        'question': str(row[2]) if row[2] else '',
        'options': [str(row[3]) if row[3] else '', str(row[4]) if row[4] else '', str(row[5]) if row[5] else '', str(row[6]) if row[6] else ''],
        'answer': str(row[7]) if len(row) > 7 and row[7] else ''
    }
    questions['single'].append(q)

# 多选题
ws = wb2['多选题']
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        continue
    q = {
        'id': int(row[0]),
        'question': str(row[2]) if row[2] else '',
        'options': [str(row[3]) if row[3] else '', str(row[4]) if row[4] else '', str(row[5]) if row[5] else '', str(row[6]) if row[6] else ''],
        'answer': str(row[7]) if len(row) > 7 and row[7] else ''
    }
    questions['multiple'].append(q)

# 名词解释
ws = wb2['名词解释']
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        continue
    questions['term'].append({
        'id': int(row[0]),
        'question': str(row[2]) if row[2] else '',
        'answer': str(row[3]) if len(row) > 3 and row[3] else ''
    })

# 简答题
ws = wb2['简答题']
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        continue
    questions['short'].append({
        'id': int(row[0]),
        'question': str(row[2]) if row[2] else '',
        'answer': str(row[3]) if len(row) > 3 and row[3] else ''
    })

# 论述题
ws = wb2['论述题']
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        continue
    questions['essay'].append({
        'id': int(row[0]),
        'question': str(row[2]) if row[2] else '',
        'answer': str(row[3]) if len(row) > 3 and row[3] else ''
    })

# ===== 3. 42天复习计划 =====
import datetime
start = datetime.date(2026, 9, 13)
weekdays_cn = ['周一','周二','周三','周四','周五','周六','周日']

schedule = [
    ('基础梳理', '总览-起源与脉络（21条）+ 英国上（1528-1922，17条）', '整理英国时间轴，标注3个最重要节点'),
    ('基础梳理', '英国下（1922-当代，16条）', '整理英国报业发展脉络图（特许制→革命→廉价报→报团）'),
    ('基础梳理', '美国上（1704-1920，24条）', '重点记忆：曾格案、便士报三报、黄色新闻'),
    ('基础梳理', '美国下（1920-当代，23条）', '整理美国主要报纸/杂志/通讯社/广电清单'),
    ('基础梳理', '法国上（1474-1871，23条）', '重点记忆：《人权宣言》、《新闻出版自由法》、哈瓦斯社'),
    ('基础梳理', '法国下（1871-当代，23条）', '整理法国报业/广电体制特点'),
    ('基础梳理', '德国上（1450s-1918，26条）', '重点记忆：《新莱茵报》、马克思恩格斯新闻实践'),
    ('基础梳理', '德国下（1918-当代，25条）', '重点记忆：纳粹新闻统制、战后重建、《图片报》'),
    ('基础梳理', '日本上（1615-1924，26条）', '重点记忆：五大报纸发展史、《朝日》《读卖》创刊'),
    ('基础梳理', '日本下（1924-当代，25条）', '重点记忆：战时统制、战后改革、NHK体制'),
    ('基础梳理', '俄罗斯-苏联上（17世纪前-1922，35条）', '重点记忆：《火星报》《真理报》、列宁新闻思想'),
    ('基础梳理', '俄罗斯-苏联下（1922-当代，34条）', '整理苏联新闻体制特点、解体后转型'),
    ('基础梳理', '发展中国家（35条）+ 通讯社事业（28条）', '重点记忆：三社四边协定、四大通讯社'),
    ('基础梳理', '广播电视（57条快速过）+ 国际传播（46条快速过）+ 第一阶段复盘', '标注第一阶段薄弱环节，列出清单'),
    ('强化刷题', '核心考点速记：第一/最早（15条）+ 重要人物（13条）', '单选题1-40'),
    ('强化刷题', '核心考点速记：重要概念（10条）', '单选题41-80'),
    ('强化刷题', '核心考点速记：重要文件（6条）', '单选题81-120'),
    ('强化刷题', '核心考点速记：各国体制（6条）+ 速记整体回顾', '单选题121-180'),
    ('强化刷题', '英国+美国高频考点回顾', '多选题全部30题'),
    ('强化刷题', '法国+德国高频考点回顾', '名词解释1-20'),
    ('强化刷题', '日本+俄罗斯高频考点回顾', '名词解释21-40'),
    ('强化刷题', '发展中国家+通讯社高频考点回顾', '名词解释41-61'),
    ('强化刷题', '广播电视+国际传播高频考点回顾', '简答题1-20'),
    ('强化刷题', '错题回顾：单选+多选错题', '简答题21-40'),
    ('强化刷题', '错题回顾：名词解释错题', '简答题41-65'),
    ('强化刷题', '高频考点强化：人物+报刊', '论述题1-15'),
    ('强化刷题', '高频考点强化：事件+文件+体制对比', '论述题16-30'),
    ('强化刷题', '第二阶段复盘：错题本整理+薄弱环节清单', '快速过一遍全部错题'),
    ('模考冲刺', '模拟卷1（限时45分钟：单选30+多选5+名词3+简答2+论述1）', '批改+错题分析，标注薄弱知识点'),
    ('模考冲刺', '模拟卷1错题相关知识点深度回顾', '针对薄弱点刷题15题'),
    ('模考冲刺', '薄弱环节突破1（根据模考结果定）', '相关题目二刷'),
    ('模考冲刺', '模拟卷2（限时45分钟）', '批改+错题分析'),
    ('模考冲刺', '模拟卷2错题相关知识点深度回顾', '针对薄弱点刷题15题'),
    ('模考冲刺', '薄弱环节突破2', '相关题目二刷'),
    ('模考冲刺', '模拟卷3（限时45分钟）', '批改+错题分析'),
    ('模考冲刺', '模拟卷3错题相关知识点深度回顾', '三套卷错题汇总'),
    ('模考冲刺', '全部错题本最终回顾（单选+多选+名词）', '高频考点再过一遍'),
    ('模考冲刺', '全部错题本最终回顾（简答+论述）+ 第三阶段复盘', '论述题答题框架整理'),
    ('考前调整', '核心考点速记全过（53条）+ 易混淆点对比记忆', '看错题本，不做新题'),
    ('考前调整', '错题本最终回顾 + 论述题素材/万能句整理', '默写3个最重要的论述题框架'),
    ('考前调整', '全真模拟（限时，用之前没做过的题组合）', '批改+心态调整，关注时间分配'),
    ('考前调整', '轻松回顾核心框架（12主题时间轴）+ 早睡', '不熬夜，准备考试用品'),
]

plan = []
for i, (phase, recite, practice) in enumerate(schedule):
    d = start + datetime.timedelta(days=i)
    plan.append({
        'day': i + 1,
        'date': d.strftime('%m/%d'),
        'weekday': weekdays_cn[d.weekday()],
        'phase': phase,
        'recite': recite,
        'practice': practice,
        'completed': False
    })

# ===== 生成data.js =====
data = {
    'materials': materials,
    'questions': questions,
    'plan': plan,
    'examDate': '2026-10-25',
    'materialCategories': list(materials.keys()),
    'stats': {
        'totalMaterials': sum(len(v) for v in materials.values()),
        'totalQuestions': sum(len(v) for v in questions.values()),
        'totalDays': 42
    }
}

js_content = '// 外国新闻事业史复习应用 - 数据文件\n'
js_content += '// 自动生成，请勿手动修改\n\n'
js_content += 'const APP_DATA = ' + json.dumps(data, ensure_ascii=False, indent=2) + ';\n'

with open(app_dir + r'\js\data.js', 'w', encoding='utf-8') as f:
    f.write(js_content)

print(f'data.js 生成完成')
print(f'  背诵资料: {data["stats"]["totalMaterials"]} 条')
print(f'  题库: {data["stats"]["totalQuestions"]} 题')
print(f'    单选: {len(questions["single"])}')
print(f'    多选: {len(questions["multiple"])}')
print(f'    名词解释: {len(questions["term"])}')
print(f'    简答: {len(questions["short"])}')
print(f'    论述: {len(questions["essay"])}')
print(f'  复习计划: {len(plan)} 天')
