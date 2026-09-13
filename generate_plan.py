# -*- coding: utf-8 -*-
"""
外国新闻事业史 - 42天复习计划生成器
考试日期：2026年10月25日
每天复习时间：1小时
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
wb.remove(wb.active)

# ============ 样式 ============
title_font = Font(name='微软雅黑', size=14, bold=True, color='FFFFFF')
title_fill = PatternFill(start_color='1F3864', end_color='1F3864', fill_type='solid')

header_font = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
header_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

cell_font = Font(name='微软雅黑', size=10)
cell_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

# 阶段颜色
phase1_fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')  # 蓝-基础
phase2_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')  # 绿-强化
phase3_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')  # 黄-冲刺
phase4_fill = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')  # 橙-考前

thin_border = Border(
    left=Side(style='thin', color='B4C6E7'),
    right=Side(style='thin', color='B4C6E7'),
    top=Side(style='thin', color='B4C6E7'),
    bottom=Side(style='thin', color='B4C6E7')
)

def style_header(ws, row, cols):
    for c in range(1, cols+1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

def style_data_cell(ws, row, col, fill=None):
    cell = ws.cell(row=row, column=col)
    cell.font = cell_font
    cell.border = thin_border
    if fill:
        cell.fill = fill
    if col in [1,2,3,5]:
        cell.alignment = center_align
    else:
        cell.alignment = cell_align

# ============ Sheet 1: 计划总览 ============
ws1 = wb.create_sheet('计划总览')
ws1.merge_cells('A1:F1')
c = ws1['A1']
c.value = '外国新闻事业史 · 42天复习计划（考试：2026年10月25日）'
c.font = title_font
c.fill = title_fill
c.alignment = Alignment(horizontal='center', vertical='center')
ws1.row_dimensions[1].height = 36

# 基本信息
info = [
    ['考试日期', '2026年10月25日（周六）', '剩余天数', '42天（9/13-10/24）'],
    ['每日复习', '1小时', '总复习时长', '42小时'],
    ['背诵资料', '国家时间轴背诵刷题表（12主题/537条）', '题库', '366题（单选180/多选30/名词61/简答65/论述30）'],
]
for i, row in enumerate(info, 3):
    for j, val in enumerate(row, 1):
        cell = ws1.cell(row=i, column=j, value=val)
        cell.font = cell_font
        cell.border = thin_border
        cell.alignment = cell_align
        if j in [1,3]:
            cell.font = Font(name='微软雅黑', size=10, bold=True)
            cell.fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')

# 四阶段
ws1.cell(row=7, column=1, value='四阶段复习策略').font = Font(name='微软雅黑', size=12, bold=True)
phase_headers = ['阶段', '时间', '天数', '目标', '每日时间分配', '关键产出']
for j, h in enumerate(phase_headers, 1):
    ws1.cell(row=8, column=j, value=h)
style_header(ws1, 8, 6)

phases = [
    ['第一阶段\n基础梳理', '9/13-9/26', '14天', '按国家/主题过一遍背诵表，建立整体时间轴框架', '45分钟背诵+15分钟整理笔记', '12主题全部过完，标注薄弱环节'],
    ['第二阶段\n强化刷题', '9/27-10/10', '14天', '高频考点强化记忆，系统刷完全部题库', '30分钟背诵+30分钟刷题', '366题全部刷完，建立错题本'],
    ['第三阶段\n模考冲刺', '10/11-10/20', '10天', '3套模拟卷+查漏补缺，突破薄弱环节', '45分钟模考/刷题+15分钟错题回顾', '3套模考完成，薄弱环节清零'],
    ['第四阶段\n考前调整', '10/21-10/24', '4天', '核心考点速记+错题回顾，调整考试状态', '40分钟回顾+20分钟心态调整', '核心框架烂熟于心，心态平稳'],
]
phase_fills = [phase1_fill, phase2_fill, phase3_fill, phase4_fill]
for i, (row, pf) in enumerate(zip(phases, phase_fills), 9):
    for j, val in enumerate(row, 1):
        cell = ws1.cell(row=i, column=j, value=val)
        cell.font = cell_font
        cell.border = thin_border
        cell.fill = pf
        if j in [1,2,3]:
            cell.alignment = center_align
        else:
            cell.alignment = cell_align
    ws1.row_dimensions[i].height = 50

# 列宽
widths1 = [14, 16, 8, 30, 28, 28]
for i, w in enumerate(widths1, 1):
    ws1.column_dimensions[get_column_letter(i)].width = w

# ============ Sheet 2: 逐日日历 ============
ws2 = wb.create_sheet('逐日复习日历')
ws2.merge_cells('A1:G1')
c = ws2['A1']
c.value = '42天逐日复习日历（每天1小时）'
c.font = title_font
c.fill = title_fill
c.alignment = Alignment(horizontal='center', vertical='center')
ws2.row_dimensions[1].height = 32

cal_headers = ['天数', '日期', '星期', '阶段', '背诵任务（约45分钟）', '刷题/练习任务（约15分钟）', '完成状态']
for j, h in enumerate(cal_headers, 1):
    ws2.cell(row=2, column=j, value=h)
style_header(ws2, 2, 7)

# 42天详细计划
import datetime
start = datetime.date(2026, 9, 13)
weekdays = ['周一','周二','周三','周四','周五','周六','周日']

schedule = [
    # (阶段, 背诵任务, 刷题任务)
    # 第一阶段：基础梳理 14天
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
    # 第二阶段：强化刷题 14天
    ('强化刷题', '核心考点速记：第一/最早（15条）+ 重要人物（13条）', '单选题1-40'),
    ('强化刷题', '核心考点速记：重要概念（10条）', '单选题41-80'),
    ('强化刷题', '核心考点速记：重要文件（6条）', '单选题81-120'),
    ('强化刷题', '核心考点速记：各国体制（6条）+ 速记整体回顾', '单选题121-180（单选全部完成）'),
    ('强化刷题', '英国+美国高频考点回顾', '多选题全部30题'),
    ('强化刷题', '法国+德国高频考点回顾', '名词解释1-20'),
    ('强化刷题', '日本+俄罗斯高频考点回顾', '名词解释21-40'),
    ('强化刷题', '发展中国家+通讯社高频考点回顾', '名词解释41-61（名词全部完成）'),
    ('强化刷题', '广播电视+国际传播高频考点回顾', '简答题1-20'),
    ('强化刷题', '错题回顾：单选+多选错题', '简答题21-40'),
    ('强化刷题', '错题回顾：名词解释错题', '简答题41-65（简答全部完成）'),
    ('强化刷题', '高频考点强化：人物+报刊', '论述题1-15'),
    ('强化刷题', '高频考点强化：事件+文件+体制对比', '论述题16-30（论述全部完成）'),
    ('强化刷题', '第二阶段复盘：错题本整理+薄弱环节清单', '快速过一遍全部错题'),
    # 第三阶段：模考冲刺 10天
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
    # 第四阶段：考前调整 4天
    ('考前调整', '核心考点速记全过（53条）+ 易混淆点对比记忆', '看错题本，不做新题'),
    ('考前调整', '错题本最终回顾 + 论述题素材/万能句整理', '默写3个最重要的论述题框架'),
    ('考前调整', '全真模拟（限时，用之前没做过的题组合）', '批改+心态调整，关注时间分配'),
    ('考前调整', '轻松回顾核心框架（12主题时间轴）+ 早睡', '不熬夜，准备考试用品'),
]

for i, (phase, recite, practice) in enumerate(schedule):
    row = i + 3
    d = start + datetime.timedelta(days=i)
    ws2.cell(row=row, column=1, value=f'Day {i+1}')
    ws2.cell(row=row, column=2, value=d.strftime('%m/%d'))
    ws2.cell(row=row, column=3, value=weekdays[d.weekday()])
    ws2.cell(row=row, column=4, value=phase)
    ws2.cell(row=row, column=5, value=recite)
    ws2.cell(row=row, column=6, value=practice)
    ws2.cell(row=row, column=7, value='□ 未完成')
    
    pf = phase1_fill if phase=='基础梳理' else phase2_fill if phase=='强化刷题' else phase3_fill if phase=='模考冲刺' else phase4_fill
    for col in range(1, 8):
        style_data_cell(ws2, row, col, pf)
    ws2.row_dimensions[row].height = 42

# 考试日
exam_row = len(schedule) + 3
ws2.cell(row=exam_row, column=1, value='考试日')
ws2.cell(row=exam_row, column=2, value='10/25')
ws2.cell(row=exam_row, column=3, value='周六')
ws2.cell(row=exam_row, column=4, value='考试')
ws2.cell(row=exam_row, column=5, value='外国新闻事业史考试')
ws2.cell(row=exam_row, column=6, value='—')
ws2.cell(row=exam_row, column=7, value='加油！')
for col in range(1, 8):
    cell = ws2.cell(row=exam_row, column=col)
    cell.font = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
    cell.fill = PatternFill(start_color='C00000', end_color='C00000', fill_type='solid')
    cell.alignment = center_align
    cell.border = thin_border
ws2.row_dimensions[exam_row].height = 36

widths2 = [8, 8, 6, 10, 42, 36, 10]
for i, w in enumerate(widths2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w
ws2.freeze_panes = 'A3'

# ============ Sheet 3: 每日任务模板 ============
ws3 = wb.create_sheet('每日任务模板')
ws3.merge_cells('A1:D1')
c = ws3['A1']
c.value = '每日复习任务记录模板（可打印/复制使用）'
c.font = title_font
c.fill = title_fill
c.alignment = Alignment(horizontal='center', vertical='center')
ws3.row_dimensions[1].height = 32

template_items = [
    ['日期', '____年____月____日（星期____）', '第几天', 'Day ____'],
    ['今日阶段', '□基础梳理 □强化刷题 □模考冲刺 □考前调整', '复习时长', '____分钟'],
    ['', '', '', ''],
    ['【背诵任务】', '', '', ''],
    ['背诵内容', '', '用时', '____分钟'],
    ['已掌握（打√）', '', '不熟悉（标记）', ''],
    ['', '', '', ''],
    ['【刷题任务】', '', '', ''],
    ['题目范围', '', '用时', '____分钟'],
    ['正确率', '____/____ = ____%', '错题编号', ''],
    ['', '', '', ''],
    ['【错题记录】', '', '', ''],
    ['错题1', '题目：____\n错因：□知识盲区 □记忆混淆 □审题失误 □其他\n正确答案：____\n相关知识点：____', '', ''],
    ['错题2', '题目：____\n错因：□知识盲区 □记忆混淆 □审题失误 □其他\n正确答案：____\n相关知识点：____', '', ''],
    ['错题3', '题目：____\n错因：□知识盲区 □记忆混淆 □审题失误 □其他\n正确答案：____\n相关知识点：____', '', ''],
    ['', '', '', ''],
    ['【今日复盘】', '', '', ''],
    ['最大收获', '', '明日重点', ''],
    ['心态评分', '□很好 □一般 □焦虑', '完成度', '□全部完成 □部分完成 □未完成'],
]

for i, row in enumerate(template_items, 3):
    for j, val in enumerate(row, 1):
        cell = ws3.cell(row=i, column=j, value=val)
        cell.font = cell_font
        cell.border = thin_border
        cell.alignment = cell_align
        if val and val.startswith('【'):
            cell.font = Font(name='微软雅黑', size=11, bold=True, color='1F3864')
            cell.fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
        if j == 1 and val and not val.startswith('【'):
            cell.font = Font(name='微软雅黑', size=10, bold=True)
    ws3.row_dimensions[i].height = 28

# 合并一些单元格
for r in [6, 10, 14, 15, 16, 17, 20]:
    ws3.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)

widths3 = [14, 40, 14, 30]
for i, w in enumerate(widths3, 1):
    ws3.column_dimensions[get_column_letter(i)].width = w

# ============ Sheet 4: 错题复盘表 ============
ws4 = wb.create_sheet('错题复盘表')
ws4.merge_cells('A1:H1')
c = ws4['A1']
c.value = '错题复盘记录表'
c.font = title_font
c.fill = title_fill
c.alignment = Alignment(horizontal='center', vertical='center')
ws4.row_dimensions[1].height = 32

err_headers = ['序号', '日期', '题型', '题目摘要', '错误答案', '正确答案', '错因分类', '相关知识点/国家']
for j, h in enumerate(err_headers, 1):
    ws4.cell(row=2, column=j, value=h)
style_header(ws4, 2, 8)

# 预填30行空行
for i in range(3, 33):
    for j in range(1, 9):
        cell = ws4.cell(row=i, column=j, value='')
        cell.font = cell_font
        cell.border = thin_border
        cell.alignment = cell_align
    ws4.row_dimensions[i].height = 30

# 错因分类说明
ws4.cell(row=34, column=1, value='错因分类说明：').font = Font(name='微软雅黑', size=10, bold=True)
ws4.cell(row=35, column=1, value='A=知识盲区（完全没记住）  B=记忆混淆（记混了相似知识点）  C=审题失误（看错题目/选项）  D=理解偏差（概念理解有误）  E=其他')
ws4.merge_cells('A35:H35')
ws4.cell(row=35, column=1).font = cell_font
ws4.cell(row=35, column=1).alignment = cell_align

widths4 = [6, 10, 8, 35, 12, 12, 10, 18]
for i, w in enumerate(widths4, 1):
    ws4.column_dimensions[get_column_letter(i)].width = w
ws4.freeze_panes = 'A3'

# ============ Sheet 5: 阶段检测表 ============
ws5 = wb.create_sheet('阶段检测表')
ws5.merge_cells('A1:E1')
c = ws5['A1']
c.value = '阶段检测与动态调参记录'
c.font = title_font
c.fill = title_fill
c.alignment = Alignment(horizontal='center', vertical='center')
ws5.row_dimensions[1].height = 32

check_headers = ['检测节点', '检测内容', '检测结果', '薄弱环节', '调整措施']
for j, h in enumerate(check_headers, 1):
    ws5.cell(row=2, column=j, value=h)
style_header(ws5, 2, 5)

checks = [
    ['第一阶段末\n（9/26）', '12主题能否复述时间轴？\n随机抽20条能否答对？\n薄弱主题是哪些？', '', '', ''],
    ['第二阶段末\n（10/10）', '题库366题是否全部刷完？\n单选/多选正确率？\n名词/简答/论述能否写出要点？\n错题本是否整理？', '', '', ''],
    ['模考1后\n（10/12）', '模考1总分/正确率？\n时间分配是否合理？\n哪类题型失分最多？', '', '', ''],
    ['模考2后\n（10/15）', '模考2是否比模考1有进步？\n薄弱环节是否改善？', '', '', ''],
    ['模考3后\n（10/18）', '模考3是否达到目标分数？\n还有哪些硬伤？', '', '', ''],
    ['考前3天\n（10/22）', '核心考点53条是否全部掌握？\n错题本是否清零？\n论述题框架是否熟练？', '', '', ''],
]

for i, row in enumerate(checks, 3):
    for j, val in enumerate(row, 1):
        cell = ws5.cell(row=i, column=j, value=val)
        cell.font = cell_font
        cell.border = thin_border
        cell.alignment = cell_align
        if j == 1:
            cell.font = Font(name='微软雅黑', size=10, bold=True)
            cell.fill = phase1_fill if i==3 else phase2_fill if i==4 else phase3_fill if i<=6 else phase4_fill
    ws5.row_dimensions[i].height = 60

widths5 = [14, 35, 20, 20, 25]
for i, w in enumerate(widths5, 1):
    ws5.column_dimensions[get_column_letter(i)].width = w

# ============ 保存 ============
output_path = r'D:\Desktop\外国新闻事业史1\外国新闻事业史_42天复习计划.xlsx'
wb.save(output_path)
print(f'复习计划已保存: {output_path}')
print(f'工作表: {wb.sheetnames}')
