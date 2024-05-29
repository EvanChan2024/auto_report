import taos
import pandas as pd
import numpy as np
from docx import Document
from docx.shared import Inches
from docx.table import Table
from docx.enum.text import WD_ALIGN_PARAGRAPH
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import os
'''
v1.0
单桥计算统计数据，绘制数据图并写入到word指定位置
支持应变、挠度、裂缝、桥面温度、环境温湿度
画图时y轴label及title还需对应修改
'''


def read_data(statement):
    conn: taos.TaosConnection = taos.connect(host="jkjc1",
                                             user="root",
                                             password="taosdata",
                                             database="db_jkjc",
                                             port=6030)

    result = conn.query(statement)
    result_list = result.fetch_all()  # 将查询结果转换为列表
    conn.close()
    return result_list


def data_calculate(data):
    maximum = np.nanmax(data)
    minimum = np.nanmin(data)
    mean = np.nanmean(data)
    chazhi = maximum-minimum
    result = [round(maximum, 4), round(minimum, 4), round(chazhi, 4)]
    return result


def analysis_nd(bridge_num, sensor_num, time_1, time_2, bridge_name, season_num, type):
    query_words = ('select ts,val from ' + '`' + bridge_num + '-' + sensor_num + '`' +
                   ' where ts >= ' + "'" + time_1 + "'" + ' and ts <= ' + "'" + time_2 + "'")
    data = read_data(query_words)
    time = [t[0] for t in data]
    data_nd = [t[1] for t in data]
    path = fig_plot(time, data_nd, sensor_num, bridge_name, season_num, type)
    if data_nd:
        tongjizhi = data_calculate(data_nd)
    else:
        tongjizhi = [0, 0, 0]
    return tongjizhi, path


def analysis_crk(bridge_num, sensor_num, time_1, time_2, bridge_name, season_num, type):
    query_words = ('select ts,val from ' + '`' + bridge_num + '-' + sensor_num + '`' +
                   ' where ts >= ' + "'" + time_1 + "'" + ' and ts <= ' + "'" + time_2 + "'")
    data = read_data(query_words)
    time = [t[0] for t in data]
    data_crk = [t[1] for t in data]
    path = fig_plot(time, data_crk, sensor_num, bridge_name, season_num, type)
    tongjizhi = data_calculate(data_crk)
    return tongjizhi, path


def analysis_tmp(bridge_num, sensor_num, time_1, time_2, bridge_name, season_num, type):
    query_words = ('select ts,val from ' + '`' + bridge_num + '-' + sensor_num + '`' +
                   ' where ts >= ' + "'" + time_1 + "'" + ' and ts <= ' + "'" + time_2 + "'")
    data = read_data(query_words)
    time = [t[0] for t in data]
    data_tmp = [t[1] for t in data]
    path = fig_plot(time, data_tmp, sensor_num, bridge_name, season_num, type)
    tongjizhi = data_calculate(data_tmp)
    return tongjizhi, path


def analysis_rhs(bridge_num, sensor_num, time_1, time_2, bridge_name, season_num, type_1, type_2):
    query_words = ('select ts,val1,val2 from ' + '`' + bridge_num + '-' + sensor_num + '`' +
                   ' where ts >= ' + "'" + time_1 + "'" + ' and ts <= ' + "'" + time_2 + "'")
    data = read_data(query_words)
    time = [t[0] for t in data]
    data_rts = [t[1] for t in data]
    data_rhs = [t[2] for t in data]
    path_1 = fig_plot(time, data_rts, sensor_num, bridge_name, season_num, type_1)
    path_2 = fig_plot(time, data_rhs, sensor_num, bridge_name, season_num, type_2)
    try:
        tongjizhi_1 = data_calculate(data_rts)
        tongjizhi_2 = data_calculate(data_rhs)
        return tongjizhi_1, tongjizhi_2, path_1, path_2
    except Exception as e:
        print(f'{sensor_num}: error')


def analysis_rsg(bridge_num, sensor_num, time_1, time_2, bridge_name, season_num, type_1, type_2):
    query_words = ('select ts,val1,val2,val3 from ' + '`' + bridge_num + '-' + sensor_num + '`' +
                   ' where ts >= ' + "'" + time_1 + "'" + ' and ts <= ' + "'" + time_2 + "'")
    data = read_data(query_words)
    time = [t[0] for t in data]
    data_ybw = [t[2] for t in data]
    data_yb = [t[3] for t in data]
    path_1 = fig_plot(time, data_ybw, sensor_num, bridge_name, season_num, type_1)
    path_2 = fig_plot(time, data_yb, sensor_num, bridge_name, season_num, type_2)
    if data_ybw:
        tongjizhi_1 = data_calculate(data_ybw)
    else:
        tongjizhi_1 = [0, 0, 0]
        print(f'{sensor_num}: YB data is void')
    if data_yb:
        tongjizhi_2 = data_calculate(data_yb)
    else:
        tongjizhi_2 = [0, 0, 0]
        print(f'{sensor_num}: YBw data is void')
    return tongjizhi_1, tongjizhi_2, path_1, path_2


def find_paragraph_by_keyword(doc, keyword):
    for paragraph in doc.paragraphs:
        # 查找包含指定关键词的标题段落
        if paragraph.style.name == 'Heading 2' and keyword in paragraph.text:
            return paragraph


def add_data_to_table_after_keyword(doc, keyword, data):
    # 查找包含指定关键词的标题段落
    target_paragraph = find_paragraph_by_keyword(doc, keyword)
    if target_paragraph:
        found_target = False
        for element in doc.element.body:
            if found_target:
                if element.tag.endswith('tbl'):
                    table = element
                    break
            if element.tag.endswith('p') and element.text == target_paragraph.text:
                found_target = True
        else:
            print('Target table not found')
            raise ValueError("Table not found after the specified keyword")

        # 找到表格并将数据添加到表格的后三列
        table = Table(table, doc)
        for row_index, row_data in enumerate(data, start=1):  # 从表格的第二行开始替换数据
            row_cells = table.rows[row_index].cells  # 获取各行的所有单元格
            for i, cell_data in enumerate(row_data):  # 遍历行中的每个单元格数据
                row_cells[-(len(row_data) - i)].text = str(cell_data)  # 填充单元格数据
                row_cells[-(len(row_data) - i)].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER  # 设置表格单元格中第一个段落的对齐方式为居中


def fig_plot(ts, val, sensor, bridge_name, season_num, type):
    b_with = 1  # 边框宽度
    # 字体宏观设置
    plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # 使用微软雅黑字体
    plt.rcParams['axes.unicode_minus'] = False  # 正常显示负号
    # 设置图片大小
    plt.figure(figsize=(10, 4), dpi=100)
    ax = plt.gca()  # 获取边框
    # 设置边框
    ax.spines['bottom'].set_linewidth(b_with)  # 图框下边
    ax.spines['left'].set_linewidth(b_with)  # 图框左边
    ax.spines['top'].set_linewidth(b_with)  # 图框上边
    ax.spines['right'].set_linewidth(b_with)  # 图框右边
    # 画图，添加图例
    plt.plot(ts, val)
    # plt.legend(loc="best")
    # 坐标轴标题设置
    plt.xlabel("时间", fontdict={'family': 'Microsoft YaHei', 'size': 10}, labelpad=5)
    plt.ylabel("主梁竖向位移(mm)", fontdict={'family': 'Microsoft YaHei', 'size': 10}, labelpad=5)
    # 刻度标签参数
    plt.xticks(fontproperties='Microsoft YaHei', size=10)
    plt.yticks(fontproperties='Microsoft YaHei', size=10)
    # 刻度线的大小长短粗细
    plt.tick_params(axis="both", which="major", direction="in", width=1, length=5, pad=5)
    # 刻度标签自定义
    # plt.yticks(np.arange(min(val), max(val) + 1, (max(val) - min(val)) / 10))

    # 设置日期格式
    date_format = mdates.DateFormatter('%M-%D')  # %Y-%M-%D %H:%M:%S
    ax.xaxis.set_major_formatter(date_format)

    # 标题及网格线
    plt.title(sensor + '主梁竖向位移波动趋势', fontproperties='Microsoft YaHei', size=12)
    plt.grid(True, linestyle='--', alpha=0.5)
    # plt.show()

    # 创建子文件夹
    path_1 = r'D:\Project\02\project3'  # 前缀地址，允许自定义
    # 创建目标文件夹路径
    path_destination = os.path.join(path_1, bridge_name, type, season_num)
    os.makedirs(path_destination, exist_ok=True)
    # 生成并保存图片
    file_path = os.path.join(path_destination, f'{sensor}.png')
    plt.savefig(file_path)
    plt.close()

    return path_destination


def add_picture_to_doc_after_table(doc, keyword, path):
    # 查找包含指定关键词的标题段落
    target_paragraph = find_paragraph_by_keyword(doc, keyword)
    if target_paragraph:
        found_target = False
        for element in doc.element.body:
            if found_target:
                if element.tag.endswith('tbl'):
                    table = element
                    break
            if element.tag.endswith('p') and element.text == target_paragraph.text:
                found_target = True
        else:
            raise ValueError("Table not found after the specified keyword")

        # 在表格后添加图片
        table_idx = doc.element.body.index(table)  # 获取表格在文档主体中的索引
        for i, image_file in enumerate(os.listdir(path)):
            new_paragraph = doc.add_paragraph()  # 创建一个新段落
            run = new_paragraph.add_run()  # 创建一个新的文本运行对象
            run.add_picture(os.path.join(path, image_file), width=Inches(6.0))  # 在运行对象中插入图片
            # 将新段落插入到表格后
            doc.element.body.insert(table_idx + i + 1, new_paragraph._element)  # 将新段落插入到表格后


if __name__ == "__main__":
    bridge = ['S342塘南河桥']
    season = 'Q1'
    time_start = '2024-04-01 00:00:00'
    time_end = '2024-04-30 23:59:59'
    # sensor_type = 'aaa'
    # sensor = 'hhh'
    df = pd.read_excel(r'D:\Project\02\project3\data\sensorinfo.xlsx', sheet_name='BRIDGE_TEST_SELFCHECK.T_BRIDGE')
    # 过滤数据，选取所需的列
    filtered_data = df[df['BRIDGENAME'].isin(bridge)][['FOREIGN_KEY', 'SENSOR_SUB_TYPE_NAME', 'SENSOR_CODE']]
    bridge_number = filtered_data['FOREIGN_KEY'].to_list()
    # 按照传感器类型分组
    grouped = filtered_data.groupby('SENSOR_SUB_TYPE_NAME')

    for sensor_type, group in grouped:
        data_list = []
        data_list_1 = []
        keyword = '666'
        keyword_1 = '666'
        path = '666'
        path_1 = '666'
        path_2 = '666'
        for sensor_code in group['SENSOR_CODE']:
            if sensor_type == '结构裂缝':
                data_single, path = analysis_crk(bridge_number[0], sensor_code, time_start, time_end, bridge[0], season, sensor_type)
                data_list.append(data_single)
                keyword = "混凝土结构裂缝监测"  # 指定关键词
            elif sensor_type == '桥面温度':
                data_single, path = analysis_tmp(bridge_number[0], sensor_code, time_start, time_end, bridge[0], season, sensor_type)
                data_list.append(data_single)
                keyword = "桥面铺装层温度监测"  # 指定关键词
            elif sensor_type == '主梁竖向位移':
                data_single, path = analysis_nd(bridge_number[0], sensor_code, time_start, time_end, bridge[0], season, sensor_type)
                data_list.append(data_single)
                keyword = "主梁竖向位移监测"  # 指定关键词
            elif sensor_type == '环境温湿度':
                type_1 = '环境温度'
                type_2 = '环境湿度'
                data_single_1, data_single_2, path_1, path_2 = analysis_rhs(bridge_number[0], sensor_code, time_start, time_end, bridge[0], season, type_1, type_2)
                data_list.append(data_single_1)
                data_list_1.append(data_single_2)
                keyword = "环境温度、湿度监测"  # 指定关键词
            elif sensor_type == '应变/温度':
                type_1 = '应变温度'
                type_2 = '应变'
                data_single_1, data_single_2, path_1, path_2 = analysis_rsg(bridge_number[0], sensor_code, time_start, time_end, bridge[0], season, type_1, type_2)
                data_list.append(data_single_1)
                data_list_1.append(data_single_2)
                keyword = "结构温度监测"  # 指定关键词
                keyword_1 = "关键截面应变监测"  # 指定关键词
            else:
                print(f'{sensor_code}: Not defined')

        if sensor_type == '应变/温度':
            doc = Document(r'D:\Project\02\project3\01.塘南河桥24年4月数据分析报告.docx')  # 读取 Word 文档
            add_data_to_table_after_keyword(doc, keyword, data_list)
            add_data_to_table_after_keyword(doc, keyword_1, data_list_1)
            add_picture_to_doc_after_table(doc, keyword, path_1)
            add_picture_to_doc_after_table(doc, keyword_1, path_2)
            doc.save(r'D:\Project\02\project3\01.塘南河桥24年4月数据分析报告.docx')  # 保存修改后的文档
            print(f'{sensor_type}: Solution is done')
        elif sensor_type == '环境温湿度':
            data_list.extend(data_list_1)  # 环境温湿度合并列表，其他的也适用，但应变/温度不适用
            doc = Document(r'D:\Project\02\project3\01.塘南河桥24年4月数据分析报告.docx')  # 读取 Word 文档
            add_data_to_table_after_keyword(doc, keyword, data_list)
            add_picture_to_doc_after_table(doc, keyword, path_1)
            add_picture_to_doc_after_table(doc, keyword, path_2)
            doc.save(r'D:\Project\02\project3\01.塘南河桥24年4月数据分析报告.docx')  # 保存修改后的文档
            print(f'{sensor_type}: Solution is done')
        else:
            data_list.extend(data_list_1)  # 环境温湿度合并列表，其他的也适用，但应变/温度不适用
            doc = Document(r'D:\Project\02\project3\01.塘南河桥24年4月数据分析报告.docx')  # 读取 Word 文档
            add_data_to_table_after_keyword(doc, keyword, data_list)
            add_picture_to_doc_after_table(doc, keyword, path)
            doc.save(r'D:\Project\02\project3\01.塘南河桥24年4月数据分析报告.docx')  # 保存修改后的文档
            print(f'{sensor_type}: Solution is done')


