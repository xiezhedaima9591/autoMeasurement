#-*- coding:utf-8 -*-

import re
import xlsxwriter
from datetime import datetime
from pointer import Pointer

class Maker(object):
    '结果文件生成器'

    def __init__(self, counter):
        self.ct = counter #统计器
        result_file_name = counter.bug_count_dict['subject_name'] + 'BugReport.xlsx'
        self.workbook = xlsxwriter.Workbook(result_file_name) #工作簿
        self.bc_ws = self.workbook.add_worksheet('bug统计') #bug统计工作表
        self.bc_p = Pointer(0, 0) #bug统计工作表行列指示器
        self.ur_ws = self.workbook.add_worksheet('待解决bug列表-负责小组') #待解决bug工作表按负责小组分组
        self.ur_p = Pointer(0, 0) #待解决bug工作表行列指示器
        if self.ct.config.classify_switch == True:
            self.urc_ws = self.workbook.add_worksheet('待解决bug列表-处理状态') #待解决bug工作表按处理状态分组
            self.urc_p = Pointer(0, 0)
        self.nb_ws = self.workbook.add_worksheet('新增bug列表') #新增bug工作表
        self.nb_p = Pointer(0, 0) #新增bug工作表行列指示器
        self.rs_ws = self.workbook.add_worksheet('已解决bug列表') #已解决bug工作表
        self.rs_p = Pointer(0, 0) #已解决bug工作表行列指示器
        self.vd_ws = self.workbook.add_worksheet('待验证bug列表') #待验证bug工作表
        self.vd_p = Pointer(0, 0) #待验证bug工作表行列指示器
        if self.ct.config.delay_switch == True:
            self.dl_ws = self.workbook.add_worksheet('延期bug列表') #延期bug工作表
            self.dl_p = Pointer(0, 0)
        self.da_ws = self.workbook.add_worksheet('数据处理') #数据处理工作表
        self.da_p = Pointer(0, 0) #数据处理工作表行列指示器
        self.fd = {}
        self._make_format()

    def _make_format(self):
        '生成所需的所有格式'
        self.fd['title_format'] = self.workbook.add_format({
            'font_name': '微软雅黑',
            'font_size': 20,
            'align': 'center',
            'valign': 'vcenter',
            })
        self.fd['date_format'] = self.workbook.add_format({
            'font_name': '微软雅黑',
            'font_size': 11,
            'align': 'right',
            'valign': 'bottom',
            })
        self.fd['seg_title_format'] = self.workbook.add_format({
            'font_name': '微软雅黑',
            'font_size': 14,
            'bold': True,
            'bg_color': '#8DB4E2',
            })
        self.fd['thead_format'] = self.workbook.add_format({
            'font_name': '微软雅黑',
            'font_size': 10,
            'font_color': 'white',
            'bg_color': '#DA9694',
            'border': True,
            'border_color': 'black',
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            })
        self.fd['thead_format'].set_text_wrap()
        self.fd['tbody_format'] = self.workbook.add_format({
            'font_name': '微软雅黑',
            'font_size': 11,
            'border': True,
            'border_color': 'black',
            'align': 'center',
            'valign': 'vcenter',
            })
        self.fd['cl_format'] = self.workbook.add_format({
            'diag_type': 2,
            'font_name': '微软雅黑',
            'font_size': 10,
            'font_color': 'white',
            'bg_color': 'DA9694',
            'border': True,
            'border_color': 'black',
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            })
        self.fd['cl_format'].set_text_wrap()
        self.fd['lh_format'] = self.workbook.add_format({
            'font_name': '微软雅黑',
            'font_size': 10,
            'top': True,
            })
        self.fd['list_format'] = self.workbook.add_format({
            'font_name': '微软雅黑',
            'font_size': 10,
            })
        self.fd['data_date'] = self.workbook.add_format()
        self.fd['data_date'].set_num_format('m/d')

    def _sort_unresolved_data(self):
        '对未解决bug表的数据排序'
        unresolved_index = 5
        temp_dict = {}
        temp_sorted = []
        sorted_groups = []
        groups = self.ct.unresolved_result.keys()
        
        for group in groups:
            temp_dict[group] = self.ct.unresolved_result[group][unresolved_index]

        temp_sorted = sorted(temp_dict.iteritems(), key = lambda x: x[1], reverse = True)
        total = temp_sorted.pop(0)
        for one_item in temp_sorted:
            sorted_groups.append(one_item[0])
        sorted_groups.append(total[0])

        return sorted_groups

    def _read_report(self):
        '读取趋势图数据'
        record_file_name = self.ct.bug_count_dict['subject_name'] + '.txt'
        record_file = open(record_file_name, 'r')
        record_str = record_file.read()
        record_file.close()
        record_matrix, item_length = self.ct.str_to_matrix(record_str)
        for item_index, one_item in enumerate(record_matrix):
            if one_item[1].isdigit():
                one_item[0] = datetime.strptime(one_item[0], '%Y-%m-%d').date()
                if item_index == 1:
                    min_date = one_item[0]
                if item_index == item_length - 1:
                    max_date = one_item[0]
                one_item[1] = int(one_item[1])
                one_item[2] = int(one_item[2])
                one_item[3] = int(one_item[3])
                one_item[4] = int(one_item[4])
            self.da_ws.write(self.da_p.cu_row, 0, one_item[0], self.fd['data_date'])
            self.da_ws.write(self.da_p.cu_row, 1, one_item[1])
            self.da_ws.write(self.da_p.cu_row, 2, one_item[2])
            self.da_ws.write(self.da_p.cu_row, 3, one_item[3])
            self.da_ws.write(self.da_p.cu_row, 4, one_item[4])
            self.da_p.next_row()

        self.da_ws.hide()

        return item_length, min_date, max_date

    def make_bug_count_sheet(self):
        '生成bug统计表'
        today = datetime.today()
        title = self.ct.bug_count_dict['subject_name'] + 'Bug统计'
        self.bc_ws.set_column(2, 2, 13.14)
        self.bc_ws.set_column('O:O', 22.43)
        self.bc_ws.set_column(0, 0, 24.38)
        self.bc_ws.set_column('B:I', 7)
        self.bc_ws.set_row(0, 30.75, {})
        self.bc_ws.merge_range(self.bc_p.cu_row, self.bc_p.cu_col,
                self.bc_p.cu_row, self.bc_p.cu_col + 
                self.bc_p.max_col - 2, title, self.fd['title_format'])
        self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col + 
                self.bc_p.max_col - 1, str(today.date()),
                self.fd['date_format'])
        self._make_all_bug() #画bug总况表
        self._make_unresolved_bug() #画待解决bug统计表

    def _make_show_name(self):
        ss_name = ''
        temp_dict = {}
        for lintxx in self.ct.lint:
            if lintxx['产品名称'] not in temp_dict:
                temp_dict[lintxx['产品名称']] = 1
            else:
                temp_dict[lintxx['产品名称']] += 1
        keys = temp_dict.keys()
        ss_name = keys[0]
        for key in keys:
            if temp_dict[key] > temp_dict[ss_name]:
                ss_name = key

        return ss_name

    def _make_all_bug(self):
        '画bug总况表'
        ss_name = self._make_show_name()  #ss_name存储项目展示名称
        self.bc_p.next_row()
        self.bc_ws.merge_range(self.bc_p.cu_row, self.bc_p.cu_col,
                self.bc_p.cu_row, self.bc_p.cu_col +
                self.bc_p.max_col - 1, '一 Bug总况',
                self.fd['seg_title_format'])
        self.bc_p.next_row()
        self.bc_ws.merge_range(self.bc_p.cu_row, self.bc_p.cu_col,
                self.bc_p.cu_row + 1, self.bc_p.cu_col, '项目名称',
                self.fd['thead_format'])
        self.bc_p.next_col()
        self.bc_ws.merge_range(self.bc_p.cu_row, self.bc_p.cu_col,
                self.bc_p.cu_row, self.bc_p.cu_col + 2, 'bug状态',
                self.fd['thead_format'])
        self.bc_ws.set_row(self.bc_p.cu_row, 27)
        self.bc_p.set_col(self.bc_p.cu_col + 3) #因为前面将3格合并，所以下一个标题要将列数加3
        self.bc_ws.merge_range(self.bc_p.cu_row, self.bc_p.cu_col,
                self.bc_p.cu_row + 1, self.bc_p.cu_col, '合计',
                self.fd['thead_format'])
        self.bc_p.next_col()
        self._add_unclosed_chart()# 添加总体未关闭bug分布图
        self.bc_p.next_row()
        self.bc_p.set_col(1)
        temp_head = ['待解决', '待验证', '延期']
        for one_head in temp_head:
            self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                             one_head, self.fd['thead_format'])
            self.bc_p.next_col()
        self.bc_ws.set_row(self.bc_p.cu_row, 47.25)
        self.bc_p.next_row()
        self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                         ss_name, self.fd['tbody_format'])
        self.bc_p.next_col()
        self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                         self.ct.bug_count_dict['unresolved_num'],
                         self.fd['tbody_format'])
        self.bc_p.next_col()
        self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                         self.ct.bug_count_dict['validate_num'],
                         self.fd['tbody_format'])
        self.bc_p.next_col()
        self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                         self.ct.bug_count_dict['delay_num'],
                         self.fd['tbody_format'])
        self.bc_p.next_col()
        self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                         self.ct.bug_count_dict['sum_num'],
                         self.fd['tbody_format'])
        self.bc_ws.set_row(self.bc_p.cu_row, 40.5)
        self.bc_p.next_row()
        self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                         '新增bug数', self.fd['list_format'])
        self.bc_p.next_col()
        self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                         self.ct.bug_count_dict['new_bug'],
                         self.fd['list_format'])
        self.bc_p.next_row()
        self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                         '解决bug数', self.fd['list_format'])
        self.bc_p.next_col()
        self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                         self.ct.bug_count_dict['resolved_bug'],
                         self.fd['list_format'])
        self.bc_p.next_row()

    def _add_unclosed_chart(self):
        item_length, min_date, max_date = self._read_report()
        first_row = 2
        last_row = item_length
        first_str = [
            '=数据处理!$B$1',
            '=数据处理!$A$%s:$A$%s' % (first_row, last_row),
            '=数据处理!$B$%s:$B$%s' % (first_row, last_row)
        ]
        second_str = [
            '=数据处理!$C$1',
            '=数据处理!$A$%s:$A$%s' % (first_row, last_row),
            '=数据处理!$C$%s:$C$%s' % (first_row, last_row)
        ]
        third_str = [
            '=数据处理!$D$1',
            '=数据处理!$A$%s:$A$%s' % (first_row, last_row),
            '=数据处理!$D$%s:$D$%s' % (first_row, last_row),
        ]
        fourth_str = [
            '=数据处理!$E$1',
            '=数据处理!$A$%s:$A$%s' % (first_row, last_row),
            '=数据处理!$E$%s:$E$%s' % (first_row, last_row)
        ]
        unclosed_chart = self.workbook.add_chart({
            'type' : 'scatter',
            'subtype' : 'straight_with_markers'
        })
        unclosed_chart.set_size({
            'width' : 700,
            'height' : 192
        })

        unclosed_chart.add_series({
            'name' : first_str[0],
            'categories' : first_str[1],
            'values' : first_str[2],
            'data_labels': {'value': True},
        })
        unclosed_chart.add_series({
            'name' : second_str[0],
            'categories' : second_str[1],
            'values' : second_str[2],
        })
        unclosed_chart.add_series({
            'name' : third_str[0],
            'categories' : third_str[1],
            'values' : third_str[2],
        })
        unclosed_chart.add_series({
            'name' : fourth_str[0],
            'categories' : fourth_str[1],
            'values' : fourth_str[2],
        })
        unclosed_chart.set_title({
            'name' : 'bug动态',
            'name_font' : {'name' : '微软雅黑', 'size': 12},
        })
        unclosed_chart.set_x_axis({
            'num_font' : {'name' : '微软雅黑'},
            'date_axis' : True,
            'min' : min_date,
            'max' : max_date,
            'major_unit' : self.ct.config.circle,
            'major_unit_type' : 'days',
        })
        unclosed_chart.set_y_axis({
            'num_font' : {'name' : '微软雅黑'},
        })
        unclosed_chart.set_legend({'position': 'bottom'})
        self.bc_ws.insert_chart(self.bc_p.cu_row, self.bc_p.cu_col,
                                unclosed_chart)

    def _make_unresolved_bug(self):
        self.bc_ws.merge_range(self.bc_p.cu_row, self.bc_p.cu_col,
                               self.bc_p.cu_row, self.bc_p.cu_col +
                               self.bc_p.max_col - 1, '二 待解决Bug统计',
                               self.fd['seg_title_format'])
        self.bc_p.next_row()

        head_list = ['         bug严重性 \n 负责小组', '1-致命', '2-严重', '3-普通',
                     '4-较低', '5-建议', '未解决\nbug数', '新增','解决']
        for one_head in head_list:
            self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                             one_head, self.fd['thead_format'])
            self.bc_p.next_col()
        temp_row = self.bc_p.cu_row #将当前位置暂时存储，用于设置图表插入位置
        temp_col = self.bc_p.cu_col #同上
        self.bc_p.next_row()
        group_list = self._sort_unresolved_data()
        length_data = len(group_list)
        for one_group in group_list:
            self.bc_ws.set_row(self.bc_p.cu_row, 18)
            pond = self.ct.unresolved_result[one_group]
            self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                             one_group, self.fd['tbody_format'])
            for item in pond:
                self.bc_p.next_col()
                self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                                 item, self.fd['tbody_format'])
            self.bc_p.next_row()
        self._make_data_source()
        self.bc_p.set_row(temp_row)
        self.bc_p.set_col(temp_col)
        self._add_severity_chart(length_data)  #添加bug严重性分布图
        self._add_unresolved_chart(length_data)  # 添加待解决bug小组分布图

    def _make_data_source(self):
        self.fd['sec_title'] = self.workbook.add_format({
            'font_name': '微软雅黑',
            'font_size': '10',
            'bold': True,
            'border': False,
        })
        self.bc_ws.merge_range(self.bc_p.cu_row, self.bc_p.cu_col,
                               self.bc_p.cu_row, self.bc_p.cu_col +
                               self.bc_p.max_col - 1,'数据来源说明',
                               self.fd['seg_title_format'])
        self.bc_p.next_row()
        self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                         '数据来源:PLM系统导出', self.fd['sec_title'])
        self.bc_p.next_row()
        self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                         '筛选条件:' + self.ct.config.filter_term,
                         self.fd['sec_title'])
        self.bc_p.next_row()
        self.bc_ws.write(self.bc_p.cu_row, self.bc_p.cu_col,
                         '导出时间:' + datetime.today().strftime('%Y-%m-%d %H:%M'),
                         self.fd['sec_title'])

    def _add_unresolved_chart(self, length_data):
        insert_row = self.bc_p.cu_row + 7 #图插入位置行数
        insert_col = self.bc_p.cu_col #图插入位置列数
        first_row = self.bc_p.cu_row #待解决bug统计第一行
        last_row = first_row + length_data - 1 #减1是因为group_list里有总计行
        data_row = self.bc_p.cu_row + 1 #去掉标题行的有效数据第一行
        height = length_data * 18 + 33 #用于让图表自适应左边表的高度
        unresolved_chart = self.workbook.add_chart({
            'type' : 'bar',
            'subtype' : 'stacked',
        })
        for cu_col in range(1,6):
            unresolved_chart.add_series({
                'name' : ['bug统计', first_row, cu_col],
                'categories' : ['bug统计', data_row, 0, last_row, 0],
                'values' : ['bug统计', data_row, cu_col, last_row, cu_col],
            })
        unresolved_chart.set_title({
            'name' : '待解决bug负责小组分布',
            'name_font' : {'name' : '微软雅黑', 'size': 12},
        })
        unresolved_chart.set_x_axis({
            'name' : 'bug数量',
            'name_font' : {'name' : '微软雅黑'},
            'num_font' : {'name' : '微软雅黑'},
        })
        unresolved_chart.set_y_axis({
            'name' : '负责人',
            'name_font' : {'name' : '微软雅黑'},
            'num_font' : {'name' : '微软雅黑'},
            'reverse' : True,
        })
        unresolved_chart.set_size({
            'width' : 483,
            'height' : height,
            'y_scale' : 1.3,
        })
        self.bc_ws.insert_chart(insert_row, insert_col, unresolved_chart)

    def _add_severity_chart(self, length_data):
        insert_row = self.bc_p.cu_row
        insert_col = self.bc_p.cu_col
        categories_row = self.bc_p.cu_row
        value_row = categories_row + length_data
        height = 5 * 18 + 33
        severity_chart = self.workbook.add_chart({'type': 'pie'})
        severity_chart.add_series({
            'name': 'bug严重性分布',
            'categories': ['bug统计', categories_row, 1, categories_row, 5],
            'values': ['bug统计', value_row, 1, value_row, 5],
            'data_labels': {'value': True},
        })
        severity_chart.set_title({
            'name': '未解决bug严重性分布图',
            'name_font': {'name' : '微软雅黑', 'size': 12},
        })
        severity_chart.set_size({
            'width': 483,
            'height': height,
            'y_scale': 1.3,
        })
        self.bc_ws.insert_chart(insert_row, insert_col, severity_chart)


    def make_unresolved_bug_sheet(self):
        data = self.ct.make_write_dict('unresolved')
        unresolved_thead = ['负责小组', 'PSR编号', '标题', '状态', '严重性',
                            '创建时间', '出现频率', '当前审阅者', 'Reject次数', '是否重新打开']
        self.ur_ws.set_column('A:A', 24.38)
        self.ur_ws.set_column('B:B', 13.25)
        self.ur_ws.set_column('C:C', 60)
        self.ur_ws.set_column('F:F', 10.38)
        self.ur_ws.merge_range(self.ur_p.cu_row, self.ur_p.cu_col,
                               self.ur_p.cu_row,
                               self.ur_p.cu_col + self.ur_p.max_col - 1,
                               '三 待解决Bug列表-按负责小组', self.fd['seg_title_format'])
        self.ur_p.next_row()
        self.ur_ws.write_row(self.ur_p.cu_row, self.ur_p.cu_col,
                             unresolved_thead, self.fd['lh_format'])
        self.ur_p.next_row()
        groups = data.keys()
        groups = sorted(groups)
        unresolved_thead.pop(0)
        for group in groups:
            if len(data[group]) != 0:
                self.ur_ws.write(self.ur_p.cu_row, self.ur_p.cu_col,
                                 group, self.fd['list_format'])
                self.ur_p.next_col()
                for item in data[group]:
                    for head in unresolved_thead:
                        self.ur_ws.write(self.ur_p.cu_row, self.ur_p.cu_col,
                                         item[head], self.fd['list_format'])
                        self.ur_p.next_col()
                        if head == '是否重新打开':
                            self.ur_p.next_row()
                            self.ur_p.set_col(1)
                self.ur_p.set_col(0)

    def make_cunresolved_bug_sheet(self):
        data = self.ct.make_classify_dict()
        unresolved_thead = ['处理状态', 'PSR编号', '标题', '状态', '严重性',
                            '创建时间', '出现频率', '当前审阅者', 'Reject次数', '是否重新打开']
        self.urc_ws.set_column('A:A', 24.38)
        self.urc_ws.set_column('B:B', 13.25)
        self.urc_ws.set_column('C:C', 60)
        self.urc_ws.set_column('F:F', 10.38)
        self.urc_ws.merge_range(self.urc_p.cu_row, self.urc_p.cu_col,
                               self.urc_p.cu_row,
                               self.urc_p.cu_col + self.urc_p.max_col - 1,
                               '四 待解决Bug列表-按处理状态', self.fd['seg_title_format'])
        self.urc_p.next_row()
        self.urc_ws.write_row(self.urc_p.cu_row, self.urc_p.cu_col,
                             unresolved_thead, self.fd['lh_format'])
        self.urc_p.next_row()
        kinds = data.keys()
        kinds = sorted(kinds)
        if '其他' in kinds:
            kinds.remove('其他')
        kinds.append('其他')
        unresolved_thead.pop(0)
        for kind in kinds:
            if len(data[kind]) != 0:
                self.urc_ws.write(self.urc_p.cu_row, self.urc_p.cu_col,
                                 kind, self.fd['list_format'])
                self.urc_p.next_col()
                for item in data[kind]:
                    for head in unresolved_thead:
                        self.urc_ws.write(self.urc_p.cu_row, self.urc_p.cu_col,
                                         item[head], self.fd['list_format'])
                        self.urc_p.next_col()
                        if head == '是否重新打开':
                            self.urc_p.next_row()
                            self.urc_p.set_col(1)
                self.urc_p.set_col(0)

    def make_new_bug_sheet(self):
        data = self.ct.make_write_dict('new')
        new_thead = ['负责小组', 'PSR编号', '标题', '状态', '严重性', '创建时间']
        self.nb_ws.set_column('A:A', 24.38)
        self.nb_ws.set_column('B:B', 13.25)
        self.nb_ws.set_column('C:C', 60)
        self.nb_ws.set_column('F:F', 10.38)
        title = '四 新增bug列表'
        if self.ct.config.classify_switch == True:
            title = '五 新增bug列表'
        self.nb_ws.merge_range(self.nb_p.cu_row, self.nb_p.cu_col,
                               self.nb_p.cu_row,
                               self.nb_p.cu_col + self.nb_p.max_col - 1,
                               title, self.fd['seg_title_format'])
        self.nb_p.next_row()
        self.nb_ws.write_row(self.nb_p.cu_row, self.nb_p.cu_col,
                             new_thead, self.fd['lh_format'])
        self.nb_p.next_row()
        groups = data.keys()
        new_thead.pop(0)
        for group in groups:
            if len(data[group]) != 0:
                self.nb_ws.write(self.nb_p.cu_row, self.nb_p.cu_col,
                                 group, self.fd['list_format'])
                self.nb_p.next_col()
                for item in data[group]:
                    for head in new_thead:
                        self.nb_ws.write(self.nb_p.cu_row, self.nb_p.cu_col,
                                         item[head], self.fd['list_format'])
                        self.nb_p.next_col()
                        if head == '创建时间':
                            self.nb_p.next_row()
                            self.nb_p.set_col(1)
                self.nb_p.set_col(0)

    def make_resolved_bug_sheet(self):
        data = self.ct.make_write_dict('resolved')
        resolved_thead = ['负责小组', 'PSR编号', '标题', '状态',
                          '严重性', '解决时间', '缺陷定位', '解决方案']
        self.rs_ws.set_column('A:A', 24.38)
        self.rs_ws.set_column('B:B', 13.25)
        self.rs_ws.set_column('C:C', 60)
        self.rs_ws.set_column('F:F', 10.38)
        title = '五 解决bug列表'
        if self.ct.config.classify_switch == True:
            title = '六 解决bug列表'
        self.rs_ws.merge_range(self.rs_p.cu_row, self.rs_p.cu_col,
                               self.rs_p.cu_row,
                               self.rs_p.cu_col + self.rs_p.max_col - 1,
                               title, self.fd['seg_title_format'])
        self.rs_p.next_row()
        self.rs_ws.write_row(self.rs_p.cu_row, self.rs_p.cu_col,
                             resolved_thead, self.fd['lh_format'])
        self.rs_p.next_row()
        groups = data.keys()
        resolved_thead.pop(0)
        for group in groups:
            if len(data[group]) != 0:
                self.rs_ws.write(self.rs_p.cu_row, self.rs_p.cu_col,
                                 group, self.fd['list_format'])
                self.rs_p.next_col()
                for item in data[group]:
                    for head in resolved_thead:
                        #print head
                        self.rs_ws.write(self.rs_p.cu_row, self.rs_p.cu_col,
                                         item[head], self.fd['list_format'])
                        self.rs_p.next_col()
                        if head == '解决方案':
                            self.rs_p.next_row()
                            self.rs_p.set_col(1)
                self.rs_p.set_col(0)

    def make_validate_bug_sheet(self):
        data = self.ct.make_write_dict('validate')
        validate_thead = ['负责小组', 'PSR编号', '标题', '创建时间',
                          '解决时间', '创建者', '出现频率', '缺陷定位',
                          '解决方案']
        self.vd_ws.set_column('A:A', 24.38)
        self.vd_ws.set_column('B:B', 13.25)
        self.vd_ws.set_column('C:C', 60)
        self.vd_ws.set_column('F:F', 10.38)
        title = '六 待验证Bug列表'
        if self.ct.config.classify_switch == True:
            title = '七 待验证Bug列表'

        self.vd_ws.merge_range(self.vd_p.cu_row, self.vd_p.cu_col,
                               self.vd_p.cu_row,
                               self.vd_p.cu_col + self.vd_p.max_col - 1,
                               title, self.fd['seg_title_format'])
        self.vd_p.next_row()
        self.vd_ws.write_row(self.vd_p.cu_row, self.vd_p.cu_col,
                             validate_thead, self.fd['lh_format'])
        self.vd_p.next_row()
        groups = data.keys()
        validate_thead.pop(0)
        for group in groups:
            if len(data[group]) != 0:
                self.vd_ws.write(self.vd_p.cu_row, self.vd_p.cu_col,
                                 group, self.fd['list_format'])
                self.vd_p.next_col()
                for item in data[group]:
                    for head in validate_thead:
                        self.vd_ws.write(self.vd_p.cu_row, self.vd_p.cu_col,
                                         item[head], self.fd['list_format'])
                        self.vd_p.next_col()
                        if head == '解决方案':
                            self.vd_p.next_row()
                            self.vd_p.set_col(1)
                self.vd_p.set_col(0)

    def make_delay_bug_sheet(self):
        data = self.ct.make_write_dict('delay')
        delay_thead = ['负责小组', 'PSR编号', '标题', '状态', '严重性',
                            '创建时间', '出现频率', '当前审阅者', 'Reject次数', '是否重新打开']
        self.dl_ws.set_column('A:A', 24.38)
        self.dl_ws.set_column('B:B', 13.25)
        self.dl_ws.set_column('C:C', 60)
        self.dl_ws.set_column('F:F', 10.38)
        title = '七 延期bug列表'
        if self.ct.config.classify_switch == True:
            title = '八 延期bug列表'
        self.dl_ws.merge_range(self.dl_p.cu_row, self.dl_p.cu_col,
                               self.dl_p.cu_row,
                               self.dl_p.cu_col + self.dl_p.max_col - 1,
                               title, self.fd['seg_title_format'])
        self.dl_p.next_row()
        self.dl_ws.write_row(self.dl_p.cu_row, self.dl_p.cu_col,
                             delay_thead, self.fd['lh_format'])
        self.dl_p.next_row()
        groups = data.keys()
        groups = sorted(groups)
        delay_thead.pop(0)
        for group in groups:
            if len(data[group]) != 0:
                self.dl_ws.write(self.dl_p.cu_row, self.dl_p.cu_col,
                                 group, self.fd['list_format'])
                self.dl_p.next_col()
                for item in data[group]:
                    for head in delay_thead:
                        self.dl_ws.write(self.dl_p.cu_row, self.dl_p.cu_col,
                                         item[head], self.fd['list_format'])
                        self.dl_p.next_col()
                        if head == '是否重新打开':
                            self.dl_p.next_row()
                            self.dl_p.set_col(1)
                self.dl_p.set_col(0)