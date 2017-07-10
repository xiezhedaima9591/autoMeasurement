#-*- coding:utf-8 -*-

import re
import os
import copy
import xlrd
from datetime import datetime

class Counter(object):
    '统计器，用来读取原始数据和对原始数据进行统计'

    def __init__(self, config, file_name = u'搜索结果.xls'):
        self.file_name = file_name#.decode('utf-8')
        self.config = config
        self.title = (
                      'PSR编号','标题','状态','产品族','负责小组','产品名称',
                      '系统版本号','严重性', '创建时间', '解决时间','Reject次数',
                      '出现频率', '当前审阅者', '缺陷定位', '解决方案', '创建者',
                      )
        self.lint = [] #用于存放原始数据
        self.bug_count_dict = {} #用于存放bug统计总况
        self.unresolved_result = {} #用于存放待解决bug统计情况
        self.new_dict = {}
        self.unresolved_dict = {}
        self.resolved_dict = {}
        self.validate_dict = {}

    def read_original_data(self):
        '读取导出数据'
        try:
            data_wb = xlrd.open_workbook(self.file_name)
        except IOError, e:
            print '导出数据文件不存在或损坏'
            print e
            return
        data_ws = data_wb.sheet_by_index(0)
        row_num = data_ws.nrows
        result = []
        for index in range(row_num):
            if index == 0:
                continue
            else:
                result.append(data_ws.row_values(index))
        self._make_result_list(result)  #将读取的数据存入到lint中

    def _make_result_list(self, result):
        '将读取的数据结构化并存储'
        parr = r'.+\|.+\|'
        time_parr = r'\d{2}/\d{2}/\d{4}'
        parr = re.compile(parr)
        time_parr = re.compile(time_parr)
        key = 0
        for i in result:
            n = 0
            lintxx = dict.fromkeys(self.title, 'None')
            for j in i:
                if self.title[n] == '负责小组':
                    j = re.sub(parr, '', j)
                if '时间' in self.title[n] and j != '':
                    j = re.match(time_parr, j).group()
                lintxx[self.title[n]] = str(self._CodeToUtf_8(j))
                n += 1
            self.lint.append(lintxx)
    
    def _CodeToUtf_8(self, data):
        '转码'
        if type(data) == 'type\'str\'':
            return data.decode('utf-8')
        else:
            return data

    def count_all_bug(self):
        '统计总体bug情况'
        if self.config.product_name == 'null':
            self.bug_count_dict['subject_name'] = self.config.subject_name.decode('GBK')
        elif self.config.subject_name == 'null':
            self.bug_count_dict['subject_name'] = self.config.product_name.decode('GBK')
        sub_parr = r'\|'
        sub_parr = re.compile(sub_parr)
        self.bug_count_dict['subject_name'] = re.sub(sub_parr,
                                            ' ',self.bug_count_dict['subject_name'])
        #有些项目名称里有|符号，因为结果文件要以项目名称命名,而|符号不能做为文件名，所以要提前去掉|符号
        self.bug_count_dict['unresolved_num'] = 0
        self.bug_count_dict['validate_num'] = 0
        self.bug_count_dict['delay_num'] = 0
        for one_item in self.lint:
            if one_item['状态'] not in ('Validate', '关闭'):
                self.bug_count_dict['unresolved_num'] += 1
                if self.config.delay_keyword in one_item['标题']:
                    self.bug_count_dict['delay_num'] += 1
            elif one_item['状态'] == 'Validate':
                self.bug_count_dict['validate_num'] += 1

        self.bug_count_dict['unresolved_num'] = \
            self.bug_count_dict['unresolved_num'] - \
            self.bug_count_dict['delay_num']

        self.bug_count_dict['sum_num'] = \
            self.bug_count_dict['unresolved_num'] + \
            self.bug_count_dict['validate_num'] + \
            self.bug_count_dict['delay_num']

    def count_unresolved_bug(self):
        '统计未解决bug情况'
        groups = []
        self.bug_count_dict['new_bug'] = 0
        self.bug_count_dict['resolved_bug'] = 0
        given_date = self.get_given_date()

        for one_item in self.lint:
            #把有数据里的小组列出来
            group = one_item['负责小组']
            self.unresolved_result[group] = []

        groups = self.unresolved_result.keys()
        importance = ('1-致命', '2-严重', '3-普通', '4-较低', '5-建议',)
        total = [0, 0, 0, 0, 0, 0, 0, 0]
        for group in groups:
            pond = [0, 0, 0, 0, 0, 0, 0, 0]
            for one_item in self.lint:
                create_date = datetime.strptime(one_item['创建时间'], '%m/%d/%Y').date()
                if one_item['解决时间'] != '':
                    resolved_date = datetime.strptime(one_item['解决时间'], '%m/%d/%Y').date()
                if group == one_item['负责小组']:
                    if one_item['状态'] not in ('Validate', '关闭') and \
                            self.config.delay_keyword not in one_item['标题']:
                        if one_item['严重性'] in importance:
                            pond[importance.index(one_item['严重性'])] += 1
                            total[importance.index(one_item['严重性'])] += 1

                        pond[5] += 1
                        total[5] += 1

                        if create_date >= given_date:
                            pond[6] += 1
                            total[6] += 1
                            self.bug_count_dict['new_bug'] += 1
                    else:
                        if resolved_date >= given_date:
                            pond[7] += 1
                            total[7] += 1
                            self.bug_count_dict['resolved_bug'] += 1
                else:
                    continue

            self.unresolved_result[group] = pond

        self.unresolved_result['总计'] = total
        self._record_report()

    def get_given_date(self):
        '将配置里的所给日期从字符串型转换成日期型'
        try:
            assert isinstance(self.config.given_date, str)
            try:
                given_date = datetime.strptime(self.config.given_date, '%Y/%m/%d').date()
            except ValueError, e:
                print '给出的日期不是正确格式'
                print e
                return
        except AssertionError, args:
            print '给出的日期不是正确的类型'
            #print e
            return
        return given_date

    def _record_report(self):
        '将本次统计结果记录在文件里，以便画趋势图时使用'
        today = datetime.today()
        today = today.strftime('%Y-%m-%d')
        today_report = [today, str(self.bug_count_dict['unresolved_num']),
                        str(self.bug_count_dict['validate_num']),
                        str(self.bug_count_dict['resolved_bug']),
                        str(self.bug_count_dict['new_bug']),
                        self.config.given_date]
        record_file_name = self.bug_count_dict['subject_name'] + '.txt'
        if os.path.exists(record_file_name):
            record_file = open(record_file_name, 'r+')
        else:
            record_file = open(record_file_name, 'w+')
        record_str = record_file.read()
        record_matrix, item_length = self.str_to_matrix(record_str)
        if item_length == 1:
            #新建文件长度会有一个空字符串被当成一条记录
            record_matrix.pop(0) #去掉空字符串
            record_matrix = self._create_record(record_matrix, today_report) #新建文件时创建记录
        elif item_length > 1 and item_length < 5:
            record_matrix = self._append_record(record_matrix, today_report) #非新建文件时，将当天统计记录添加进去
        else:
            try:
                assert item_length > 1, '记录长度小于1，数据出错'
                record_matrix = self._replace_record(record_matrix, today_report, 1)
            except AssertionError, e:
                print e
                return
        title_item = record_matrix.pop(0)
        record_matrix = sorted(record_matrix, key = lambda x: x[0])
        record_matrix.insert(0, title_item)
        record_str = self.matrix_to_str(record_matrix)
        record_file = open(record_file_name, 'w+')
        record_file.write(record_str)
        record_file.close()

    def _record_validate(self):
        record_file_name = self.bug_count_dict['subject_name'] + '_validate.txt'
        if os.path.exists(record_file_name):
            os.remove(record_file_name)
        psr_list = []
        for pond in self.lint:
            if pond['状态'] == 'Validate':
                psr_list.append(pond['PSR编号'])
        record_file = open(record_file_name, 'w+')
        for item in psr_list:
            record_file.write(item + '\n')
        record_file.close()

    def _read_validate(self):
        record_file_name = self.bug_count_dict['subject_name'] + '_validate.txt'
        try:
            record_file = open(record_file_name, 'r')
        except IOError:
            return []
        record_str = record_file.read()
        record_file.close()
        item_flag = r'\n'
        psr_list =[item for item in re.split(item_flag,record_str)]
        return psr_list



    def _create_record(self, record_matrix, today_report):
        title_item = ['记录日期', '待解决', '待验证', '解决', '新增', '节点日期']
        record_matrix.append(title_item)
        record_matrix.append(today_report)
        return record_matrix

    def _append_record(self, record_matrix, today_report):
        today = today_report[0]
        for index, one_item in enumerate(record_matrix):
            if index == 0:
                continue
            else:
                record_date = one_item[0]
                if record_date == today: #同一天对同一个项目进行统计
                    if self.config.given_date in one_item: #如果给出日期没有变
                        return record_matrix  #则不用再进行记录，将原来的内容返回
                    else: #如果给出日期变化了
                        record_matrix = self._replace_record(record_matrix, today_report, len(record_matrix) - 1)
                        return record_matrix  #把原来的那条记录替换成新的给出日期统计的数据
        record_matrix.append(today_report)
        return record_matrix

    def _replace_record(self, record_matrix, today_report, flag):
        today = today_report[0]
        for index, one_item in enumerate(record_matrix):
            if index == 0:
                continue
            else:
                record_date = one_item[0]
                if record_date == today and (self.config.given_date in one_item):
                    return record_matrix
        record_matrix.pop(flag) #记录文件的条目数超过五个时把第一条删除后追加新数据，同一统计日期不同节点日期时，把最后一条数据替换
        record_matrix.append(today_report)
        return record_matrix

    def str_to_matrix(self, string):
        '字符串转二维列表'
        item_flag = r'\n\s*\n'
        field_flag = r'\s*'
        temp_item = []
        record_item = [one_item for one_item in re.split(item_flag, string)]
        item_length = len(record_item)
        for one_item in record_item:
            temp_item.append([one_field.strip() for one_field in re.split(field_flag, one_item)])
        return temp_item, item_length
    
    def matrix_to_str(self, matrix):
        '二维列表转字符串'
        temp_item = []
        for one_item in matrix:
            temp_item.append('    '.join(one_item))
        string = '\n\n'.join(temp_item)
        return string

    def make_write_dict(self, keyword):
        '生成写入各个sheet的数据'
        given_date = self.get_given_date()
        group_dict = {}
        result_dict = {}
        for one_item in self.lint:
            group_dict[one_item['负责小组']] = 0

        groups = group_dict.keys()

        for group in groups:
            result_dict[group] = []

        for group in groups:
            temp_item  = []
            for one_item in self.lint:
                if one_item['负责小组'] == group:
                    if keyword == 'new':
                        create_date = datetime.strptime(one_item['创建时间'], '%m/%d/%Y').date()
                        if create_date >= given_date:
                            temp_item.append(one_item)
                    elif keyword == 'unresolved':
                        if one_item['状态'] not in ('Validate', '关闭') and \
                            self.config.delay_keyword not in one_item['标题']:
                            one_item['是否重新打开'] = ' '
                            psr_list = self._read_validate()
                            if one_item['PSR编号'] in psr_list:
                                one_item['是否重新打开'] = '是'
                            temp_item.append(one_item)
                    elif keyword == 'resolved':
                        if one_item['状态'] in ('Validate', '关闭'):
                            temp_item.append(one_item)
                    elif keyword == 'validate':
                        if one_item['状态'] == 'Validate':
                            temp_item.append(one_item)
                    elif keyword == 'delay':
                        if one_item['状态'] not in ('Validate', '关闭') and \
                            self.config.delay_keyword in one_item['标题']:
                            one_item['是否重新打开'] = ' '
                            psr_list = self._read_validate()
                            if one_item['PSR编号'] in psr_list:
                                one_item['是否重新打开'] = '是'
                            temp_item.append(one_item)
            result_dict[group] = temp_item

        return result_dict

    def make_classify_dict(self):
        given_date = self.get_given_date()
        result_dict = {}
        kinds = self.config.kind_keywords
        classify_keywords = self.config.classify_keywords
        for index, kind in enumerate(kinds):
            kinds[index]  = self._CodeToUtf_8(kind)
        for index, classify in enumerate(classify_keywords):
            classify_keywords[index] = self._CodeToUtf_8(classify)
        for kind in kinds:
            result_dict[kind] = []

        other_item = []
        temp_lint = copy.deepcopy(self.lint)

        for classify in zip(classify_keywords,kinds):
            temp_item = []
            for one_item in temp_lint:
                if one_item['状态'] not in ('Validate', '关闭') and \
                                self.config.delay_keyword not in one_item['标题']:
                    if classify[0] in one_item['标题']:
                        one_item['是否重新打开'] = ' '
                        psr_list = self._read_validate()
                        if one_item['PSR编号'] in psr_list:
                            one_item['是否重新打开'] = '是'
                        temp_item.append(one_item)
                        temp_lint.remove(one_item)

            result_dict[classify[1]] = temp_item

        for one_item in temp_lint:
            if one_item['状态'] not in ('Validate', '关闭') and \
                            self.config.delay_keyword not in one_item['标题']:
                one_item['是否重新打开'] = ' '
                psr_list = self._read_validate()
                if one_item['PSR编号'] in psr_list:
                    one_item['是否重新打开'] = '是'
                other_item.append(one_item)

        result_dict['其他'] = other_item

        print len(result_dict)


        self._record_validate()
        return result_dict


if __name__ == '__main__':
    import os
    import sys
    from config import Config
    reload(sys)
    sys.setdefaultencoding('utf-8')
    os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
    cf = Config()
    cf.read_configfile()
    ct = Counter(cf)
    ct.read_original_data()
    ct.count_all_bug()
    ct.count_unresolved_bug()
    ct.make_write_dict('new')

