# -*- coding:utf-8 -*-
import os
import sys
from plmoperator import Operator
from counter import Counter
from maker import Maker
from config import Config

class Measure(object):
    def __init__(self):
        self.config = Config()
        self.config.read_configfile()
        if (self.config.subject_name == 'null'
            and self.config.product_name == 'null') \
            or (self.config.subject_name != 'null'
                and self.config.product_name != 'null'):#产品名称和所属项目两者必填且只能填一个，别一个用填字符串null
            raise Exception('wrong config!')  #所以两者都为null或都不为null要报配置错误
        self.operator = Operator(self.config)
        self.counter = Counter(self.config)
        self.maker = None

    def run(self):
        self.operator.login()
        self.operator.find_sys_window()
        self.operator.operate_main()
        self.counter.read_original_data()
        self.counter.count_all_bug()
        self.counter.count_unresolved_bug()
        self.maker = Maker(self.counter)
        self.maker.make_bug_count_sheet()
        self.maker.make_unresolved_bug_sheet()
        if self.config.classify_switch == True:
            self.maker.make_cunresolved_bug_sheet()
        self.maker.make_new_bug_sheet()
        self.maker.make_resolved_bug_sheet()
        self.maker.make_validate_bug_sheet()
        if self.config.delay_switch == True:
            self.maker.make_delay_bug_sheet()
        self.maker.workbook.close()


if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf-8')
    os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
    m = Measure()
    m.run()
