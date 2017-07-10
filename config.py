# -*-coding:utf-8-*-
import re


class Config(object):
    '配置器'

    def __init__(self, file_name='./config.ini'):
        self.file_name = file_name
        self.username = ''
        self.password = ''
        self.subject_name = ''
        self.product_name = ''
        self.given_date = ''
        self.circle = ''
        self.delay_keyword = ''
        self.filter_term = ''
        self.classify_keywords = []
        self.kind_keywords = []
        self.classify_switch = True
        self.delay_switch = True

    def read_configfile(self):
        item_flag = r'\n\s*\n'
        nav_flag = '='
        try:
            config_file = open(self.file_name, 'r')
        except IOError, e:
            print '找不到该文件或打开失败'
            print e
            return

        config_str = config_file.read()
        config_file.close()
        config_item = [one.strip() for one in re.split(item_flag, config_str)]
        for one_item in config_item:
            attr, value = [one.strip() for one in one_item.split(nav_flag)]
            if attr not in ('classify_keywords', 'kind_keywords', 'classify_switch', 'delay_switch'):
                statement = "self.%s = '%s'" % (attr, value)
            else:
                statement = "self.%s = %s" % (attr, value)
            exec (statement)


            # print self.username
            # print self.password
            # print self.product_name
            # print self.given_date
            # print self.circle


if __name__ == '__main__':
    c = Config('./config.ini')
    c.read_configfile()
    print c.subject_name
    print c.product_name