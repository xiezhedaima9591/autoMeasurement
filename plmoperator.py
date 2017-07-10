#-*-coding:utf-8-*-

import os
import re
import time
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from config import Config


class Operator(object):
    '此类用于操作PLM系统导出数据'

    def __init__(self, config):
        fp = webdriver.FirefoxProfile()
        fp.set_preference("browser.download.folderList", 2)
        fp.set_preference("browser.download.manager.showWhenStarting", False)
        fp.set_preference("browser.download.dir", os.getcwd())
        fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel")
        self.config = config
        try:
            self.driver = webdriver.Firefox(firefox_profile = fp)
        except exceptions.WebDriverException, e:
            print '浏览器不存在'
            print e
            return
        

    def login(self):
        '登录操作'
        self.driver.get('http://plm.kedacom.com:7001/Agile/')
        try:
            username_text = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'ui_username')))
            password_text = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'ui_password')))
            login_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.NAME, 'submit')))
        except exceptions.TimeoutException, e:
            print '相关元素未找到'
            print e
            return
        username_text.clear()
        password_text.clear()
        username_text.send_keys(self.config.username)
        password_text.send_keys(self.config.password)
        login_button.click()

    def find_sys_window(self):
        '将活动窗口切换到系统主窗口'
        windows = self.driver.window_handles
        len_of_windows = len(windows)
        index_window = 0
        target_title = r'Agile Product Lifecycle Management.+'

        while index_window < len_of_windows:
            self.driver.switch_to_window(windows[index_window])
            title_now = self.driver.title
            title_match = re.match(target_title, title_now)
            if title_match:
                break
            else:
                index_window += 1

    def operate_main(self):
        '操作系统主窗口'
        self._show_search_filed_panel() #显示出搜索字段面板
        self._select_search_filed()  #选择搜索字段
        self._input_subject_name()  #输入项目名称操作
        self._search()  #搜索
        self._select_data_filed()  #选择所需数据字段
        wait_flag_path = "//table[@id='QUICKSEARCH_TABLE']//span[@title='%s']" % ('缺陷描述.负责小组')
        try:
            wait_flag = WebDriverWait(self.driver, 300).until(EC.presence_of_element_located((By.XPATH, wait_flag_path)))
            #以表头‘负责小组’是否加载完成为标准来判断导出数据是否加载完成
        except exceptions.TimeoutException, e:
            print '定位错误'
            print e
            return
        self._export_data()  #导出数据
        print u'等待文件下载中...'
        while not os.path.exists(u'搜索结果.xls') or os.stat(u'搜索结果.xls').st_size == 0 or os.path.exists(u'搜索结果.xls.part'):
            os.system('cls')
            for i in range(5):
                time.sleep(1)
                print '.',
            #pass #等待导出文件下载完成
        print
        self.driver.quit()
    
    def _show_search_filed_panel(self):
        '操作显示出搜索字段面板'
        try:
            param_search = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'top_paramSearch')))
        except exceptions.TimeoutException, e:
            print '相关元素未找到'
            print e
            return
        time.sleep(2)
        param_search.click()
        time.sleep(2)
        adv_path = "//div[@id='basicSearchButtonBar']//p/a[contains(text,%s)]" % ("高级".decode('GBK'))
        try:
            adv_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, adv_path)))
        except exceptions.TimeoutException, e:
            print '相关元素未找到'
            print e
            return
        adv_button.click()

    def _select_search_filed(self):
        '选择搜索字段'
        bo_str = "//div[@id='main_search_content']/div[@id='advanced_search']//div[@id='advancedClassOptions']"
        try:
            basic_options = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, bo_str)))
        except exceptions.TimeoutException, e:
            print '相关元素未找到'
            print e
            return
        try:
            selects = WebDriverWait(basic_options, 10).until(EC.presence_of_all_elements_located((By.TAG_NAME, 'select')))
        except exceptions.TimeoutException, e:
            print '相关元素未找到'
            print e
        one_select = selects[0]
        one_select = Select(one_select)
        one_select.select_by_value('4750')
        two_select = selects[1]
        two_select = Select(two_select)
        two_select.select_by_value('2474118')
        three_select_path = "//div[@id='advancedOptionContent']//div[@class='filter_attributes nodrag']//select[contains(@class,'nodrag')]"
        try:
            three_select = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, three_select_path)))
        except exceptions.TimeoutException, e:
            print '相关元素未找到'
            print e
            return
        three_select = Select(three_select)
        if self.config.product_name == 'null':
            three_select.select_by_value('1550')
        elif self.config.subject_name == 'null':
            three_select.select_by_value('1549')
    
    def _input_subject_name(self):
        '输入项目名称'
        click_input_path = "//div[@id='advancedOptionContent']//div[@class='filter_value nodrag']//input[contains(@id, 'input-edit_mode_agile')]"
        try:
            click_input = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, click_input_path)))
        except exceptions.TimeoutException, e:
            print '相关元素未找到'
            print e
            return
        click_input.click()
        show_floater_path = "//div[@id='select-control']//a[contains(title,%s)]" % ('搜索以添加')
        try:
            show_floater_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, show_floater_path)))
        except exceptions.TimeoutException, e:
            print '相关元素未找到'
            print e
            return
        show_floater_button.click()
        text_input_path = "//div[@id='floater_window']//input[starts-with(@id, 'floater_search_text_agile') and @name='search']"
        try:
            text_input = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, text_input_path)))
        except exceptions.TimeoutException, e:
            print '相关元素未找到'
            print e
            return
        if self.config.subject_name != 'null':
            text_input.send_keys(self.config.subject_name.decode('GBK'))
        elif self.config.product_name != 'null':
            text_input.send_keys(self.config.product_name.decode('GBK'))
        quick_search_path = "//div[@id='floater_window']//span[contains(@id,'search-yui')]"
        try:
            quick_search = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, quick_search_path)))
        except exceptions.TimeoutException, e:
            print '相关元素未找到'
            print e
            return
        quick_search.click()
        target_item_path = "//div[@id='floater_window']//table[@class='yui-dt-table']/tbody[2]/tr[1]/td[1]"
        try:
            target_item = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, target_item_path)))
        except exceptions.TimeoutException, e:
            print '相关元素未找到'
            print e
            return
        ActionChains(self.driver).double_click(target_item).perform()

    def _search(self):
        '执行搜索'
        search_button_path = "//div[@id='advancedSearchButtonBar']//p//a[contains(@id,'advSearch-as-yui')]"
        try:
            search_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, search_button_path)))
        except exceptions.TimeoutException, e:
            print '相关元素未找到'
            print e
            return
        search_button.click()

    def _select_data_filed(self):
        '选择所需数据字段'
        format_button_path = "//div[@id='advSearchColumnOne']//a[contains(@id, 'advFormat-as-yui')]"
        try:
            format_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, format_button_path)))
        except exceptions.TimeoutException, e:
            print '未找到相关字段'
            print e
            return
        format_button.click()
        try:
            format_tab = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'formatColumnsTab')))
        except exceptions.TimeoutException, e:
            print '未找到相关字段1'
            print e
            return
        format_tab.click()
        delete_col1_text = '工作流 (封面)'
        try:
            delete_col1 = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, delete_col1_text)))
        except exceptions.TimeoutException, e:
            print '未找到相关字段2'
            print e
            return
        ActionChains(self.driver).double_click(delete_col1).perform()
        delete_col2_text = 'PSR 类型 (封面)'
        try:
            delete_col2 = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, delete_col2_text)))
        except exceptions.TimeoutException, e:
            print '未找到相关字段3'
            print e
            return
        ActionChains(self.driver).double_click(delete_col2).perform()
        
        toggles_path = "//div[@id='formatContentEl']//div[@class=' lt_column nodrag']/div[starts-with(@id, 'sc_control')]//span[starts-with(@id, 'agileTreeNodeToggle')]"
        try:
            toggles = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, toggles_path)))
        except exceptions.TimeoutException, e:
            print '未找到相关字段4'
            print e
            return
        for toggle in toggles:
            toggle.click()

        need_list = ['产品线/产品族', '负责小组', '产品名称', '系统版本号',
                     '严重性', '创建日期', '解决时间', 'Reject次数', '出现频率',
                     '当前审阅者', '缺陷定位说明   [Resolve必填]',
                     '解决方案   [Resolve必填]', '创建者']
        for need in need_list:
            try:
                one_item = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, need)))
            except exceptions.TimeoutException, e:
                print '未找到相关字段'
                print e
                return
            ActionChains(self.driver).double_click(one_item).perform()

        applicate_text = '应用'
        try:
            applicate_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, applicate_text)))
        except exceptions.TimeoutException, e:
            print '未找到相关字段'
            print e
            return
        applicate_button.click()
        close_span_path = "//div[@id='palette-handlebar']/following-sibling::*[1]"
        try:
            close_span = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, close_span_path)))
        except exceptions.TimeoutException, e:
            print '未找到相关字段'
            print e
            return
        close_span.click()

    def _export_data(self):
        filename = u'搜索结果.xls'
        if os.path.exists(filename):
            os.remove(filename)
        more_button_path = "//div[starts-with(@id, 'table_actions')]/div[@class='column_one no_width']/p/a[2]"
        try:
            more_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, more_button_path)))
        except exceptions.TimeoutException, e:
            print '未找到相关元素'
            print e
            return
        more_button.click()
        export_text = '导出 (xls)'
        try:
            export_button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, export_text)))
        except exceptions.TimeoutException, e:
            print '未找到相关元素'
            print e
            return
        export_button.click()

if __name__ == '__main__':
    c = Config()
    c.read_configfile()
    o = Operator(c)
    o.login()
    o.find_sys_window()
    try:
        o.operate_main()
    except exceptions.UnexpectedAlertPresentException, e:
        print 'PLM系统出错，请重新运行程序'
        print e
    


