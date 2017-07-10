#-*- coding:utf-8 -*-

class Pointer(object):
    '写入位置指示器'

    def __init__(self, cu_row = 0, cu_col = 0, max_col = 15):
        self.cu_row = cu_row
        self.cu_col = cu_col
        self.max_col = max_col

    def next_row(self):
        self.cu_row += 1
        self.cu_col = 0

    def next_col(self):
        if self.cu_col == self.max_col:
            return
        self.cu_col += 1

    def set_row(self, row):
        self.cu_row = row
        self.cu_col = 0
    
    def set_col(self, col):
        self.cu_col = col

