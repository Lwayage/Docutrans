##############################################################
# 代码复制自https://www.cnblogs.com/sunnyroot/p/10797358.html #
# 以下代码在原代码的基础上在comp2filename函数中添加了一行判断    #
##############################################################

# -*- coding: utf-8 -*-
# @Time    : 2019/4/30 13:32
# @Author  : shine
# @File    : mix_sort.py
"""
基于字符串数字混合排序的Python脚本
"""

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False

def find_continuous_num(astr, c):
    num = ''
    try:
        while not is_number(astr[c]) and c < len(astr):
            c += 1
        while is_number(astr[c]) and c < len(astr):
            num += astr[c]
            c += 1
    except:
        pass
    if num != '':
        return int(num)

def comp2filename(file1, file2):
    smaller_length = min(len(file1), len(file2))
    for c in range(0, smaller_length):
        if not is_number(file1[c]) and not is_number(file2[c]):
            if file1[c] < file2[c]:
                return True
            if file1[c] > file2[c]:
                return False
            if file1[c] == file2[c]:
                if c == smaller_length - 1:
                    if len(file1) < len(file2):
                        return True
                    else:
                        return False
                else:
                    continue
        if is_number(file1[c]) and not is_number(file2[c]):
            return True
        if not is_number(file1[c]) and is_number(file2[c]):
            return False
        if is_number(file1[c]) and is_number(file2[c]):
            if find_continuous_num(file1, c) < find_continuous_num(file2, c):
                return True
            if find_continuous_num(file1, c) == find_continuous_num(file2, c):   # 原代码没有这一判断
                continue                                                         # 加这一判断后可识别同一文件名里的多组数字
            else:
                return False

def sort_list_by_name(lst):
    for i in range(1, len(lst)):
        x = lst[i]
        j = i
        while j > 0 and comp2filename(x, lst[j-1]):
            lst[j] = lst[j-1]
            j -= 1
        lst[j] = x
    return lst