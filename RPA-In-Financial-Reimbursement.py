# -*- encoding:utf-8 -*-

"""
@作者：Charlie
@备注：隐私相关已做处理
"""

import time
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.keys import Keys
import csv
import openpyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

root = tk.Tk()
root.title("请操作")
root.geometry("400x100+500+300")
sum_name = []


# txt转换处理
def txt_to_excel(txt_address):
    input_file = txt_address
    output_file = txt_address[:-4] + '.xlsx'

    wb = openpyxl.Workbook()
    ws = wb.worksheets[0]

    with open(input_file, 'rt', encoding="ANSI") as data:
        reader = csv.reader(data, delimiter='|')
        ws.append(['账号', '名称', '', '', '状态', '注释', '', '提示'])
        for row in reader:
            if len(row) == 6:
                row.pop(1)
                row.pop(2)
                row.pop(3)
                row[2] = row[2].rstrip()
                row.insert(2, '')
                row.insert(3, '')
                row.insert(4, '成功')
                row.insert(6, '')
                row.insert(7, '')
                ws.append(row)
    wb.save(output_file)
    return output_file


# excel处理
def excel_matching(excel_addr):
    name_dict = {}
    df_map = pd.read_excel(excel_addr, dtype=str)
    for row in df_map.values:
        name = row[1]
        if name not in name_dict:
            bool_index1 = df_map['名称'].str.contains(name)
            filter_data = list(df_map['状态'][bool_index1])
            filter_data2 = list(df_map['注释'][bool_index1])

            # 过滤状态为失败的元素
            for index, item in enumerate(filter_data):
                if item != '成功':
                    filter_data.pop(index)
                    filter_data2.pop(index)

            name_dict[name] = filter_data2
    sum_name.append(name_dict)


# 生成匹配字典

def build_nameBigDict():
    # 定位工作文件夹
    global filenames
    tkinter.messagebox.showinfo('提示', '请选择要处理的文件夹（里面有要处理的excel或txt文件）')
    Folderpath = filedialog.askdirectory()
    inputdir = Folderpath
    for parents, dirnames, filenames in os.walk(inputdir):
        for filename in filenames:
            txt_file = filename[-4:]
            if txt_file == '.txt':
                excel_addr = txt_to_excel(inputdir + '/' + filename)
                excel_matching(excel_addr)
            else:
                excel_matching(inputdir + '/' + filename)
    return filenames


# 数据初始化
filename_s = build_nameBigDict()

# 浏览器初始化

# 创建Chrome浏览器对象，这会在电脑上在打开一个浏览器窗口
# 放入你谷歌浏览器驱动的位置
browser = webdriver.Chrome(executable_path="")

# 隐性等待，最长等30秒
# browser.implicitly_wait(5)

# 通过浏览器向服务器发送URL请求
browser.get("")

sleep(2)

# 设置浏览器的大小
browser.maximize_window()

# 输入用户名和密码
browser.find_element_by_xpath('//*[@id="login_username"]').send_keys('')
browser.find_element_by_xpath('//*[@id="login_password"]').send_keys('')

# 点击登录
browser.find_element_by_xpath('//*[@id="login_button"]').send_keys(Keys.ENTER)

# 定位到新页面
for handle in browser.window_handles:
    browser.switch_to.window(handle)

# 定位到iframe
iframe = browser.find_element_by_xpath(
    '//*[@id="main"]')
# 切换到iframe
browser.switch_to.frame(iframe)

# 点击“更多”
browser.find_element_by_class_name("sectionMoreIco").click()

# 退出iframe
browser.switch_to.default_content()

# 定位到新页面
for handle in browser.window_handles:
    browser.switch_to.window(handle)

# 定位到iframe
iframe = browser.find_element_by_xpath(
    '//*[@id="main"]')
# 切换到iframe
browser.switch_to.frame(iframe)

# 点击“查询条件”
browser.find_element_by_xpath("//td[text()='--查询条件--']").click()

# 点击“发起人”
sleep(1)

# 如果是某些情况下拉框不可点，可以试试这个方法
# Select(browser.find_element_by_class_name("common_drop_down.w100b")).select_by_index(3)
# browser.execute_script("document.getElementsByClassName('common_drop_down w100b')[0].style.display = 'block';")

browser.find_element_by_xpath("//a[text()='发起人']").click()

# 设置每页显示1条
try:
    loc = browser.find_element_by_xpath(
        '//*[@id="rpInputChange"]')  # 定位该元素
    loc.send_keys(Keys.CONTROL + 'a')  # 全选
    loc.send_keys(Keys.DELETE)  # 删除，清空
    loc.send_keys(1)  # 写入新的值

    # 点击“GO”
    browser.find_element_by_xpath('//*[@id="grid_go"]').click()
except Exception as err:
    tkinter.messagebox.showinfo('提示', '可能有广告或消息弹窗挡住了，请关闭后再点击按钮重试')
    loc = browser.find_element_by_xpath(
        '//*[@id="rpInputChange"]')  # 定位该元素
    loc.send_keys(Keys.CONTROL + 'a')  # 全选
    loc.send_keys(Keys.DELETE)  # 删除，清空
    loc.send_keys(1)  # 写入新的值

    # 点击“GO”
    browser.find_element_by_xpath('//*[@id="grid_go"]').click()

# 开始

# 第几个表
for index, excel_dict in enumerate(sum_name):
    # 该表第几个姓名
    for a_name in excel_dict:
        # 输入姓名
        browser.find_element_by_xpath(
            '//*[@id="sender"]').clear()
        browser.find_element_by_xpath(
            '//*[@id="sender"]').send_keys(
            a_name)
        # 点击“搜索”
        browser.find_element_by_class_name('common_button.search_buttonHand').click()
        # 设置第1页
        sleep(2)
        try:
            loc2 = browser.find_element_by_xpath(
                '//*[@id="gridId_pDiv"]/div[1]/span[3]/input')  # 定位该元素
            loc2.send_keys(Keys.CONTROL + 'a')  # 全选
            loc2.send_keys(Keys.DELETE)  # 删除，清空
            loc2.send_keys(1)  # 写入新的值

            # 点击“GO”
            browser.find_element_by_xpath('//*[@id="grid_go"]').click()
        except Exception as err:
            tkinter.messagebox.showinfo('提示', '可能有广告或消息弹窗挡住了，请关闭后再点击按钮重试')
            loc2 = browser.find_element_by_xpath(
                '//*[@id="gridId_pDiv"]/div[1]/span[3]/input')  # 定位该元素
            loc2.send_keys(Keys.CONTROL + 'a')  # 全选
            loc2.send_keys(Keys.DELETE)  # 删除，清空
            loc2.send_keys(1)  # 写入新的值

            # 点击“GO”
            browser.find_element_by_xpath('//*[@id="grid_go"]').click()

        # 获取总页数
        sleep(1)
        Text = browser.find_element_by_class_name("total_page")
        a_str = Text.text
        sum_page = int(a_str[1:-1])
        i = 0
        h = 0
        while (i < sum_page):
            # handles1 = browser.window_handles
            # print("one:",handles1)
            i = i + 1
            # 判断是否为报销单
            try:

                reimbursement = browser.find_element_by_class_name("color_black").text
            except Exception as err:
                print('tiao')
                pass
            else:
                print(reimbursement, type(reimbursement))
                if "报销" in reimbursement:
                    print("在")
                    # 点击事项
                    try:
                        try:
                            print("1")
                            browser.find_element_by_xpath("//span[text()='" + a_name + "']").click()
                            print("2")
                        except Exception as err:
                            print('3')
                            browser.find_element_by_xpath("//div[text()='" + a_name + "']").click()
                            print('4')
                    except Exception as err:
                        tkinter.messagebox.showinfo('提示', '可能有网络延迟，请调整好网络点击按钮重试')
                        sleep(3)
                        if (h == 0):
                            i = i - 1
                            h = 1
                        continue
                    sleep(3)
                    while (True):
                        print(5)
                        # 切换到新网页
                        for handle in browser.window_handles:
                            browser.switch_to.window(handle)
                        print(6)
                        # 设置浏览器大小
                        browser.maximize_window()
                        try:

                            print(7)
                            # handles2 = browser.window_handles
                            # print("two:",handles2)

                            # 定位到iframe
                            iframe2 = browser.find_element_by_xpath(
                                '//*[@id="componentDiv"]')
                            # 切换到iframe
                            browser.switch_to.frame(iframe2)
                        except Exception as err:
                            tkinter.messagebox.showinfo('提示', '可能有网络延迟，请调整好网络点击按钮重试')
                            pass
                        else:
                            break

                    while (True):
                        try:
                            # 定位到iframe
                            iframe1 = browser.find_element_by_xpath(
                                '//*[@id="zwIframe"]')
                            # 切换到iframe
                            browser.switch_to.frame(iframe1)

                            # 获取编号
                            Number = WebDriverWait(browser, 7).until(
                                EC.presence_of_element_located((By.CLASS_NAME, "xdRichTextBox.validate"))

                            )
                            number = Number.text
                            print(number)

                            if number in excel_dict[a_name]:
                                print("找到了")
                                print(excel_dict[a_name])
                                excel_dict[a_name].remove(number)

                                # 点击 是
                                browser.find_element_by_xpath("(//label[text()='是'])[2]").click()
                                # 退出iframe
                                browser.switch_to.default_content()
                                # 点击同意
                                browser.find_element_by_xpath("//label[text()='同意']").click()

                                # 点击提交
                                # browser.find_element_by_xpath('//*[@id="_dealSubmit"]').click()
                                print("点击提交")
                                browser.close()
                                sleep(2)
                            else:
                                # 系统有表里没有
                                with open("" + time.strftime('%Y.%m.%d', time.localtime(time.time())) + "系统有表里没有.txt",
                                          "a+") as f:
                                    f.write(
                                        filename_s[index] + " 中 " + a_name + "的编号：" + number + " 系统里有但在表里没找到" + '\n')
                                # 关闭窗口
                                browser.close()
                                sleep(1)
                        except Exception as err:
                            tkinter.messagebox.showinfo('提示', '可能有网络延迟，请调整好网络点击按钮重试')
                            pass
                        else:
                            break

            # 定位到新页面
            for handle in browser.window_handles:
                browser.switch_to.window(handle)
            # 定位到iframe
            iframe = browser.find_element_by_xpath(
                '//*[@id="main"]')
            # 切换到iframe
            browser.switch_to.frame(iframe)
            try:
                # 点击下一页
                browser.find_element_by_xpath('//*[@id="gridId_pDiv"]/div[1]/a[3]').click()
                print("下一页")
            except Exception as err:
                tkinter.messagebox.showinfo('提示', '可能有广告或消息弹窗挡住了，请关闭后再点击按钮重试')
                # 点击下一页
                browser.find_element_by_xpath('//*[@id="gridId_pDiv"]/div[1]/a[3]').click()
            sleep(1)
            while (True):
                try:
                    # 定位到新页面
                    for handle in browser.window_handles:
                        browser.switch_to.window(handle)
                    # 定位到iframe
                    iframe = browser.find_element_by_xpath(
                        '//*[@id="main"]')
                    # 切换到iframe
                    browser.switch_to.frame(iframe)
                except Exception as err:
                    tkinter.messagebox.showinfo('提示', '可能有网络延迟，请调整好网络点击按钮重试')
                    pass
                else:
                    break

        # 看看excel_dict[a_name]的list是否为空，不为空，把列表元素写入文件说filename_s[index]什么表什么名字设么编号未找到
        if len(excel_dict[a_name]) != 0:
            for item in excel_dict[a_name]:
                with open(
                        "" + time.strftime('%Y.%m.%d', time.localtime(time.time())) + "表有系统没有.txt",
                        "a+") as f:
                    f.write(filename_s[index] + " 中 " + a_name + "的编号：" + item + " 在系统里没找到或该单子名字中不包含“报销”二字" + '\n')

tkinter.messagebox.showinfo('提示', '已完成')
