# coding=utf-8
# pip install xlrd==1.2.0 -i http://pypi.douban.com/simple/ --trusted-host pypi.douban.com
# pip install xlwt -i http://pypi.douban.com/simple/ --trusted-host pypi.douban.com
# pyinstaller -F xxxx.py
# pyinstaller -F -i bb.ico student_compare.py -n 疫情统计 --noconsole
# pip install pyinstaller -i http://pypi.douban.com/simple/ --trusted-host pypi.douban.com
# pip install -i https://pypi.tuna.tsinghua.edu.cn/simple pyinstaller

"""
try:
    __version__ = version(__name__)
except Exception:
    pass
"""
# __version__="0.2.5"   # 这行是用来解决python库报错的问题

from os.path import abspath, dirname, join
import sys
from urllib.request import urlretrieve
import xlrd
import xlwt
import os
from datetime import datetime
import tkinter as tk
import tkinter.messagebox  # 这个是消息框，对话框的关键
from requests_html import HTMLSession

excel_name = "student__down_1521.xlsx"


# 爬取中高风险地区列表
def get_dangerous_area():
    url = 'http://www.gd.gov.cn/gdywdt/zwzt/yqfk/content/post_3021711.html'

    session = HTMLSession()
    res = session.get(url)

    print("爬取风险地区的url= {}".format(url))
    # 地区html列表
    area_list = res.html.xpath("//div[@class='zw']/p")

    # 处理地区列表
    new_area_list = []
    for area in area_list:
        areaName = "".join([a.strip() for a in area.xpath('//text()')])
        # print("发布的危险地区：{}".format(areaName))
        new_area_list.append(areaName)
    new_area_list1 = [area for area in new_area_list if area != '']

    mid_sentence_index = new_area_list1.index("中风险地区：")
    high_area_list = new_area_list1[1:mid_sentence_index]  # 高风险地区列表
    mid_area_list = new_area_list1[mid_sentence_index + 1:]  # 中风险地区列表
    # print(new_area_list1)
    print("高风险 ->{}".format(high_area_list))
    print("中风险 ->{}".format(mid_area_list))
    return high_area_list, mid_area_list


# 下载excel
def down_excel(url):
    print("下载的excel的url-> {}".format(url))
    del_old_excel()
    global excel_name
    try:
        urlretrieve(url, excel_name)
        line = ""
        try:
            f = open(excel_name)
            line = f.readline()
            print("下载的excel文件第一行内容-> {}".format(line))
        except:
            print("Excel成功下载！")
        if line.__contains__("500") and line.__contains__("code"):
            print("Excel下载失败！url链接过期")
            raise RuntimeError('下载excel文件失败，抛出异常')
    except Exception as e:
        import traceback
        traceback.format_exc()
        print("Excel下载失败！，url错误")
        raise RuntimeError('下载excel文件失败，抛出异常')


#  读取Excel列表并进行对比
def analyse_student():
    global excel_name
    dangerous_areas = get_dangerous_area()

    student_excel = xlrd.open_workbook(excel_name)
    sign_sheet = student_excel.sheet_by_name("已签到")
    high_list = [['高风险地区']]
    mid_list = [['中风险地区']]
    extra_list = [['港澳地区']]

    for i in range(sign_sheet.nrows):
        row = sign_sheet.row_values(i)
        area_arr = sign_sheet.cell_value(i, 10), sign_sheet.cell_value(i, 11), sign_sheet.cell_value(i, 12)
        area_arr = [area for area in area_arr if area.strip() != '']

        # 判断该学生所在地区是否包含在高危险地区内
        if is_dangerous(dangerous_areas[0], area_arr):
            print("高风险地区：-> {}".format(row))
            high_list.append(row)
        # 判断该学生所在地区是否包含在中危险地区内
        if is_dangerous(dangerous_areas[1], area_arr):
            print("中风险地区：-> {}".format(row))
            mid_list.append(row)
        if is_extra_area(area_arr):
            print("港澳风险地区：-> {}".format(row))
            extra_list.append(row)
    result_excel = xlwt.Workbook()
    sheet = result_excel.add_sheet("汇总")
    rowNum = write_excel(sheet, [sign_sheet.row_values(0)])
    rowNum = write_excel(sheet, high_list, rowNum)
    rowNum = write_excel(sheet, mid_list, rowNum + 2)
    rowNum = write_excel(sheet, extra_list, rowNum + 2)
    file_name = "{}.xls".format(get_excel_time())
    try:
        result_excel.save(file_name)
        return True
    except Exception as e:
        import traceback
        traceback.format_exc()
        print("保存失败，将以新的文件名保存！")
        result_excel.save("{}_1.xls".format(get_excel_time()))


# 判断该学生所在地区是否包含在危险地区内
def is_dangerous(dangerous_areas, area_arr):
    for dangerous_area in dangerous_areas:
        flag = True if len(area_arr) > 0 else False
        for area1 in area_arr:
            if not dangerous_area.__contains__(area1):
                flag = False
        if flag:
            return True
    return False


# 判断是否是港澳地区
def is_extra_area(area_arr):
    extra_areas = ("香港特别行政区", "澳门特别行政区")
    for extra_area in extra_areas:
        for area1 in area_arr:
            if area1.__contains__(extra_area):
                return True
    return False


# 获取统计表生成的时间
def get_excel_time():
    student_excel = xlrd.open_workbook(excel_name)
    one_sheet = student_excel.sheet_by_name("综合")
    # return "{}".format(one_sheet.cell_value(1, 0)).split(" ")[0]
    return "{}({})".format(one_sheet.cell_value(0, 0), one_sheet.cell_value(1, 0).split(" ")[0])


# 从指定行号开始写,返回写完之后的行号
def write_excel(sheet, rowList, start_row_num=0):
    for i in range(len(rowList)):
        row = rowList[i]
        for j in range(len(row)):
            sheet.write(start_row_num + i, j, row[j])

    return start_row_num + len(rowList)


# 读取文件中的url
def get_url():
    try:
        f = open("url.txt")
        line = f.readline()
        print("下载的excel的url-> {}".format(line))
        return line
    except Exception as e:
        import traceback
        traceback.format_exc()
        print("读取url文件失败")

# 删除旧excel
def del_old_excel():
    if os.path.exists(excel_name):
        os.remove(excel_name)

def show_gui():
    window = tk.Tk()
    w, h = window.winfo_screenwidth(), window.winfo_screenheight()
    window.title('统计学生疫情地区信息')
    window.geometry('500x100+{}+{}'.format(int((w - 500) / 2), int((h - 100) / 2)))

    # user information
    tk.Label(window, text='请输入Excel下载地址: ').place(x=50, y=20)

    var_usr_name = tk.StringVar()
    var_usr_name.set('example@python.com')
    entry_usr_name = tk.Entry(window, textvariable=var_usr_name, fg="red")
    entry_usr_name.place(x=180, y=20, width=300, height=20)

    def exec():
        print(datetime.now())
        excel_url = var_usr_name.get()
        retult = False

        try:
            down_excel(excel_url)
            retult = analyse_student()
        except Exception as e:
            tk.messagebox.showerror('Error', '下载excel文件失败，请输入正确未失效的Excel链接')
        if retult:
            tk.messagebox.showinfo(title='结果提示', message='Success! ')
        del_old_excel()

    # login and sign up button
    btn_login = tk.Button(window, text=' 确定 ', command=exec)
    btn_login.place(x=200, y=60)

    window.mainloop()

if __name__ == '__main__':
    log_file = open('log.txt', 'w')
    sys.stdout = log_file
    print(datetime.now())
    show_gui()
