import xlrd
import xlwt
import time
from datetime import datetime



# workbook = xlrd.open_workbook(r"E:\test\kltest.xlsx")
# sheet = workbook.sheets()[0]  # 读取第一个sheet
# nrows = sheet.nrows  # 行数
# first_row_values = sheet.row_values(0)  # 第一行数据
# list = []
# num = 1
# for row_num in range(1, nrows):
#     row_values = sheet.row_values(row_num)
#     # if row_values:
#     #     str_obj = {}
#     for i in range(len(first_row_values)):
#         ctype = sheet.cell(num, i).ctype
#         cell = sheet.cell_value(num, i)
#         print(cell, ctype)
#         if ctype == 2 and cell % 1 == 0.0:  # ctype为2且为浮点
#             cell = int(cell)  # 浮点转成整型
#             cell = str(cell)  # 转成整型后再转成字符串，如果想要整型就去掉该行
#         elif ctype == 3:
#             date = datetime(*xlrd(cell, 0))
#             cell = date.strftime('%Y/%m/%d %H:%M:%S')
#         elif ctype == 4:
#             cell = True if cell == 1 else False
#         str_obj[first_row_values[i]] = cell
#     list.append(str_obj)
#     num = num + 1
# print(list)

# print(sheet.cell(1, 0))
# print(xlrd.xldate_as_datetime(sheet.cell(1, 0).value, 0))

# theads = ["序号", "大类", "产品名称", "规格型号", "单位", "数量", "单价RMB", "金额RMB", "备注", "到料时间"]
# wt = xlwt.Workbook()
# wtb = wt.add_sheet("Sheet1")
# wtb.write_merge(0, 0, 0, 9, "沙特公司基层单位领料确认单")
# wtb.write_merge(1, 1, 0, 9, "ZPEBINT/JL-SA-YMD-005-2014")
# wtb.write_merge(2, 2, 0, 9, "单位： %s %s" % ("test_department", "2018/08/17"))
# for i in range(len(theads)):
#     wtb.write(3, i, theads[i])
# wt.save(r"e:\test\abc.xls")

# rt = xlrd.open_workbook(r"E:\test\2018年收料记录 - 副本.xlsx")
# rtb = rt.sheet_by_name("2018年收料记录")
# if type(rtb.row_values(8)[11]) == str:
#     print(111)
# else:
#     print(222)

rt = xlrd.open_workbook(r"E:\test\new\5队领料单.xls")
rtb = rt.sheet_by_name("境外未领")

result = []
print(rtb.row_values(5))