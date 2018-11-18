import os
import xlrd
import xlwt
from xlutils.copy import copy
import traceback


class ExcelHelper(object):
    def __init__(self, **kwargs):
        # 初始化参数
        self.source_table = kwargs.get("source_table")
        self.source_sheet = kwargs.get("source_sheet")
        self.title_rowx = kwargs.get("title_rowx")
        self.filter_date = kwargs.get("filter_date")
        self.filter_department = kwargs.get("filter_department")
        self.target_table_path = kwargs.get("target_table_path")
        self.target_sheet = kwargs.get("target_sheet")

    def read_data(self):
        try:
            # 1、读取数据
            # 打开表
            rd = xlrd.open_workbook(self.source_table, )
            # 打开sheet
            rtb = rd.sheet_by_name(self.source_sheet)
            # 获取sheet的行数
            tb_rows = rtb.nrows
            #
            result = []
            # 打印标题
            # print(rtb.row_values(self.title_rowx))
            # 打印第一行数据
            # print(rtb.row_values(self.title_rowx+1))
            # 原始标题
            # ['ENGLISH NAME', '日期', '大类', '物资名称', '规格型号', '部件号', '单位', '数量', '单价（SR)', '金额', '（单价）人民币', '计划单位', '备注', '合同号', '供货厂家']
            # 需要的标题
            # 大类	产品名称	规格型号	单位	数量	单价RMB	金额RMB		备注	到料时间
            # 原始数据
            # ['', 43282.0, 47.0, '接头', '样品1', '', '只', 50.0, 12.0, 600.0, 21.178584, 1.0, '', '2018-0574', 'JX-180507001']
            # 将数字转换为日期
            for row in range(self.title_rowx + 1, tb_rows):
                date = xlrd.xldate_as_datetime(rtb.cell(row, 1).value, 0).strftime("%Y/%m/%d")
                department = rtb.row_values(row)[11]
                if type(department) == str:
                    continue

                if self.filter_date == date and self.filter_department == department:
                    # 使用日期和部门进行过滤数据
                    tmp_str = "{big_type}@@@{name}@@@{version}@@@{unit}@@@{num}@@@{price_rmb}@@@{total_price_rmb}@@@{note}@@@{date}".format(
                        big_type=int(rtb.row_values(row)[2]),
                        name=rtb.row_values(row)[3],
                        version=rtb.row_values(row)[4],
                        unit=rtb.row_values(row)[6],
                        num=int(rtb.row_values(row)[7]),
                        price_rmb=round(float(rtb.row_values(row)[10]), 2),
                        total_price_rmb=round(rtb.row_values(row)[7] * rtb.row_values(row)[10], 2),
                        note=rtb.row_values(row)[12],
                        date=date
                    )
                    tmp_list = tmp_str.split('@@@')
                    fin_list = []
                    for i in tmp_list:
                        if tmp_list.index(i) == 0 or tmp_list.index(i) == 4:
                            fin_list.append(int(i))
                        elif tmp_list.index(i) == 5 or tmp_list.index(i) == 6:
                            fin_list.append(round(float(i), 2))
                        else:
                            fin_list.append(i)
                    result.append(fin_list)
            # print(result)
            return result
        except Exception as e:
            print(traceback.format_exc())

    def write_data(self, data):
        try:
            """
           1、如果没有表格，就创建表格和表头 
           2、如果有表格，就追加数据
           """
            # 创建目标文件路径，如果不存在，就新建
            if not os.path.exists(self.target_table_path):
                os.makedirs(self.target_table_path)
            # 查看文件是否存在
            file_name = os.path.join(self.target_table_path, str(self.filter_department) + "队领料单.xls")
            if os.path.exists(file_name):
                # 文件存在
                rt = xlrd.open_workbook(file_name)
                rtb = rt.sheet_by_name(self.target_sheet)
                # 获取原表格行数
                rtb_nrows = rtb.nrows
                nid = rtb.row_values(rtb_nrows - 1)[0]
                wt = copy(rt)
                wtb = wt.get_sheet(0)
                for i in range(len(data)):
                    # 增加序号
                    wtb.write(rtb_nrows + i, 0, int(nid) + i + 1)
                    # print(int(nid))
                    for j in range(len(data[i])):
                        wtb.write(rtb_nrows + i, j + 1, data[i][j])

                wt.save(file_name)
            else:
                # 文件不存在
                # 增加表头
                theads = ["序号", "大类", "产品名称", "规格型号", "单位", "数量", "单价RMB", "金额RMB", "备注", "到料时间"]
                wt = xlwt.Workbook()
                wtb = wt.add_sheet(self.target_sheet)
                wtb.write_merge(0, 0, 0, 9, "沙特公司基层单位领料确认单")
                wtb.write_merge(1, 1, 0, 9, "ZPEBINT/JL-SA-YMD-005-2014")
                wtb.write_merge(2, 2, 0, 9, "单位： %s %s" % (self.filter_department, data[0][8]))
                for i in range(len(theads)):
                    wtb.write(3, i, theads[i])
                # 增加数据
                for i in range(len(data)):
                    # 增加序号
                    wtb.write(i + 4, 0, i + 1)
                    for j in range(len(data[i])):
                        # 增加真是数据
                        wtb.write(i + 4, j + 1, data[i][j])

                # 保存文件
                wt.save(file_name)
        except Exception as e:
            print(traceback.format_exc())


if __name__ == "__main__":
    """
    1、python中都是从0开始计数的

    2、source_table:源数据表
    3、source_sheet:源数据表中的sheet名字
    4、title_rowx:源数据表中标题所在的行数
    5、filter_date:源数据表中使用此参数进行日期过滤
    6、filter_department:源数据表中使用此参数进行部门过滤
    7、target_table_path:目标保存数据表的目录
    8、target_sheet:目标数据保存sheet
    """
    # -----------------------------------------------
    """
    需修改的参数
    """
    source_table = r"E:\test\2018年收料记录 - 副本.xlsx"
    source_sheet = "2018年收料记录"
    target_table_path = r"E:\test\new"
    departments = [12]
    filter_date = "2018/07/05"
    target_sheet = "境外未领"
    # -----------------------------------------------
    try:
        for department in departments:
            eh = ExcelHelper(source_table=source_table,
                             source_sheet=source_sheet,
                             title_rowx=2,
                             filter_date=filter_date,
                             filter_department=department,
                             target_table_path=target_table_path,
                             target_sheet=target_sheet)
            data = eh.read_data()
            # print(data)
            eh.write_data(data)
    except Exception as e:
        print(traceback.format_exc())
