from openpyxl import load_workbook
from tkinter import Tk
from re import match
from warnings import filterwarnings
from win32api import MessageBox
from win32con import MB_OK
from tkinter import filedialog

# 忽略警告
filterwarnings("ignore")

# 消除TK模块的小窗口
root = Tk()
root.withdraw()

# 文件路径提取（限制文件类型）
file_path = filedialog.askopenfilename(filetypes=[('XLSX','xlsx')])

if file_path.strip() == '':
    # 第一个参数0，不知道啥意思，第二个参数是框体文本，第三个是标题文本，第四个是反馈键
    MessageBox(0, '         未选择文件', '', MB_OK)

else:
    try:
        # 获取sheet页,此处是list对象
        wb1 = load_workbook(file_path, data_only=True)
        sheets1 = wb1.sheetnames

        # 此处是sheet对象
        sheet1 = wb1[sheets1[0]]

        # 最大行数
        max_row = sheet1.max_row

        # 最大列数
        max_column = sheet1.max_column

        # 选取特定列（考勤号码、日期、签到时间、签退时间）
        zlist = [2, 4, 8, 9, ]

        number_list = []

        for m in range(3, max_row+1):
            serial_num = "%08d" % int(sheet1.cell(m, 2).value)
            date_obj = match(r'^(\d{4})-(\d{1,2})-(\d{1,2})$', str(sheet1.cell(m, 4).value))
            year = date_obj.group(1)
            mon = "%02d" % int(date_obj.group(2))
            day = "%02d" % int(date_obj.group(3))
            date = year + mon + day
            clock_in = sheet1.cell(m, 8).value

            if clock_in == None:
                pass

            else:
                time_obj = match(r'^(\d{2}):(\d{2}):(\d{2})$', clock_in)
                convert_data = 'P100001' + date + str(time_obj.group(1)) + str(time_obj.group(2)) + str(time_obj.group(3)) + date + str(time_obj.group(1)) + \
                    str(time_obj.group(2)) + str(time_obj.group(3)) + serial_num
                number_list.append(convert_data)
            clock_out = sheet1.cell(m, 9).value

            if clock_out == None:
                pass

            else:
                time_obj = match(r'^(\d{2}):(\d{2}):(\d{2})$', clock_out)
                convert_data = 'P200001' + date + str(time_obj.group(1)) + str(time_obj.group(2)) + str(time_obj.group(3)) + date + str(time_obj.group(1)) + \
                    str(time_obj.group(2)) + str(time_obj.group(3)) + serial_num
                number_list.append(convert_data)
        # 关闭excel
        wb1.close()

    except (TypeError,ValueError,AttributeError):
        # 异常处理
        MessageBox(0, '         文档内容无效', '', MB_OK)

    else:
        filename = number_list[0][7:11] + '年' + number_list[0][11:13] + '月' + '考勤转换数据.txt'
        file = open(filename, 'w')
        file.write('\n'.join(number_list))
        file.close()
        MessageBox(0, '           执行完毕', '', MB_OK)
