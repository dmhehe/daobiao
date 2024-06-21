import re
import json
from openpyxl import load_workbook



ATTR_ROW = 6 #属性行，前6行都是属性行，没有数据的
variable_name_pattern = re.compile(r'^[a-zA-Z_]\w*$')

def getSheetName(xls_file_path):
    # 打开一个 xlsx 文件
    workbook = load_workbook(filename=xls_file_path, data_only=True)
    
    # 获取所有工作表的名字
    sheets_names = workbook.sheetnames
    
    # 定义一个正则表达式匹配符合变量命名规则的名称
    # 变量名可以使用字母、数字、下划线，但不能以数字开头
    variable_name_pattern = re.compile(r'^[a-zA-Z_][a-zA-Z0-9_]*$')
    
    # 过滤掉以 "Sheet" 开头和不符合变量命名规则的工作表名称
    filtered_sheet_names = [
        name for name in sheets_names
        if not name.startswith("Sheet") and variable_name_pattern.match(name)
    ]
    
    return filtered_sheet_names

#留下有用的部分
def filter_usedata(data):
    if len(data) < 5:
        raise Exception("表数据有问题1！！小于5行")

#    for i, list1 in enumerate(data):
#        print("4444444444444", i, list1)


    max_col = len(data[0])
    for i in range(max_col):
        list1 = []
        for j in range(5):
            list1.append(data[j][i].strip())
        if all(s == "" for s in list1):
            max_col = i
            break
        
    max_row = len(data)
    for i, row_data in enumerate(data):
        list1 = []
        for j in range(max_col):
            list1.append(row_data[j])
        if all(s == "" for s in list1):
            max_row = i
            break

    new_data = []

    for i in range(max_row):
        new_line = []
        new_data.append(new_line)
        for j in range(max_col):
            new_line.append(data[i][j])


    if len(new_data) < 5:
        raise Exception("表有效数据有问题2！！小于5行")



    # print("6666666666666666", new_data)
    return new_data



# #获取xls 里面 某个表 里面数据 二维字符串返回
# def read_sheet_data(xls_file_path, sheet_name):
#     # 打开 xls 文件
#     workbook = xlrd.open_workbook(xls_file_path)
    
#     # 通过名称获取工作表
#     sheet = workbook.sheet_by_name(sheet_name)
    
#     # 读取工作表数据到二维字符串列表
#     data = []
#     for row_idx in range(sheet.nrows):
#         row_data = []
#         for col_idx in range(sheet.ncols):
#             # 读取单元格数据并转换为字符串
#             cell_value = sheet.cell_value(row_idx, col_idx)
#             # 处理数据类型，确保结果是字符串
#             print("ssssssssss", cell_value)
#             if isinstance(cell_value, str):
#                 row_data.append(cell_value.strip())
#             else:
#                 # 对于非字符串类型，如数字，转换为字符串
#                 row_data.append(str(cell_value).strip())
#         data.append(row_data)
#     data = filter_usedata(data)
#     return data


def read_sheet_data(xls_file_path, sheet_name):
    # 打开 xlsx 文件
    workbook = load_workbook(filename=xls_file_path, data_only=True)
    
    # 通过名称获取工作表
    sheet = workbook[sheet_name]
    
    # 读取工作表数据到二维字符串列表
    data = []
    for row in sheet.iter_rows(values_only=True):
        row_data = []
        for cell_value in row:
            # 处理数据类型，确保结果是字符串
            print("ssssssssss", cell_value)
            if isinstance(cell_value, str):
                row_data.append(cell_value.strip())
            elif cell_value is None:
                row_data.append('')
            else:
                # 对于非字符串类型，如数字，转换为字符串
                row_data.append(str(cell_value).strip())
        data.append(row_data)
    
    data = filter_usedata(data)
    return data


#返回属性字典，‘1,a=2,b=3’
def getAttrDict(text, mainName="main"):
    text = text.replace("，", ",")
    data_dict = {}
    attr_pair_list = text.split(",")
    for attr_pair_text in attr_pair_list:
        txt1 = attr_pair_text.strip()
        if not txt1:
            continue
        
        if txt1.find("=") >= 0:
            list1 = txt1.split("=")
            data_dict[list1[0]] = list1[1]
        else:
            data_dict[mainName] = txt1

    return data_dict


def getNumList(text):
    text = text.replace("，", ",")
    return eval("["  + text + "]")


def getStringList(text):
    text = text.replace("，", ",")
    list1 = text.split(",")
    list2 = []
    for str1 in list1:
        list2.append(str1.strip())

    return list2


def convert_string_to_number(s):
    try:
        # 尝试将字符串转换为整数
        return int(s)
    except ValueError:
        try:
            # 如果转换为整数失败，尝试转换为浮点数
            return float(s)
        except ValueError:
            # 如果转换为整数和浮点数都失败，返回原始字符串
            return s


g_Start_Flag = "//****&&*****start****&&****"
g_End_Flag = "//****&&*****end****&&****"
def writeFile(file_path, str1):
    content = ""
    # 打开并读取文件
    with open(file_path, 'r', encoding='utf-8') as file1:
        # 读取文件的全部内容
        content = file1.read()

    start_text = ""
    end_text = "" 
    start_idx = content.index("g_Start_Flag")
    end_idx = content.index("g_End_Flag")
    if start_idx >= 0:
        start_text = content[0:start_idx]

    if end_idx >= 0:
        end_text = content[end_idx+len(g_End_Flag):]


    text = start_text + g_Start_Flag + "\n" + str1 + "\n" + g_End_Flag + end_text

    with open(file_path, 'w', encoding='utf-8') as file2:
        # 写入文本
        file2.write(text)

def writeJson(file_path, data):
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

