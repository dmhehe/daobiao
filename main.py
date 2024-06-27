from typing import Any
import xls_tool

import os
import zipfile


class SheetData:
    def __init__(self, xls_file_path, sheet_name) -> None:
        self.m_line_data_list = []
        self.file_name = sheet_name  #文件名就是 表名
        
        
        self.m_xls_data = xls_tool.read_sheet_data(xls_file_path, sheet_name)
        xls_data = self.m_xls_data
        
        
        self.m_colum = len(xls_data[0])
        self.m_row = len(xls_data)


        xls_attr_dict = {}#整个表的属性
        for i in range(self.m_colum):
            attr_dict = xls_tool.getAttrDict(xls_data[0][i])
            xls_attr_dict.update(attr_dict)
        self.m_xls_attr_dict = xls_attr_dict

        self.m_attr_name_list = []#属性名字
        for i in range(self.m_colum):
            attr_dict = xls_tool.getAttrDict(xls_data[1][i])
            self.m_attr_name_list.append(attr_dict)


        self.m_client_attr_list = []#客户端属性
        for i in range(self.m_colum):
            attr_dict = xls_tool.getAttrDict(xls_data[2][i])
            self.m_client_attr_list.append(attr_dict)


        self.m_server_attr_list = []#服务端属性
        for i in range(self.m_colum):
            attr_dict = xls_tool.getAttrDict(xls_data[3][i])
            self.m_server_attr_list.append(attr_dict)

        self.m_type_list = []#类型
        for i in range(self.m_colum):
            attr_dict = xls_tool.getAttrDict(xls_data[4][i])
            self.m_type_list.append(attr_dict)


        self.m_desc_list = []#描述
        for i in range(self.m_colum):
            attr_dict = xls_tool.getAttrDict(xls_data[5][i])
            self.m_desc_list.append(attr_dict)

    def addOneLineData(self, data):
        self.m_line_data_list.append(data)

    def isNewLineData(self, data):
        return data in self.m_line_data_list

    def getAttrName(self, i, j, isClient=True):
        attrName = self.m_attr_name_list[j]["main"]
        if isClient:
            useName = self.m_client_attr_list[j]["main"]
        else:
            useName = self.m_server_attr_list[j]["main"]

        if useName not in ["1", "0"]:
            return useName
        
        return attrName
    

    def getRawValue(self, i, j, isClient=True):
        value_str = self.m_xls_data[i][j]
        return value_str
    
    def getTsValueType(self, j, isClient=True):
         typeName = self.m_type_list[j]["main"]
         if typeName == "num":
             return "number"
         elif typeName == "str":
             return "string"
         elif typeName == "str[]":
            return "string[]"
         elif typeName == "num[]":
            return "number[]"
         else:
             return "any"
        
    def getValue(self, i, j, isClient=True):
        typeName = self.m_type_list[j]["main"]
        value_str = self.m_xls_data[i][j]
        if typeName == "num":
            if value_str == "":
                return 0
            
            return xls_tool.convert_string_to_number(value_str)
        
        elif typeName == "str":
            return self.m_xls_data[i][j]

        elif typeName == "str[]":
            if value_str == "":
                return []
            
            ans_list = xls_tool.getStringList(value_str)
            self.addOneLineData(ans_list)
            return ans_list
        
        elif typeName == "num[]":
            if value_str == "":
                return []
            
            ans_list = xls_tool.getNumList(value_str) 
            self.addOneLineData(ans_list)
            return ans_list
        
        elif typeName in ("raw", ""):
            if value_str == "":
                return None
            
            return xls_tool.convert_string_to_number(value_str)
    
    def isUseColum(self, i, isClient=True):
        if isClient:
            useName = self.m_client_attr_list[i]["main"]
        else:
            useName = self.m_server_attr_list[i]["main"]

        if useName == "0":
            return False
        
        return True


    def getTransAllData(self, isClient=True):
        mainType = self.m_xls_attr_dict["main"]
        if mainType == "1": #第一种类型  {key:obj}
            return self.getDataByType1(isClient)
        elif mainType == "2": #[obj]
            return self.getDataByType2(isClient)
        elif mainType == "3":#{key:{key2:obj}}
            return self.getDataByType3(isClient)
        elif mainType == "4":#{key:[obj]}
            return self.getDataByType4(isClient)
        elif mainType == "5":#{key:{key2:{key3:obj}}}
            return self.getDataByType5(isClient)
        elif mainType == "6":#{key:{key2:[obj]}}
            return self.getDataByType6(isClient)

    
    def getDataByType1(self, isClient=True):
        ans_data = {}
        
        useKeyDict = {}
        for i in range(6, self.m_row):
            line_data = {}
            key = None
            for j in range(0, self.m_colum):
                if not self.isUseColum(j, isClient):
                    continue

                value = self.getValue(i, j, isClient)
                valueName = self.getAttrName(i, j, isClient)
                line_data[valueName] = value
                if j == 0:
                    key = value

            if key == None:
                raise Exception(f"在({i}, {j}), 关键key为空")
            
            if key in useKeyDict:
                raise Exception(f"出现了重复的key {key}")
            else:
                useKeyDict[key] = 1
            self.addOneLineData(line_data)
            ans_data[key] = line_data
        return ans_data
    
    def getDataByType2(self, isClient=True):
        ans_data = []
        for i in range(6, self.m_row):
            line_data = {}
            for j in range(0, self.m_colum):

                if not self.isUseColum(j, isClient):
                    continue

                value = self.getValue(i, j, isClient)
                valueName = self.getAttrName(i, j, isClient)
                line_data[valueName] = value
            self.addOneLineData(line_data)
            ans_data.append(line_data)
        return ans_data
    

    def getDataByType3(self, isClient=True):
        ans_data = {}
        useKeyDict = {}

        curDict = None
        curVal = None
        for i in range(6, self.m_row):
            id_val = self.getValue(i, 0, isClient)
            id_raw = self.getRawValue(i, 0, isClient)
            id_name = self.getAttrName(i, 0, isClient)
            
            if id_raw != "":
                curDict = {}
                ans_data[id_val] = curDict
                curVal = id_val
                
                useKeyDict = {}

            line_data = {id_name:curVal}

            key = None
            for j in range(1, self.m_colum):
                if not self.isUseColum(j, isClient):
                    continue

                value = self.getValue(i, j, isClient)
                valueName = self.getAttrName(i, j, isClient)
                line_data[valueName] = value
                if j == 1:
                    key = value

            if key == None:
                raise Exception(f"在({i}, {j}), 关键key为空")
            
            if key in useKeyDict:
                raise Exception(f"出现了重复的key {key}")
            else:
                useKeyDict[key] = 1
            self.addOneLineData(line_data)
            curDict[key] = line_data
        return ans_data
    
    def getDataByType4(self, isClient=True):
        ans_data = {}
        

        curList = None
        curVal = None
        for i in range(6, self.m_row):
            id_val = self.getValue(i, 0, isClient)
            id_raw = self.getRawValue(i, 0, isClient)
            id_name = self.getAttrName(i, 0, isClient)

            if id_raw != "":
                curList = []
                ans_data[id_val] = curList
                curVal = id_val

            line_data = {id_name:curVal}

            for j in range(1, self.m_colum):
                if not self.isUseColum(j, isClient):
                    continue

                value = self.getValue(i, j, isClient)
                valueName = self.getAttrName(i, j, isClient)
                line_data[valueName] = value
            self.addOneLineData(line_data)
            curList.append(line_data)
        return ans_data

    def getDataByType5(self, isClient=True):
        ans_data = {}
        useKeyDict = {}

        parentDict = None
        curDict2 = None
        curVal1 = None
        curVal2 = None
        
        for i in range(6, self.m_row):
            id_val = self.getValue(i, 0, isClient)
            id_raw = self.getRawValue(i, 0, isClient)
            id_name = self.getAttrName(i, 0, isClient)
            id_val2 = self.getValue(i, 1, isClient)
            id_raw2 = self.getRawValue(i, 1, isClient)
            id_name2 = self.getAttrName(i, 1, isClient)

            if id_raw != "":
                parentDict = {}
                ans_data[id_val] = parentDict
                curVal1 = id_val

            if id_raw2 != "":
                curDict2 = {}
                parentDict[id_val2] = curDict2
                curVal2 = id_val2
                useKeyDict = {}

            line_data = {id_name:curVal1, id_name2:curVal2}

            key = None
            for j in range(2, self.m_colum):
                if not self.isUseColum(j, isClient):
                    continue

                value = self.getValue(i, j, isClient)
                valueName = self.getAttrName(i, j, isClient)
                line_data[valueName] = value
                if j == 2:
                    key = value

            if key == None:
                raise Exception(f"在({i}, {j}), 关键key为空")
            
            if key in useKeyDict:
                raise Exception(f"出现了重复的key {key}")
            else:
                useKeyDict[key] = 1
            self.addOneLineData(line_data)
            curDict2[key] = line_data
        return ans_data
    

    def getDataByType6(self, isClient=True):
        ans_data = {}

        parentDict = None
        curList2 = None
        curVal1 = None
        curVal2 = None
        for i in range(6, self.m_row):
            id_val = self.getValue(i, 0, isClient)
            id_raw = self.getRawValue(i, 0, isClient)
            id_name = self.getAttrName(i, 0, isClient)
            id_val2 = self.getValue(i, 1, isClient)
            id_raw2 = self.getRawValue(i, 1, isClient)
            id_name2 = self.getAttrName(i, 1, isClient)

            if id_raw != "":
                parentDict = {}
                ans_data[id_val] = parentDict
                curVal1 = id_val

            if id_raw2 != "":
                curList2 = []
                parentDict[id_val2] = curList2
                curVal2 = id_val2

            line_data = {id_name:curVal1, id_name2:curVal2}

            for j in range(2, self.m_colum):
                if not self.isUseColum(j, isClient):
                    continue

                value = self.getValue(i, j, isClient)
                valueName = self.getAttrName(i, j, isClient)
                line_data[valueName] = value
            self.addOneLineData(line_data)
            curList2.append(line_data)
        return ans_data
    
    

    
    
    
    def getTsStatementData(self):
        mainType = self.m_xls_attr_dict["main"]
        if mainType == "1":
            return self.getTsStatement1(True)
        elif mainType == "2":
            return self.getTsStatement2(True)
        elif mainType == "3":
            return self.getTsStatement3(True)
        elif mainType == "4":
            return self.getTsStatement4(True)
        elif mainType == "5":#{key:{key2:{key3:obj}}}
            return self.getTsStatement5(True)
        elif mainType == "6":#{key:{key2:[obj]}}
            return self.getTsStatement6(True)
    
    
    def getTsStatement1(self, isClient=True):
        
# interface MyObject {
#     property1: string;
#     property2: number;
#     // 可以根据需要添加更多属性
# }

# // 声明一个接口，其中键是字符串，对应的值是 MyObject 类型的对象
# export interface MyObjectMap {
#     [key: string]: MyObject;
# }
        ListBBB = []
        for j in range(0, self.m_colum):
            if not self.isUseColum(j, isClient):
                continue
            value = self.getTsValueType(j, isClient)
            valueName = self.getAttrName(0, j, isClient)
            ListBBB.append("    " + valueName + ": " + value + ";")
        
        mapText = """export interface AAADataMap {
    [key: string]: AAAData;
}"""
        objText = """export interface AAAData {
    BBB
}"""

        AAA = self.file_name

        mapText = mapText.replace("AAA", AAA)
        objText = objText.replace("AAA", AAA).replace("BBB", "\n".join(ListBBB))
        return objText +"\n\n"+ mapText


    def getTsStatement2(self, isClient=True):
        
# interface MyObject {
#     property1: string;
#     property2: number;
#     // 可以根据需要添加更多属性
# }
# export type MyObjectArray = MyObject[];
        ListBBB = []
        for j in range(0, self.m_colum):
            if not self.isUseColum(j, isClient):
                continue
            value = self.getTsValueType(j, isClient)
            valueName = self.getAttrName(0, j, isClient)
            ListBBB.append("    " + valueName + ": " + value + ";")
        
        mapText = """export type AAADataArray = AAAData[];"""
        objText = """export interface AAAData {
    BBB
}"""

        AAA = self.file_name

        mapText = mapText.replace("AAA", AAA)
        objText = objText.replace("AAA", AAA).replace("BBB", "\n".join(ListBBB))
        return objText +"\n\n"+ mapText
        
        

    def getTsStatement3(self, isClient=True):
        ListBBB = []
        for j in range(0, self.m_colum):
            if not self.isUseColum(j, isClient):
                continue
            value = self.getTsValueType(j, isClient)
            valueName = self.getAttrName(0, j, isClient)
            ListBBB.append("    " + valueName + ": " + value + ";")
        
        mapText = """export interface AAADataMap {
    [key: string]: AAADataMap2;
}"""

        map2Text = """export interface AAADataMap2 {
    [key: string]: AAAData;
}"""

        objText = """export interface AAAData {
    BBB
}"""

        AAA = self.file_name

        mapText = mapText.replace("AAA", AAA)
        map2Text = map2Text.replace("AAA", AAA)
        objText = objText.replace("AAA", AAA).replace("BBB", "\n".join(ListBBB))
        return objText +"\n\n"+ map2Text + "\n\n"+ mapText
    

    def getTsStatement4(self, isClient=True):
        ListBBB = []
        for j in range(0, self.m_colum):
            if not self.isUseColum(j, isClient):
                continue
            value = self.getTsValueType(j, isClient)
            valueName = self.getAttrName(0, j, isClient)
            ListBBB.append("    " + valueName + ": " + value + ";")
        map2Text = """export interface AAADataMap {
    [key: string]: AAADataArray;
}"""
        mapText = """export type AAADataArray = AAAData[];"""
        objText = """export interface AAAData {
    BBB
}"""

        AAA = self.file_name

        mapText = mapText.replace("AAA", AAA)
        map2Text = map2Text.replace("AAA", AAA)
        objText = objText.replace("AAA", AAA).replace("BBB", "\n".join(ListBBB))
        return objText +"\n\n"+ mapText+"\n\n"+ map2Text

    def getTsStatement5(self, isClient=True):
        ListBBB = []
        for j in range(0, self.m_colum):
            if not self.isUseColum(j, isClient):
                continue
            value = self.getTsValueType(j, isClient)
            valueName = self.getAttrName(0, j, isClient)
            ListBBB.append("    " + valueName + ": " + value + ";")
        
        mapText = """export interface AAADataMap {
    [key: string]: AAADataMap2;
}"""

        map2Text = """export interface AAADataMap2 {
    [key: string]: AAADataMap3;
}"""

        map3Text = """export interface AAADataMap3 {
    [key: string]: AAAData;
}"""



        objText = """export interface AAAData {
    BBB
}"""

        AAA = self.file_name

        mapText = mapText.replace("AAA", AAA)
        map2Text = map2Text.replace("AAA", AAA)
        map3Text = map3Text.replace("AAA", AAA)
        objText = objText.replace("AAA", AAA).replace("BBB", "\n".join(ListBBB))
        return objText +"\n\n"+ map3Text + "\n\n"+ map2Text + "\n\n"+ mapText


    def getTsStatement6(self, isClient=True):
        ListBBB = []
        for j in range(0, self.m_colum):
            if not self.isUseColum(j, isClient):
                continue
            value = self.getTsValueType(j, isClient)
            valueName = self.getAttrName(0, j, isClient)
            ListBBB.append("    " + valueName + ": " + value + ";")
        map3Text = """export interface AAADataMap {
    [key: string]: AAADataMap2;
}"""

        map2Text = """export interface AAADataMap2 {
    [key: string]: AAADataArray;
}"""
        mapText = """export type AAADataArray = AAAData[];"""
        objText = """export interface AAAData {
    BBB
}"""

        AAA = self.file_name

        mapText = mapText.replace("AAA", AAA)
        map2Text = map2Text.replace("AAA", AAA)
        map3Text = map3Text.replace("AAA", AAA)
        objText = objText.replace("AAA", AAA).replace("BBB", "\n".join(ListBBB))
        return objText +"\n\n"+ mapText+"\n\n"+ map2Text+"\n\n"+ map3Text
        



    def makeJsonFile(self, file_path, sheet_name, isClient=True):
        
        txt = self.getTransAllData(isClient)
        xls_tool.writeJson(file_path, txt)
        
        
    def makeLuaFile(self, file_path, sheet_name, isClient=True):
        startText = """daobiao = daobiao or {}\ndaobiao.AAA = """.replace("AAA", sheet_name)
        data = self.getTransAllData(isClient)
        xls_tool.writeLua(file_path, data, startText, self.m_line_data_list)
        
    def makeTSFile(self, file_path, sheet_name, isClient=True):
        txt = ""
        if isClient:
            txt = self.getTsStatementData()
        xls_tool.writeFile(file_path, txt)
    






def compress_json_files(folder_path, output_zip_path):
    
    directory = os.path.dirname(output_zip_path)
    if not os.path.exists(directory):
        os.makedirs(directory)
    
    # 创建一个新的 zip 文件
    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # 遍历文件夹中的所有文件和子文件夹
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                # 获取文件的完整路径
                file_path = os.path.join(root, file)
                # 将文件添加到 zip 文件中，并使用相对路径
                # zipfile 处理文件名时默认使用 UTF-8 编码
                zipf.write(file_path, os.path.relpath(file_path, folder_path))
                
                


#生成json的文件夹，项目里面没有用的，给看看参考而已
g_JsonPath = "D:/daobiao/json"

#生成ts声明代码的地方，项目有用的，因为是声明，上线编译成js就会没有掉的了
g_TsPath = "D:/daobiao/ts"

#bin文件就是所有json的zip压缩算法的压缩包，项目有用的， 要放到项目daobiao文件夹里面 
g_PackPath = "D:/daobiao/pack"

#xlsx的文件夹
g_XlsxFloderPath = "D:/daobiao/xlsx"



#lua的文件夹
g_LuaPath = "D:/daobiao/lua"



def make_one_sheet(xls_file_path, sheet_name):
    objSheetData = SheetData(xls_file_path, sheet_name)  
    objSheetData.makeJsonFile(g_JsonPath + "/" + sheet_name + ".json", sheet_name)
    objSheetData.makeTSFile(g_TsPath + "/" + sheet_name + ".d.ts", sheet_name)
    
    objSheetData.makeLuaFile(g_LuaPath + "/" + sheet_name + ".lua", sheet_name, False)



def make_all_xlsx(xls_floder_path):
    for root, dirs, files in os.walk(xls_floder_path):
        for file in files:
            if file.endswith('.xlsx'):
                xls_file_path = os.path.join(root, file)
                sheets_names = xls_tool.getSheetName(xls_file_path)
                for sheet_name in sheets_names:
                    make_one_sheet(xls_file_path, sheet_name)
    


def delete_all_files_in_folder(folder_path):
    # 检查文件夹是否存在
    if not os.path.exists(folder_path):
        print(f"The folder {folder_path} does not exist.")
        return
    
    # 删除文件夹内的所有文件
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f"Failed to delete {file_path}. Reason: {e}")


def main():
    
    delete_all_files_in_folder(g_JsonPath)
    delete_all_files_in_folder(g_TsPath)
    delete_all_files_in_folder(g_PackPath)
    delete_all_files_in_folder(g_LuaPath)
    
    
    
    make_all_xlsx(g_XlsxFloderPath)
    compress_json_files(g_JsonPath, g_PackPath+"/daobiao.bin")
    print("完成！！！！")
    
# main()



