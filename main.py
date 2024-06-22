from typing import Any
import xls_tool

import os
import zipfile


class SheetData:
    def __init__(self, xls_file_path, sheet_name) -> None:
        self.m_newline_data = []
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

    def addNewLineData(self, data):
        self.m_newline_data.append(data)

    def isNewLineData(self, data):
        return data in self.m_newline_data

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
            
            return xls_tool.getStringList(value_str)
        
        elif typeName == "num[]":
            if value_str == "":
                return []
            return xls_tool.getNumList(value_str) 
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


    def getClientData(self):
        mainType = self.m_xls_attr_dict["main"]
        if mainType == "1": #第一种类型  {key:obj}
            return self.getDataByType1(True)
        elif mainType == "2": #[obj]
            return self.getDataByType2(True)
        elif mainType == "3":#{key:{key2:obj}}
            return self.getDataByType3(True)
        elif mainType == "4":#{key:[obj]}
            return self.getDataByType4(True)
        elif mainType == "5":#{key:{key2:{key3:obj}}}
            return self.getDataByType5(True)
        elif mainType == "6":#{key:{key2:[obj]}}
            return self.getDataByType6(True)

    
    def getDataByType1(self, isClient=True):
        ans_data = {}
        self.addNewLineData(ans_data)
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

            ans_data[key] = line_data
        return ans_data
    
    def getDataByType2(self, isClient=True):
        ans_data = []
        self.addNewLineData(ans_data)
        for i in range(6, self.m_row):
            line_data = {}
            for j in range(0, self.m_colum):

                if not self.isUseColum(j, isClient):
                    continue

                value = self.getValue(i, j, isClient)
                valueName = self.getAttrName(i, j, isClient)
                line_data[valueName] = value
                
            ans_data.append(line_data)
        return ans_data
    

    def getDataByType3(self, isClient=True):
        ans_data = {}
        self.addNewLineData(ans_data)
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
                self.addNewLineData(curDict)
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

            curDict[key] = line_data
        return ans_data
    
    def getDataByType4(self, isClient=True):
        ans_data = {}
        self.addNewLineData(ans_data)

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
                self.addNewLineData(curList)

            line_data = {id_name:curVal}

            for j in range(1, self.m_colum):
                if not self.isUseColum(j, isClient):
                    continue

                value = self.getValue(i, j, isClient)
                valueName = self.getAttrName(i, j, isClient)
                line_data[valueName] = value

            curList.append(line_data)
        return ans_data

    def getDataByType5(self, isClient=True):
        ans_data = {}
        self.addNewLineData(ans_data)
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
                self.addNewLineData(parentDict)

            if id_raw2 != "":
                curDict2 = {}
                parentDict[id_val2] = curDict2
                curVal2 = id_val2
                self.addNewLineData(curDict2)
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

            curDict2[key] = line_data
        return ans_data
    

    def getDataByType6(self, isClient=True):
        ans_data = {}
        self.addNewLineData(ans_data)

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
                self.addNewLineData(parentDict)

            if id_raw2 != "":
                curList2 = []
                parentDict[id_val2] = curList2
                curVal2 = id_val2
                self.addNewLineData(curList2)

            line_data = {id_name:curVal1, id_name2:curVal2}

            for j in range(2, self.m_colum):
                if not self.isUseColum(j, isClient):
                    continue

                value = self.getValue(i, j, isClient)
                valueName = self.getAttrName(i, j, isClient)
                line_data[valueName] = value

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
        



    def makeJsonFile(self, file_path, isClient=True):
        txt = ""
        if isClient:
            txt = self.getClientData()
        xls_tool.writeJson(file_path, txt)
        
    def makeTSFile(self, file_path, isClient=True):
        txt = ""
        if isClient:
            txt = self.getTsStatementData()
        xls_tool.writeFile(file_path, txt)
    





def compress_json_files(folder_path, zip_file_path):
    
    directory = os.path.dirname(zip_file_path)
    if not os.path.exists(directory):
        os.makedirs(directory)
    
    with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith('.json'):
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, os.path.relpath(file_path, folder_path))


g_JsonPath = "D:/daobiao/json"
g_TsPath = "D:/daobiao/ts"
g_PackPath = "D:/daobiao/pack"

g_XlsxFloderPath = "D:/daobiao/xlsx"

def make_one_sheet(xls_file_path, sheet_name):
    objSheetData = SheetData(xls_file_path, sheet_name)  
    objSheetData.makeJsonFile(g_JsonPath + "/" + sheet_name + ".json")
    objSheetData.makeTSFile(g_TsPath + "/" + sheet_name + ".d.ts")



def make_all_xlsx(xls_floder_path):
    for root, dirs, files in os.walk(xls_floder_path):
        for file in files:
            if file.endswith('.xlsx'):
                xls_file_path = os.path.join(root, file)
                sheets_names = xls_tool.getSheetName(xls_file_path)
                for sheet_name in sheets_names:
                    make_one_sheet(xls_file_path, sheet_name)
    

def main():
    make_all_xlsx(g_XlsxFloderPath)
    compress_json_files(g_JsonPath, g_PackPath+"/daobiao.zip")
    
main()