import xls_tool



class SheetData:
    def __init__(self, xls_file_path, sheet_name) -> None:
        self.m_newline_data = {}

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

        self.m_type_list = []#类型名字
        for i in range(self.m_colum):
            attr_dict = xls_tool.getAttrDict(xls_data[4][i])
            self.m_type_list.append(attr_dict)


        self.m_desc_list = []#类型名字
        for i in range(self.m_colum):
            attr_dict = xls_tool.getAttrDict(xls_data[5][i])
            self.m_desc_list.append(attr_dict)

    def addNewLineData(self, data):
        self.m_newline_data[data] = True

    def isNewLineData(self, data):
        return data in self.m_newline_data

    def getAttrName(self, i, j, isClient=True):
        attrName = self.m_attr_name_list[i]["main"]
        if isClient:
            useName = self.m_client_attr_list[i]["main"]
        else:
            useName = self.m_server_attr_list[i]["main"]

        if useName not in ["1", "0"]:
            return useName
        
        return attrName
    

    def getRawValue(self, i, j, isClient=True):
        value_str = self.m_xls_data[i][j]
        return value_str

    def getValue(self, i, j, isClient=True):
        typeName = self.m_type_list[i]["main"]
        value_str = self.m_xls_data[i][j]
        if typeName == "num":
            if value_str == "":
                return 0
            
            return float(value_str)
        
        elif typeName == "str":
            return self.m_xls_data[i][j]

        elif typeName == "strlist":
            if value_str == "":
                return []
            
            return xls_tool.getStringList(value_str)
        
        elif typeName == "numlist":
            if value_str == "":
                return []
            return xls_tool.getNumList(value_str) 
        elif typeName in ("raw", ""):
            if value_str == "":
                return None
            
            try:
                a = float(value_str)
                return a
            except:
                return value_str
    
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
        if mainType == "1":
            return self.getDataByType1(True)
        elif mainType == "2":
            return self.getDataByType2(True)
        elif mainType == "3":
            return self.getDataByType3(True)
        elif mainType == "4":
            return self.getDataByType4(True)
        elif mainType == "5":
            return self.getDataByType5(True)
        elif mainType == "6":
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
        for i in range(6, self.m_row):
            id_val = self.getValue(0, j, isClient)
            id_raw = self.getRawValue(0, j, isClient)

            if id_raw != "":
                curDict = {}
                ans_data[id_val] = curDict
                self.addNewLineData(curDict)
                useKeyDict = {}

            line_data = {}

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
        for i in range(6, self.m_row):
            id_val = self.getValue(0, j, isClient)
            id_raw = self.getRawValue(0, j, isClient)
            id_name = self.getAttrName(0, j, isClient)

            if id_raw != "":
                curList = []
                ans_data[id_val] = curList
                self.addNewLineData(curList)

            line_data = {id_name:id_val}

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
        for i in range(6, self.m_row):
            id_val = self.getValue(0, j, isClient)
            id_raw = self.getRawValue(0, j, isClient)
            id_name = self.getAttrName(0, j, isClient)
            id_val2 = self.getValue(1, j, isClient)
            id_raw2 = self.getRawValue(1, j, isClient)
            id_name2 = self.getAttrName(1, j, isClient)

            if id_raw != "":
                parentDict = {}
                ans_data[id_val] = parentDict
                self.addNewLineData(parentDict)

            if id_raw2 != "":
                curDict2 = {}
                parentDict[id_val2] = curDict2
                self.addNewLineData(curDict2)
                useKeyDict = {}

            line_data = {id_name:id_val, id_name2:id_val2}

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
        for i in range(6, self.m_row):
            id_val = self.getValue(0, j, isClient)
            id_raw = self.getRawValue(0, j, isClient)
            id_name = self.getAttrName(0, j, isClient)
            id_val2 = self.getValue(1, j, isClient)
            id_raw2 = self.getRawValue(1, j, isClient)
            id_name2 = self.getAttrName(1, j, isClient)

            if id_raw != "":
                parentDict = {}
                ans_data[id_val] = parentDict
                self.addNewLineData(parentDict)

            if id_raw2 != "":
                curList2 = []
                parentDict[id_val2] = curList2
                self.addNewLineData(curList2)

            line_data = {id_name:id_val, id_name2:id_val2}

            for j in range(2, self.m_colum):
                if not self.isUseColum(j, isClient):
                    continue

                value = self.getValue(i, j, isClient)
                valueName = self.getAttrName(i, j, isClient)
                line_data[valueName] = value

            curList2.append(line_data)
        return ans_data
    
    def makeJsonFile(self, file_path, isClient=True):
        txt = ""
        if isClient:
            txt = self.getClientData()
        xls_tool.writeJson(file_path, txt)
    
def main():
    xls_file_path = "rrr.xls"
    sheets_names = xls_tool.getSheetName(xls_file_path)
    print("11111111111111", sheets_names)

    for sheet_name in sheets_names:
        print('222222222222222222222222', xls_tool.read_sheet_data(xls_file_path, sheet_name))
        make_one_sheet(xls_file_path, sheet_name)


def make_one_sheet(xls_file_path, sheet_name):
    objSheetData = SheetData(xls_file_path, sheet_name)  
    objSheetData.makeJsonFile()

    
    
        

main()