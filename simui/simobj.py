import simui.simmain as simmain
import os
import json

import simui.simtimer as simtimer
import weakref

g_add_index = 0


def get_add_index():
    global g_add_index
    g_add_index += 1
    return g_add_index





g_class_name_to_class = {
    
}

g_class_Dict = {
    
}


g_dec_Dict = {
    
}

g_class_attr_Dict = {
    
}

g_meth_Dict = {

}


g_func_Dict = {

}

def get_class_by_name(name):
    return g_class_name_to_class[name]



def get_meth_list_by_cls(cls):
    list1 = []
    keylist = list(dir(cls))
    for keyName in keylist:
        funcTemp = getattr(cls, keyName, None)
        if (callable(funcTemp) == False):
            continue
        data = g_meth_Dict.get(funcTemp, None)
        if data:
            list1.append(data)
    return list1

def get_all_func_values():
    return list(g_func_Dict.values())



def get_dec_data_list(cls):
    data_list = g_dec_Dict.get(cls, None)
    if data_list == None:
        return []
    return data_list


def get_dec_class_attr_list(cls):
    args = g_class_Dict.get(cls, None)
    if args == None:
        return []
    return args["attr_list"]


def get_class_attr_to_data(cls, attr_name):
    return g_class_attr_Dict.get(cls, {}).get(attr_name, None)
    
def add_class(cls, args):
    g_class_Dict[cls] = args
    
    g_class_name_to_class[cls.__name__] = cls

    list1 = g_dec_Dict.get(cls, [])
    g_dec_Dict[cls] = list1

    attr_dict = g_class_attr_Dict.get(cls, {})
    g_class_attr_Dict[cls] = attr_dict
    
    for i, item_data in enumerate(args["attr_list"]):
        attr_name, show_name, att_cls, args_dict = (item_data + [None, None, None, None])[:4]
        if type(attr_name) != str:
            raise "attr_name should be str !!"
        
        if show_name != None and type(show_name) != str:
            raise "show_name should be str or None !!"
        
        if att_cls != None and type(att_cls) != str:
            raise "att_cls should be str or None !!"

        if args_dict != None and type(args_dict) != dict:
            raise "args_dict should be dict or None !!"
        
        if args_dict == None:
            args_dict = {}
        iPrior = args_dict.get("iPrior", 0)
        item_data = {
            "attr_name": attr_name,
            "show_name": show_name,
            "att_cls": att_cls,
            "args_dict": args_dict,
            "add_index":get_add_index(),
            "type": "class_attr",
            "iPrior":iPrior,
        }

        list1.append(item_data)
        attr_dict[attr_name] = item_data
        

def add_meth(func, args):
    iPrior = args.get("iPrior", 0)
    item_data = {
        "func": func,
        "args_dict": args,
        "add_index":get_add_index(),
        "type": "meth",
        "iPrior":iPrior,
    }
    g_meth_Dict[func] = item_data



def add_func(func, args):
    iPrior = args.get("iPrior", 0)
    name = "!!main_func!!"
    list1 = g_dec_Dict.get(name, [])
    g_dec_Dict[name] = list1
    item_data = {
        "func": func,
        "args_dict": args,
        "add_index":get_add_index(),
        "type": "func",
        "iPrior":iPrior,
    }
    g_func_Dict[func] = item_data







def sim_func(sName=None, args=None):
    if sName != None and type(sName) != str:
        raise "sName should be str or None !!"
    
    if args != None and type(args) != dict:
        raise "args should be dict or None !!"
    
    
    if not args:
        args = {}


    args["sName"] = sName
    
    
    # 定义一个函数装饰器
    def dec_func(func):
        print("add sim_func!!!!", func.__name__)
        add_func(func, args)
        
        return func
    return dec_func


def sim_meth(sName=None, args=None):
    if sName != None and type(sName) != str:
        raise "sName should be str or None !!"
    
    if args != None and type(args) != dict:
        raise "args should be dict or None !!"
    
    if not args:
        args = {}


    args["sName"] = sName
    
    
    # 定义一个函数装饰器
    def dec_meth(func):
        print("add sim_meth!!!!", func.__name__)
        add_meth(func, args)
        
        return func
    return dec_meth



def __get_save_json_data(self):
    attr_list = get_dec_data_list(self.__class__)
    ans_data = {}
    for item_data in attr_list:
        attr_name = item_data["attr_name"]
        att_cls = item_data["att_cls"]
        if attr_name in self.__dict__:
            if g_class_name_to_class.get(att_cls, None):
                ans_data[attr_name] = self.__dict__[attr_name].__get_save_json_data()
            else:
                ans_data[attr_name] = self.__dict__[attr_name]
    return ans_data

def __load_json(self, ans_data):
    attr_list = get_dec_data_list(self.__class__)
    for item_data in attr_list:
        attr_name = item_data["attr_name"]
        att_cls = item_data["att_cls"]
        if attr_name in ans_data:
            if g_class_name_to_class.get(att_cls, None):
                new_obj = g_class_name_to_class[att_cls]()
                new_obj.__load_json(ans_data[attr_name])
                new_obj._set_parent(self)
                setattr(self, attr_name, new_obj)
            else:
                setattr(self, attr_name, ans_data[attr_name])


def bind_json(self, file_name, key_name=None, save_gap=1000):

    folder_name, _ = os.path.split(file_name)
    
    folder_name = folder_name.strip()
    if folder_name:
        os.makedirs(folder_name, exist_ok=True)
    
    self.__bind_save_gap = save_gap
    self.__bind_file_name = file_name
    self.__bind_key_name = key_name
    file_path = file_name  # 替换为你的文件路径

    
    data = get_file_data(file_path)
    if data:
        if self.__bind_key_name != None:
            data = data.get(self.__bind_key_name, None)
        # print(data)  # 这里可以对读取的 JSON 数据进行处理
        if data:
            self.__load_json(data)
    else:
        print("File does not exist or empty.")
        
def __save_json_now(self):
    file_path = self.__bind_file_name  # 替换为你的文件路径
    
    if file_path is None:
        return
    
    data = self.__get_save_json_data()
    new_data = data
    if self.__bind_key_name != None:
        old_data = get_file_data(file_path)
        if old_data:
            old_data[self.__bind_key_name] = data
            new_data = old_data
        else:
            new_data = {self.__bind_key_name: data}


    with open(file_path, 'w', encoding='utf-8') as file:
        json.dump(new_data, file, ensure_ascii=False, indent=4)

def get_file_data(file_path):
    if os.path.exists(file_path):
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
            return data

def _check_save_json(self, change_attr_name=None):
    if getattr(self, "_parent_ref", None) != None:
        self._parent_ref()._check_save_json(change_attr_name)
        return
    
    if getattr(self, "__bind_save_gap", None) != None:
        def cb():
            self.__save_json_now()
        
        
        key_name = str(id(self)) + "_save"
        simtimer.add_timer(key_name, self.__bind_save_gap, cb)
        


def _set_parent(self, parent):
    self._parent_ref = weakref.ref(parent)


def sim_class(sName=None, attr_list=None, args=None):
    if sName != None and type(sName) != str:
        raise "sName should be str or None !!"
    
    if args != None and type(args) != dict:
        raise "args should be dict or None !!"
    
    if attr_list != None and type(attr_list) != list:
        raise "attr_list should be list or None !!"

    if not args:
        args = {}
    args["sName"] = sName

    if not attr_list:
        attr_list = []
    args["attr_list"] = attr_list

    # 定义一个类装饰器
    def dec_class(cls):
        add_class(cls, args)
        print("add_class !!!!", cls)

        # # 定义一个新的方法
        # 将新方法添加到类中
        cls.__get_save_json_data = __get_save_json_data
        cls._check_save_json = _check_save_json
        cls.__load_json = __load_json
        cls.__save_json_now = __save_json_now
        cls._set_parent = _set_parent
        cls.bind_json = bind_json
        
        def custom_setattr(self, name, value):
            # print(f"Setting attribute {name} with value {value}")
            value = simmain.check_allUI_update_value(self, name, value)
            super(cls, self).__setattr__(name, value)

        setattr(cls, '__setattr__', custom_setattr)

        return cls
    
    return dec_class




