import keyword
import re
import tkinter as tk
from tkinter import ttk

from tkinterdnd2 import TkinterDnD,DND_FILES

from simui.simobj import sim_func
from simui.simobj import sim_class
from simui.simobj import sim_meth
from simui.simmain import add_show_obj, get_root_TK



g_pattern = re.compile(r'^[a-zA-Z_]\w*$')




def is_valid_variable_name(name):
    # 检查是否是保留字
    if keyword.iskeyword(name):
        return False
    
    # 检查是否符合标识符规则
    
    return bool(g_pattern.match(name))


g_code_text = """
@sim_class("",  attr_list)
class key_name:
    pass
    
result = key_name()
"""



def init_sim_global(key_name, attr_list):

    if key_name.strip() == "" or not is_valid_variable_name(key_name):
        raise Exception("key_name 有问题" + key_name)
        return 
    
    
    globals_dict = globals()
    locals_dict = locals()
    code = g_code_text.replace("key_name", key_name)

    # 在指定的全局和局部作用域中执行代码
    exec(code, globals_dict, locals_dict)

    # 从指定的作用域中获取结果
    result = locals_dict.get('result')
    result.bind_json("__save/"+key_name + ".json")
    add_show_obj(result)
    
    return result
