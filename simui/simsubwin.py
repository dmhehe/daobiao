import keyword
import re
import tkinter as tk
from tkinter import ttk

from tkinterdnd2 import TkinterDnD,DND_FILES

from simui.simwin import SimWin
import simui.simmain

from simui.simobj import sim_func
from simui.simobj import sim_class
from simui.simobj import sim_meth


class SimSubWin(SimWin):
    def __init__(self):
        prarent_root = simui.simmain.get_root_TK()
        root = tk.Toplevel(prarent_root)
        super().__init__(root)
        self.m_close_cb_list = []
            
    def show(self, title=None):
        if not title:
           title = "UIWin test"

        # 绑定窗口关闭事件
        self.m_root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        self.initTK(self.m_root, title)
        self.refesh_ui()
        simui.simmain.add_show_win(self)

        
        
    def add_close_callback(self, callback):
        if callback not in self.m_close_cb_list:
            self.m_close_cb_list.append(callback)
            
    def on_close(self):
        simui.simmain.remove_show_win(self)
        self.m_root.destroy()
        
        if simui.get_is_root_withdraw():
            simui.destroy()
            
        try:
            for cb in self.m_close_cb_list:
                if cb:
                    cb()
        except Exception as e:
            print("close cb Error!!!", e)
            

    
    def close(self):
        self.on_close()
    

def show_sub_win(obj_list, title="", close_cb=None, func_list=[]):
    if type(obj_list) != list:
        obj_list = [obj_list]
    else:
        obj_list = list(obj_list)
        
    subwin = new_sub_win(func_list)
    win_obj = subwin()
    obj_list.append(win_obj)
    win_obj.set_show_obj_list(obj_list)
    if close_cb:
        win_obj.add_close_callback(close_cb)
    win_obj.show(title)
    return win_obj




g_win_code_text = """
class key_name(SimSubWin):
"""
g_one_func_code_text = """
    @sim_meth(func_name, func_args)
    def FFFF(self):
        if func_cb:
            func_cb()
"""

g_win_Index = 0


def new_sub_win(func_list):
    global g_win_Index
    g_win_Index += 1
    key_name = "TempSubWin" + str(g_win_Index)
    globals_dict = globals()
    locals_dict = locals()
    code_str = g_win_code_text.replace("key_name", key_name)

    if func_list:
        if type(func_list[0]) != list:
            func_list = [func_list]
    
    for i, funData in enumerate(func_list):
        func, funcName, funcArgs = (funData + [None, None])[:3]
        if funcName == None:
            funcName = func.__name__
        if funcArgs == None:
            funcArgs = {}
        globals_dict["func_name"+str(i)] = funcName
        globals_dict["func_cb"+str(i)] = func
        globals_dict["func_args"+str(i)] = funcArgs
        
        func_code_str = g_one_func_code_text.replace("func_name", "func_name"+str(i))
        func_code_str = func_code_str.replace("func_cb", "func_cb"+str(i))
        func_code_str = func_code_str.replace("func_args", "func_args"+str(i))
        func_code_str = func_code_str.replace("FFFF", "FFFF"+str(i))
        
        code_str = code_str + func_code_str
    else:
        code_str = code_str + "    pass\n"
    print("code_str", code_str)
    
    # 在指定的全局和局部作用域中执行代码
    exec(code_str, globals_dict, locals_dict)


    cls = locals_dict.get(key_name)
     
    return cls



g_show_func_code_text = """
class key_name:
    @sim_meth(func_name, func_args)
    def FFFF(self):
        if func_cb:
            func_cb()
"""

g_show_func_index = 0

def get_show_func_obj(func, func_name=None, func_args=None):
    global g_show_func_index
    g_show_func_index += 1
    key_name = "TempFuncObj" + str(g_win_Index)
    globals_dict = globals()
    locals_dict = locals()
    code_str = g_show_func_code_text.replace("key_name", key_name)
    
    
    if func_name == None:
        func_name = func.__name__
    if func_args == None:
        func_args = {}

    globals_dict["func_name"] = func_name
    globals_dict["func_cb"] = func
    globals_dict["func_args"] = func_args
    
    # 在指定的全局和局部作用域中执行代码
    exec(code_str, globals_dict, locals_dict)


    cls = locals_dict.get(key_name)
     
    return cls()