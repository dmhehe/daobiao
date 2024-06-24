
from simui.simdict import SimDict
from simui.simlist import SimList
import simui.simobj as simobj

import tkinter as tk
from tkinter import ttk

from tkinterdnd2 import TkinterDnD,DND_FILES
import re

from PIL import Image, ImageTk




g_show_win_list = []

g_root_TK = TkinterDnD.Tk()

g_show_obj_list = []

import simui.simwin as simwin

def add_show_win(win_obj):
    g_show_win_list.append(win_obj)


def has_show_win():
    if g_show_win_list:
        return True
    return False



def remove_show_win(win_obj):
    if win_obj in g_show_win_list:
        g_show_win_list.remove(win_obj)

def check_allUI_update_value(obj, name, value):
    for win_obj in g_show_win_list:
        if win_obj.is_show_obj(obj):
            value = win_obj.update_value(obj, name, value)
    return value

def get_root_TK():
    return g_root_TK


def show_ui(title=None):
    global g_root_TK
    
    win_obj = simwin.SimWin(g_root_TK)
    win_obj.set_show_obj_list(g_show_obj_list)
    win_obj.show(title)


def add_show_obj(obj):
    if obj not in g_show_obj_list:
        print("add_show_obj !!!!", obj)
        g_show_obj_list.append(obj)


















