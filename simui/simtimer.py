import json

import weakref

import simui.simmain as simmain

timer_ids = {}  # 用于存储定时器 ID 的字典

def add_timer(key_name, iTime, cb):
    root_TK = simmain.get_root_TK()
    # print(f"Timer callback for {key_name}!{iTime}{cb}")
    del_timer(key_name)
    # 设置下一个定时器回调，间隔为 1000ms (1秒)
    timer_ids[key_name] = root_TK.after(iTime, cb)

def del_timer(key_name):
    root_TK = simmain.get_root_TK()
    if key_name in timer_ids:
        root_TK.after_cancel(timer_ids[key_name])
        # print(f"Timer for {key_name} cancelled!")
        del timer_ids[key_name]


