
import simui.simobj
import simui.simmain

from simui.simobj import sim_func, sim_class, sim_meth

from simui.simmain import show_ui, add_show_obj

from simui.uiobj import init_sim_global


from simui.simtimer import add_timer, del_timer

from simui.simsubwin import show_sub_win, get_show_func_obj

vesion = "0.1.3"


g_dlg_num = 0

from tkinter import messagebox

def show_popup(text):
    root = simui.simmain.get_root_TK()
    global g_dlg_num
    g_dlg_num += 1
    messagebox.showinfo("message", text, master=root)
    g_dlg_num -= 1


def show_yesno(text="确认吗？", yes_text = "确认"):
    root = simui.simmain.get_root_TK()
    global g_dlg_num
    g_dlg_num += 1
    
    response = messagebox.askyesno(yes_text, text, master=root)
    g_dlg_num -= 1

    return response


is_root_withdraw = False
def get_is_root_withdraw():
    return is_root_withdraw

def withdraw():
    global is_root_withdraw
    is_root_withdraw = True
    root = simui.simmain.get_root_TK()
    root.withdraw()
    
def mainloop():
    root = simui.simmain.get_root_TK()
    
    if is_root_withdraw and not simui.simmain.has_show_win():
        return
    
    root.mainloop()
    
    
def destroy():
    import simtimer
    def func1():
        if simui.simmain.has_show_win() or g_dlg_num > 0:
            simtimer.add_timer("simui.destroy",  100, func1)
            return
        root = simui.simmain.get_root_TK()
        root.destroy()
        
    simtimer.add_timer("simui.destroy",  100, func1)