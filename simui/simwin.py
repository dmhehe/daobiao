import os
from tkinter import font
from simui.simdict import SimDict
from simui.simlist import SimList
import simui.simobj as simobj

import tkinter as tk
from tkinter import ttk

from tkinterdnd2 import TkinterDnD,DND_FILES
import re

from PIL import Image, ImageTk
from tkinter import PhotoImage

import simui.simmain
from simui.simexec import exec_and_return
from PIL import Image, ImageTk
import sys

g_win_width = 600
g_default_one_item_h = 30
g_default_gap = 10

g_purple_image = Image.new('RGB', (100, 100), color='purple')

    # 将图像转换为Tkinter支持的格式
g_purple_tk_image = ImageTk.PhotoImage(g_purple_image)


g_float_range_pattern = r'float\(([-+]?[0-9]*\.?[0-9]+), ([-+]?[0-9]*\.?[0-9]+)\)'


g_int_range_pattern = re.compile(r'int\((-?\d+), (-?\d+)\)')





# 构造资源文件的完整路径
checked_image_path = "resources/checked.png"
unchecked_image_path = "resources/unchecked.png"

g_checked_image = None
g_unchecked_image = None



if not os.path.exists(checked_image_path):
    folder_name, file_name = os.path.split(sys.executable)
    checked_image_path = os.path.join(folder_name,  "Lib", "simui", "resources", "checked.png")
    unchecked_image_path = os.path.join(folder_name,  "Lib", "simui", "resources", "unchecked.png")


# 加载图片
if os.path.exists(checked_image_path):
    g_checked_image = Image.open(checked_image_path) 
    g_unchecked_image = Image.open(unchecked_image_path)


def resize_image(image, new_width, new_height):
    image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
    return ImageTk.PhotoImage(image)



class SimWin():
    def __init__(self, root):
        self.m_root = root
        self.m_inner_frame = None
        self.m_int_validator = None #整形验证功能
        self.m_float_validator = None
        self.m_show_obj_list = []
        
    def initTK(self, root, title=None):
        if not title:
            title = "SimUI Demo"

        # root = TkinterDnD.Tk()
        root.title(title)
        # 创建一个滚动条
        scrollbar = tk.Scrollbar(root)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        
        # 创建一个 Canvas
        canvas = tk.Canvas(root, width=g_win_width, height=800)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 鼠标滚轮事件绑定
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))

        # 创建一个内部框架
                
        inner_frame = tk.Frame(canvas)
        self.m_inner_frame = inner_frame
        on_configure = self.get_on_configure_func(canvas)
        inner_frame.bind("<Configure>", on_configure)  # 确保内容变化时能调整滚动范围
        
        
        # 关联滚动条与 Canvas
        scrollbar.config(command=canvas.yview)
        canvas.config(yscrollcommand=scrollbar.set, scrollregion=canvas.bbox("all"))

        self.m_int_validator = root.register(self.on_int_validate)
        
        self.m_float_validator = root.register(self.on_float_validate_input)
        
        canvas.create_window((0, 0), window=self.m_inner_frame, anchor=tk.NW)
        
        
        
    def get_on_configure_func(self, canvas):
        def on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        return on_configure



    def on_int_validate(self, P):
        if P.isdigit() or P == "":
            return True
        else:
            return False


    def on_float_validate_input(self, P):
        if P == "":
            return True

        try:
            float(P)
            return True
        except ValueError:
            return False



    def make_float_validate_input(self, min_value, max_value, node_dict):
        def validate_input(P):
            entry = node_dict["entry"]
            if P == "":
                return True

            try:
                float_value = float(P)
                
                
                if max_value<float_value:
                    entry.config(validate="none")
                    entry.delete(0, tk.END)
                    entry.insert(0, str(max_value))
                    entry.config(validate="key")
                    return False
                elif min_value>float_value:
                    entry.config(validate="none")
                    entry.delete(0, tk.END)
                    entry.insert(0, str(min_value)) 
                    entry.config(validate="key")
                    return False
                
                
                return min_value <= float_value <= max_value
            except ValueError:
                return False
        return validate_input



    def make_int_validate_input(self, min_value, max_value, data_dict):
        def validate_input(P):
            scale_node = data_dict["scale_node"]
            entry = data_dict["entry_int_range"]
            if P == "":
                return True

            try:
                int_value = int(P)
                
                
                if max_value<int_value:
                    entry.config(validate="none")
                    entry.delete(0, tk.END)
                    entry.insert(0, str(max_value))
                    entry.config(validate="key")
                    scale_node.set(max_value)
                    return False
                elif min_value>int_value:
                    entry.config(validate="none")
                    entry.delete(0, tk.END)
                    entry.insert(0, str(min_value)) 
                    entry.config(validate="key")
                    scale_node.set(min_value)
                    return False
                
                
                return min_value <= int_value <= max_value
            except ValueError:
                return False
        return validate_input


    def make_on_int_scale_changed(self, data_dict, obj, attr_name):
        def on_scale_changed(value):
            integer_value = int(float(value))
            # print(f"Value selected: {integer_value}")
            entry_int_range = data_dict["entry_int_range"]
            
            entry_int_range.config(validate="none")
            entry_int_range.delete(0, tk.END)  # 删除所有文本
            entry_int_range.insert(0, str(integer_value))  # 插入新值
            entry_int_range.config(validate="key")
            
            self.raw_set_attr(obj, attr_name, integer_value)
            obj._check_save_json(attr_name)
            
        return on_scale_changed



    def make_on_float_scale_changed(self, data_dict, obj, attr_name):
        def on_scale_changed(value):
            float_value = float(value)
            # print(f"Value selected: {float_value}")
            entry_float_range = data_dict["entry_float_range"]
            
            entry_float_range.config(validate="none")
            entry_float_range.delete(0, tk.END)  # 删除所有文本
            entry_float_range.insert(0, str(float_value))  # 插入新值
            entry_float_range.config(validate="key")
            
            
            self.raw_set_attr(obj, attr_name, float_value)
            obj._check_save_json(attr_name)
            
        return on_scale_changed



    def get_percent_width(self, percent):
        return int(g_win_width * percent)

    def get_left_with(self):
        return self.get_percent_width(0.3)

    def get_right_with(self):
        return self.get_percent_width(0.7) - 20


    def get_percent_right_with(self, percent):
        right_width = self.get_right_with()
        return int(right_width * percent)


    def get_percent_left_with(self, percent):
        left_width = self.get_left_with()
        return int(left_width * percent)


        
    def make_TK_node(self, obj, iCurY, add_all_func):
        
        
        cls = obj.__class__
        
        meth_list = simobj.get_meth_list_by_cls(cls) #有的方法
        
        attr_list = simobj.get_dec_data_list(cls) # 有的属性
        
        func_list = [] # 有的方法
        if add_all_func:
            func_list = simobj.get_all_func_values()
        
        add_all_func = False
        
        
        show_list = attr_list + meth_list + func_list
        
        # print("11111111111111", attr_list)
        # print("222222222222", meth_list)
        # print("222222222222", func_list)
        show_list = sorted(show_list, key=lambda x: (-1*x["iPrior"], x["add_index"]))
        
        show_dict = {}
        obj.__dict__["__showing"] = True
        obj.__dict__["__show_dict"] = show_dict
        
        root = self.m_inner_frame
        
        for one_data in show_list:
            stype = one_data["type"]
            args_dict = one_data["args_dict"]
            
            b_show = args_dict.get("show", True)
            if not b_show:
                continue
            
            if stype == "class_attr":
                iCurY, next_show_obj = self.ui_add_one_attr(obj, cls, root, iCurY, one_data, show_dict)
                if next_show_obj:
                    iCurY = self.make_TK_node(next_show_obj, iCurY, add_all_func)
                
            elif stype == "func":
                iCurY = self.ui_add_one_func(cls, root, iCurY, one_data, show_dict)
                
            elif stype == "meth":
                iCurY = self.ui_add_one_meth(obj, cls, root, iCurY, one_data, show_dict)
        
        return iCurY
        
    def ui_add_one_attr(self, obj, cls, root, iCurY, one_data, show_dict, start_x = 0):
        next_show_obj = None
        
        left_width = self.get_left_with()
        right_width = self.get_right_with()
        
        left_start = start_x
        right_start = start_x + left_width + 20 
        
        attr_name, show_name, att_cls, args_dict = one_data["attr_name"], one_data["show_name"], one_data["att_cls"], one_data["args_dict"]
        if not args_dict:
            args_dict = {}
            one_data["args_dict"] = args_dict
            
        if not show_name:
            show_name = attr_name
        
        if not att_cls:
            att_cls = ""
        
        iGap = g_default_gap
        iHeight = g_default_one_item_h

        
        obj_type = type(obj)
        if obj_type == SimList:
            value = obj[int(attr_name)]
        elif obj_type == SimDict:
            value = obj[attr_name]
        else:
            value = getattr(obj, attr_name, None)

        if value == None:
            value = args_dict.get("default", None)
        
        
        b_show_name = args_dict.get("show_name", True)
        # 创建不可编辑的标签
        
        if b_show_name:
            label_show_name = tk.Label(root, text=show_name, bg="lightblue")
            
            name_x = left_start + 10
            name_y = iCurY
            name_w = left_width
            name_h = iHeight
            if args_dict.get("name_x", ""):
                name_x = exec_and_return(args_dict["name_x"])

            if args_dict.get("name_y", ""):
                name_y = iCurY + exec_and_return(args_dict["name_y"])
            
            if args_dict.get("name_w", ""):
                name_w = exec_and_return(args_dict["name_w"])
            
            if args_dict.get("name_h", ""):
                name_h = exec_and_return(args_dict["name_h"])
            
            label_show_name.place(x=name_x, y=name_y, width=name_w, height=name_h)
            
        # 创建可编辑的文本框
        entry = None
        checkbox = None
        checkbox_var = None
        combo = None
        text_node = None
        scale_node = None
        entry_int_range = None
        entry_float_range = None
        img_label = None
        
        btn_list_up = None
        btn_list_del = None
        list_num_label = None
        
        if att_cls == "int":
            
            if value == None:
                value = 0
                self.raw_set_attr(obj, attr_name, value)
            
            entry = tk.Entry(root, validate="key", validatecommand=(self.m_int_validator, '%P'), bg="#dddd00")
            
        elif att_cls == "float":
            
            if value == None:
                value = 0.0
                self.raw_set_attr(obj, attr_name, value)
            
            entry = tk.Entry(root, validate="key", validatecommand=(self.m_float_validator, '%P'), bg="#dd3333")
        
        elif att_cls.startswith("float("):
            fa = 0.0
            fb = 100.0
            
            try:
                match = re.match(g_float_range_pattern, att_cls.strip())
                if match:
                    fa = float(match.group(1))
                    fb = float(match.group(2))
                if fa > fb:
                    raise ValueError("fa > fb")
            except:
                print("float() parse error!!!")
            
            
            if value == None:
                value = fa
                self.raw_set_attr(obj, attr_name, value)
                
                
            node_dict = {}
            scale_node = tk.Scale(root, from_=fa, to=fb, orient=tk.HORIZONTAL, command=self.make_on_float_scale_changed(node_dict, obj, attr_name))
            iHeight += 10
            
            
            left_right_width = self.get_percent_right_with(0.3)
            right_right_width = self.get_percent_right_with(0.7)
            right_right_start = right_start + left_right_width
            scale_node.place(x=right_right_start, y=iCurY-5, width=right_right_width, height=iHeight)
            
            
            temp_func = self.make_float_validate_input(fa, fb, node_dict)
            temp_func_validator = root.register(temp_func)
            entry_float_range = tk.Entry(root, validate="key", validatecommand=(temp_func_validator, '%P'), bg="#dd1515")
            entry_float_range.place(x=right_start, y=iCurY, width=left_right_width, height=iHeight)
            
            node_dict["entry_float_range"] = entry_float_range
            node_dict["scale_node"] = scale_node
            
            
            if value:
                scale_node.set(value)
            entry_float_range.config(validate="none")
            entry_float_range.delete(0, tk.END)  # 删除所有文本
            entry_float_range.insert(0, str(value))  # 插入新值
            entry_float_range.config(validate="key")
            
            entry_float_range.bind("<KeyRelease>", self.make_float_range_entry_change_func(node_dict, obj, attr_name))
            
        elif att_cls.startswith("int("):
            ia = 0
            ib = 100
            
            try:
                match = re.match(g_int_range_pattern, att_cls.strip())
                if match:
                    ia = int(match.group(1))
                    ib = int(match.group(2))
                if ia > ib:
                    raise ValueError("ia > ib")
            except:
                print("int() parse error!!!")

            if value == None:
                value = ia
                self.raw_set_attr(obj, attr_name, value)


            left_right_width = self.get_percent_right_with(0.3)
            right_right_width = self.get_percent_right_with(0.7)
            right_right_start = right_start + left_right_width

            data_dict = {}
            scale_node = tk.Scale(root, from_=ia, to=ib, orient=tk.HORIZONTAL, resolution=1, command=self.make_on_int_scale_changed(data_dict, obj, attr_name))
            iHeight += 10
            
            scale_node.place(x=right_right_start, y=iCurY-5, width=right_right_width, height=iHeight)
            
            temp_func = self.make_int_validate_input(ia, ib, data_dict)
            temp_func_validator = root.register(temp_func)
            entry_int_range = tk.Entry(root, validate="key", validatecommand=(temp_func_validator, '%P'), bg="#dd1515")
            
            entry_int_range.place(x=right_start, y=iCurY, width=left_right_width, height=iHeight)
            
            
            data_dict["entry_int_range"] = entry_int_range
            data_dict["scale_node"] = scale_node

            if value:
                scale_node.set(value)
            
            entry_int_range.config(validate="none")
            entry_int_range.delete(0, tk.END)  # 删除所有文本
            entry_int_range.insert(0, str(value))  # 插入新值
            entry_int_range.config(validate="key")
            entry_int_range.bind("<KeyRelease>", self.make_int_range_entry_change_func(data_dict, obj, attr_name))
            
            
        elif att_cls == "bool":
            
            if value == None:
                value = False
                
                
            var = tk.IntVar()
            checkbox_var = var
            on_checkbox_clicked = self.make_on_checkbox_clicked(var, obj, attr_name)
            

            
            checkbox_x = right_start + right_width*0.5
            checkbox_y = iCurY
            checkbox_w = iHeight
            checkbox_h = iHeight
            if args_dict.get("x", ""):
                checkbox_x = exec_and_return(args_dict["x"])

            if args_dict.get("y", ""):
                checkbox_y = iCurY + exec_and_return(args_dict["y"])
            
            if args_dict.get("w", ""):
                checkbox_w = exec_and_return(args_dict["w"])
            
            if args_dict.get("h", ""):
                checkbox_h = exec_and_return(args_dict["h"])

            
            if g_unchecked_image:
                img1 = resize_image(g_unchecked_image, checkbox_w, checkbox_h)
                img2 = resize_image(g_checked_image, checkbox_w, checkbox_h)
                checkbox = tk.Checkbutton(root, text="", variable=var, command=on_checkbox_clicked, image=img1, selectimage=img2, indicatoron=False)
                checkbox.img1 = img1
                checkbox.img2 = img2
            else:
                checkbox = tk.Checkbutton(root, text="", variable=var, command=on_checkbox_clicked)
                
            checkbox.place(x=checkbox_x, y=checkbox_y, width=checkbox_w, height=checkbox_h)
            
            if value == True:
                self.raw_set_attr(obj, attr_name, True)
                var.set(1)
            else:
                self.raw_set_attr(obj, attr_name, False)
                var.set(0)
        elif att_cls == "box":
                        
            # 创建列表
            options = args_dict.get("options", [])
            show_options = args_dict.get("show_options", [])
            if not show_options:
                show_options = options

            if len(options) != len(set(options)):
                raise ValueError("options must be unique")
            if len(show_options) != len(set(show_options)):
                raise ValueError("show_options must be unique")
            
            if len(options) != len(show_options):
                raise ValueError("len(options) != len(show_options)")
            
            show_to_value_dict = {}
            value_to_show_dict = {}
            for i in range(len(show_options)):
                show_to_value_dict[show_options[i]] = options[i]
            for i in range(len(options)):
                value_to_show_dict[options[i]] = show_options[i]
            
            # 创建 Combobox 组件并设置为只读模式
            combo = ttk.Combobox(root, values=show_options, state="readonly")
            combo.show_to_value_dict = show_to_value_dict
            combo.value_to_show_dict = value_to_show_dict
            
            
            # 设置默认选中的值
            if options:
                if value in options:
                    self.raw_set_attr(obj, attr_name, value)
                    combo.set(value_to_show_dict[value])
                else:
                    value = options[0]
                    self.raw_set_attr(obj, attr_name, value)
                    combo.set(value_to_show_dict[value])

            handle_selection = self.make_on_handle_selection(combo, obj, attr_name)
            # 绑定事件处理程序
            combo.bind("<<ComboboxSelected>>", handle_selection)
            
            
            def on_combobox_scroll(event):
                return "break"  # 阻止 Combobox 响应滚轮事件


            # 绑定 Combobox 的滚轮事件处理函数
            combo.bind("<MouseWheel>", on_combobox_scroll)
            
            box_x = right_start
            box_y = iCurY
            box_w = right_width
            box_h = iHeight
            if args_dict.get("x", ""):
                box_x = exec_and_return(args_dict["x"])

            if args_dict.get("y", ""):
                box_y = iCurY + exec_and_return(args_dict["y"])
            
            if args_dict.get("w", ""):
                box_w = exec_and_return(args_dict["w"])
            
            if args_dict.get("h", ""):
                box_h = exec_and_return(args_dict["h"])

            
            combo.place(x=box_x, y=box_y, width=box_w, height=box_h)
            
            
        elif att_cls == "text":
            
            if value == None:
                value = ""
                self.raw_set_attr(obj, attr_name, value)
                
            text_node = tk.Text(root, bg="lightgreen")
            
            line_num = args_dict.get("line", 5)
            iHeight = line_num*iHeight
            
            
            
            
            text_node_x = right_start
            text_node_y = iCurY
            text_node_w = right_width
            text_node_h = iHeight
            if args_dict.get("x", ""):
                text_node_x = exec_and_return(args_dict["x"])

            if args_dict.get("y", ""):
                text_node_y = iCurY + exec_and_return(args_dict["y"])
            
            if args_dict.get("w", ""):
                text_node_w = exec_and_return(args_dict["w"])
            
            if args_dict.get("h", ""):
                text_node_h = exec_and_return(args_dict["h"])

            
            
            
            text_node.place(x=text_node_x, y=text_node_y, width=text_node_w, height=text_node_h)
            text_node.delete("1.0", tk.END)  # 删除所有文本
            
            text_node.insert("1.0", str(value))  # 插入新值
        
            text_node.bind("<KeyRelease>", self.make_text_node_change_func(text_node, obj, attr_name))
            text_node.drop_target_register(DND_FILES)
            text_node.dnd_bind('<<Drop>>', self.make_text_node_on_drop_func(text_node, obj, attr_name))
        
            scrollbar_y = tk.Scrollbar(root, command=text_node.yview)
            scrollbar_y.place(x=g_win_width-15, y=iCurY, height=iHeight)
            text_node.config(yscrollcommand=scrollbar_y.set)
            
        elif att_cls == "image":
            
            
            # 定义要显示的图片大小
            target_width = 100  # 设定目标宽度
            target_height = 100  # 设定目标高度
            iHeight = target_height
            # 调整图像大小
            
            tk_image = g_purple_tk_image #加载失败就显示紫色图片
            try:
                image = Image.open(value)
                resized_image = image.resize((target_width, target_height))
                # 将调整大小后的图片转换为Tkinter支持的格式
                tk_image = ImageTk.PhotoImage(resized_image)
            except:
                pass


            img_x = right_start
            img_y = iCurY
            img_w = target_width
            img_h = iHeight
            if args_dict.get("x", ""):
                img_x = exec_and_return(args_dict["x"])

            if args_dict.get("y", ""):
                img_y = iCurY + exec_and_return(args_dict["y"])
            
            if args_dict.get("w", ""):
                img_w = exec_and_return(args_dict["w"])
            
            if args_dict.get("h", ""):
                img_h = exec_and_return(args_dict["h"])


            # 创建标签以显示调整后的图片
            img_label = tk.Label(root, image=tk_image)
            img_label.image = tk_image
            img_label.place(x=img_x, y=img_y, width=img_w, height=img_h)
            img_label.drop_target_register(DND_FILES)
            img_label.dnd_bind('<<Drop>>', self.make_img_on_drop_func(img_label, obj, attr_name))

        elif att_cls == "SimList":
            label_value = 0
            if value and type(value) == list:
                label_value = len(value)
            
            show_vaule = args_dict.get("show_vaule", True)
            
            
            btn_width = 50
            label_width = 60
            myCurStart = right_start
            
            if show_vaule:
                btn_list_up = tk.Button(root, text="↑", command=self.make_show_vaule_up_btn_func(obj, attr_name, args_dict))
                btn_list_up.place(x=myCurStart, y=iCurY, width=btn_width, height=iHeight)
            else:
                btn_list_down = tk.Button(root, text="↓", command=self.make_show_vaule_down_btn_func(obj, attr_name, args_dict))
                btn_list_down.place(x=myCurStart, y=iCurY, width=btn_width, height=iHeight)
            myCurStart += btn_width + 10

            list_num_label = tk.Label(root, text=label_value, bg="lightblue")
            
            btn_list_del = tk.Button(root, text="-", command=self.make_list_del_btn_func(obj, attr_name, list_num_label))
            btn_list_del.place(x=myCurStart, y=iCurY, width=btn_width, height=iHeight)
            myCurStart += btn_width + 10

                
            list_num_label.place(x=myCurStart, y=iCurY, width=label_width, height=iHeight)
            myCurStart += label_width + 10
            
            btn_list_add = tk.Button(root, text="+", command=self.make_list_add_btn_func(obj, attr_name, list_num_label))
            btn_list_add.place(x=myCurStart, y=iCurY, width=btn_width, height=iHeight)
            myCurStart += btn_width + 10
            
            
        elif att_cls == "SimDict":
            label_value = 0
            if value and type(value) == SimDict:
                label_value = len(value)
            
            show_vaule = args_dict.get("show_vaule", True)
            
            
            btn_width = 50
            label_width = 60
            myCurStart = right_start
            
            if show_vaule:
                btn_dict_up = tk.Button(root, text="↑", command=self.make_show_vaule_up_btn_func(obj, attr_name, args_dict))
                btn_dict_up.place(x=myCurStart, y=iCurY, width=btn_width, height=iHeight)
            else:
                btn_dict_down = tk.Button(root, text="↓", command=self.make_show_vaule_down_btn_func(obj, attr_name, args_dict))
                btn_dict_down.place(x=myCurStart, y=iCurY, width=btn_width, height=iHeight)
            myCurStart += btn_width + 10

            dict_num_label = tk.Label(root, text=label_value, bg="lightblue")
            
            btn_dict_del = tk.Button(root, text="-", command=self.make_dict_del_btn_func(obj, attr_name, dict_num_label))
            btn_dict_del.place(x=myCurStart, y=iCurY, width=btn_width, height=iHeight)
            myCurStart += btn_width + 10

                
            dict_num_label.place(x=myCurStart, y=iCurY, width=label_width, height=iHeight)
            myCurStart += label_width + 10
            
            btn_dict_add = tk.Button(root, text="+", command=self.make_dict_add_btn_func(obj, attr_name, dict_num_label))
            btn_dict_add.place(x=myCurStart, y=iCurY, width=btn_width, height=iHeight)
            myCurStart += btn_width + 10
            
        elif att_cls == "str" or att_cls == "string":
            if value == None:
                value = ""
                self.raw_set_attr(obj, attr_name, value)
                
            entry = tk.Entry(root, bg="lightgreen")
        elif att_cls == "" or att_cls == None:
            entry = tk.Entry(root, bg="lightgreen")
        
        else:
            
            cls = simobj.get_class_by_name(att_cls)
            if cls:
                if value == None:
                    value = cls()
                    value._set_parent(obj)
                    self.raw_set_attr(obj, attr_name, value)
                next_show_obj = value
                    
        
        if entry:
            entry_x = right_start
            entry_y = iCurY
            entry_w = right_width
            entry_h = iHeight
            if args_dict.get("x", ""):
                entry_x = exec_and_return(args_dict["x"])

            if args_dict.get("y", ""):
                entry_y = iCurY + exec_and_return(args_dict["y"])
            
            if args_dict.get("w", ""):
                entry_w = exec_and_return(args_dict["w"])
            
            if args_dict.get("h", ""):
                entry_h = exec_and_return(args_dict["h"])
            
            entry.place(x=entry_x, y=entry_y, width=entry_w, height=entry_h)
            entry.bind("<KeyRelease>", self.make_entry_change_func(obj, attr_name))
            entry.drop_target_register(DND_FILES)
            entry.dnd_bind('<<Drop>>', self.make_on_drop_func(entry, obj, attr_name))
            entry.config(validate="none")
            entry.delete(0, tk.END)  # 删除所有文本
            
            entry.insert(0, str(value))  # 插入新值
            entry.config(validate="key")

    
        if args_dict.get("need_h", ""):
            need_h = exec_and_return(args_dict["need_h"])
            iCurY += need_h + iGap
        else:
            iCurY += iHeight + iGap
        
        
        ui_dct = {
            "entry" : entry,
            "checkbox" : checkbox,
            "checkbox_var" : checkbox_var,
            "combo" : combo,
            "text_node" : text_node,
            "scale_node" : scale_node,
            "entry_int_range" : entry_int_range,
            "entry_float_range" : entry_float_range,
            "img_label" : img_label,
            "list_num_label":list_num_label,
        }
        
        show_dict[attr_name] = ui_dct
        
        if att_cls == "SimList":
            show_vaule = args_dict.get("show_vaule", True)
            if value and show_vaule:
                new_obj = value
                for i in range(len(value)):
                    new_one_data = {}
                    new_one_data["attr_name"] = i
                    new_one_data["show_name"] = None
                    new_one_data["att_cls"] = None
                    new_one_data["args_dict"] = {}
                    new_show_dict = {}
                    new_obj.__dict__["__show_dict"] = new_show_dict
                    iCurY = self.ui_add_one_attr(new_obj, cls, root, iCurY, new_one_data, new_show_dict, start_x+30)
        
        elif att_cls == "SimDict":
            show_vaule = args_dict.get("show_vaule", True)
            if value and show_vaule:
                new_obj = value
                keys = value.keys()
                for key in keys:
                    new_one_data = {}
                    new_one_data["attr_name"] = key
                    new_one_data["show_name"] = None
                    new_one_data["att_cls"] = None
                    new_one_data["args_dict"] = {}
                    new_show_dict = {}
                    new_obj.__dict__["__show_dict"] = new_show_dict
                    iCurY = self.ui_add_one_attr(new_obj, cls, root, iCurY, new_one_data, new_show_dict, start_x+30)
                    
        
        return iCurY, next_show_obj




    def ui_add_one_func(self, cls, root, iCurY, one_data, show_dict):
        left_width = self.get_left_with()
        right_width = self.get_right_with()
        btn_start = int((g_win_width - right_width)/2)
        
        iGap = g_default_gap
        iHeight = g_default_one_item_h
        
        func2 = one_data["func"]
        args_dict = one_data["args_dict"]
        
        sName = args_dict["sName"]
        if sName == None or len(sName.strip()) == 0:
            sName = func2.__name__
            
        btn_color = args_dict.get("color", "#00ffff")
        btn_x = btn_start
        btn_y = iCurY
        btn_w = self.get_right_with()
        btn_h = iHeight
        if args_dict.get("x", ""):
            btn_x = exec_and_return(args_dict["x"])

        if args_dict.get("y", ""):
            btn_y = iCurY + exec_and_return(args_dict["y"])
        
        if args_dict.get("w", ""):
            btn_w = exec_and_return(args_dict["w"])
        
        if args_dict.get("h", ""):
            btn_h = exec_and_return(args_dict["h"])

        
        # 创建一个按钮
        button = tk.Button(root, text=sName, command=func2, bg=btn_color)
        button.place(x=btn_x, y=btn_y, width=btn_w, height=btn_h)

        if args_dict.get("need_h", ""):
            need_h = exec_and_return(args_dict["need_h"])
            iCurY += need_h + iGap
        else:
            iCurY += btn_h + iGap
        
        return iCurY

    def ui_add_one_meth(self, obj, cls, root, iCurY, one_data, show_dict):
        left_width = self.get_left_with()
        right_width = self.get_right_with()
        btn_start = int((g_win_width - right_width)/2)

        funcTemp = one_data["func"]
        args_dict = one_data["args_dict"]
        iGap = g_default_gap
        iHeight = g_default_one_item_h
        
        
        btn_color = args_dict.get("color", "#ff00ff")


        btn_x = btn_start
        btn_y = iCurY
        btn_w = right_width
        btn_h = iHeight
        if args_dict.get("x", ""):
            btn_x = exec_and_return(args_dict["x"])

        if args_dict.get("y", ""):
            btn_y = iCurY + exec_and_return(args_dict["y"])
        
        if args_dict.get("w", ""):
            btn_w = exec_and_return(args_dict["w"])
        
        if args_dict.get("h", ""):
            btn_h = exec_and_return(args_dict["h"])


        sName = args_dict["sName"]
        if sName == None or len(sName.strip()) == 0:
            sName = funcTemp.__name__
        # 创建一个按钮
        button = tk.Button(root, text=sName, command=self.make_btn_func(funcTemp, obj), bg=btn_color)
        button.place(x=btn_x, y=btn_y, width=btn_w, height=btn_h)
        
        if args_dict.get("need_h", ""):
            need_h = exec_and_return(args_dict["need_h"])
            iCurY += need_h + iGap
        else:
            iCurY += btn_h + iGap
            
        return iCurY

    def make_on_checkbox_clicked(self, var, obj, attr_name):
        def on_checkbox_clicked():
            val = False
            if var.get() == 1:
                print("勾选框被选中")
                val = True
                
            
            self.raw_set_attr(obj, attr_name, val)
            obj._check_save_json(attr_name)
                
        return on_checkbox_clicked


    def raw_set_attr(self, obj, attr_name, value):
        ans_type = type(obj)
        if ans_type == SimList:
            obj[int(attr_name)] = value
        elif ans_type == SimDict:
            obj[attr_name] = value
        else:
            obj.__dict__[attr_name] = value


    def make_on_handle_selection(self, combo, obj, attr_name):

        def handle_selection(event):
            selected_value = combo.get()
            save_value = combo.show_to_value_dict[selected_value]
            print(f"Selected value: {selected_value} ({save_value})")
            
            self.raw_set_attr(obj, attr_name, save_value)
            obj._check_save_json(attr_name)
                
        return handle_selection

    def make_entry_change_func(self, obj, attr_name):
        def on_entry_changed(event):
            new_text = event.widget.get().strip()  # 获取输入框中的文本内容
            # print("Entry content changed to:", new_text)
            
            self.raw_set_attr(obj, attr_name, new_text)
            obj._check_save_json(attr_name)
        return on_entry_changed




    def make_int_range_entry_change_func(self, data_dict, obj, attr_name):
        def on_entry_changed(event):
            new_text = event.widget.get().strip()  # 获取输入框中的文本内容
            # print("Entry content changed to:", new_text)
            val = int(new_text)
        
            scale_node = data_dict["scale_node"]
            scale_node.set(val)
            
            self.raw_set_attr(obj, attr_name, val)
            obj._check_save_json(attr_name)
        return on_entry_changed


    def make_float_range_entry_change_func(self, data_dict, obj, attr_name):
        def on_entry_changed(event):
            new_text = event.widget.get().strip()  # 获取输入框中的文本内容
            # print("Entry content changed to:", new_text)
            val = float(new_text)
        
            scale_node = data_dict["scale_node"]
            scale_node.set(val)
            
            self.raw_set_attr(obj, attr_name, val)
            obj._check_save_json(attr_name)
        return on_entry_changed



    def make_text_node_change_func(self, text_node, obj, attr_name):
        def on_text_node_changed(event):
            new_text = text_node.get("1.0", tk.END)
            print("text_node content changed to:", new_text)
            
            self.raw_set_attr(obj, attr_name, new_text)
            obj._check_save_json(attr_name)
        return on_text_node_changed



    def make_text_node_on_drop_func(self, text_node, obj, attr_name):
        def text_node_on_drop(event):
            # 获取拖拽的文件路径
            # file_path = event.data
            # 在 Entry 中显示文件路径
            temp = text_node.tk.splitlist(event.data)
            text_content = ""
            
            for i in range(len(temp)):
                text_content += temp[i] + "\n"
                
            text_node.insert(tk.END, text_content)
            new_text = text_node.get("1.0", tk.END)
            
            self.raw_set_attr(obj, attr_name, new_text)
            obj._check_save_json(attr_name)
            
        return text_node_on_drop




    def make_on_drop_func(self, entry, obj, attr_name):
        def on_drop(event):
            # 获取拖拽的文件路径
            # file_path = event.data
            # 在 Entry 中显示文件路径
            temp = entry.tk.splitlist(event.data)
            file_path = temp[0]
            entry.delete(0, tk.END)  # 清空 Entry 中的内容
            entry.insert(0, file_path)  # 显示拖拽的文件路径
            
            
            self.raw_set_attr(obj, attr_name, file_path)
            obj._check_save_json(attr_name)
        return on_drop


    def refresh_import_by_path(self, img_label, file_path):
        target_width = 100  # 设定目标宽度
        target_height = 100  # 设定目标高度
        tk_image = g_purple_tk_image #加载失败就显示紫色图片
        try:
            image = Image.open(file_path)
            resized_image = image.resize((target_width, target_height))
            # 将调整大小后的图片转换为Tkinter支持的格式
            tk_image = ImageTk.PhotoImage(resized_image)
        except:
            pass
        img_label.config(image=tk_image)
        img_label.image = tk_image

    def make_img_on_drop_func(self, img_label, obj, attr_name):
        def on_drop(event):
            # 获取拖拽的文件路径
            # file_path = event.data
            # 在 Entry 中显示文件路径
            temp = img_label.tk.splitlist(event.data)

            file_path = temp[0]
            # print("file_path:", file_path)
            
            self.refresh_import_by_path(img_label, file_path)
            
            
            self.raw_set_attr(obj, attr_name, file_path)
            obj._check_save_json(attr_name)
        return on_drop


    def make_list_add_btn_func(self, obj, attr_name, list_label):

        def on_add_btn():
            
            oldData = obj.__dict__.get(attr_name, None)
            
            if oldData == None:
                oldData = SimList()
                oldData.set_parent(obj)
            
            oldData.append(None)
            
            list_label.config(text=str(len(oldData)))
            
            self.raw_set_attr(obj, attr_name, oldData)
            obj._check_save_json(attr_name)
            
            self.refesh_ui()
        return on_add_btn


    def make_list_del_btn_func(self, obj, attr_name, list_label):

        def on_add_btn():
            
            oldData = obj.__dict__.get(attr_name, None)
            
            if oldData == None:
                oldData = SimList()
                oldData.set_parent(obj)
            
            if len(oldData) >= 1:
                oldData.pop()
            
            list_label.config(text=str(len(oldData)))
            
            self.raw_set_attr(obj, attr_name, oldData)
            obj._check_save_json(attr_name)
            
            self.refesh_ui()
        return on_add_btn


    def make_dict_del_btn_func(self, obj, attr_name, list_label):

        def on_add_btn():
            
            oldData = obj.__dict__.get(attr_name, None)
            
            if oldData == None:
                oldData = SimDict()
                oldData.set_ui_obj(self)
                oldData.set_parent(obj)
            
            if len(oldData) >= 1:
                list1 = oldData.keys()
                list1.sort()
                oldData.pop(list1[-1])
            
            list_label.config(text=str(len(oldData)))
            
            self.raw_set_attr(obj, attr_name, oldData)
            obj._check_save_json(attr_name)
            
            self.refesh_ui()
        return on_add_btn


    def make_dict_add_btn_func(self, obj, attr_name, list_label):

        def on_add_btn():
            
            oldData = obj.__dict__.get(attr_name, None)
            
            if oldData == None:
                oldData = SimDict()
                oldData.set_ui_obj(self)
                oldData.set_parent(obj)
            
            ilen = len(oldData)
            oldData[ilen+1] = None
            
            list_label.config(text=str(len(oldData)))
            
            self.raw_set_attr(obj, attr_name, oldData)
            obj._check_save_json(attr_name)
            
            self.refesh_ui()
        return on_add_btn


    def make_show_vaule_up_btn_func(self, obj, attr_name, args_dict):
        def on_btn():
            args_dict["show_vaule"] = False
            self.refesh_ui()
            
        return on_btn


    def make_show_vaule_down_btn_func(self, obj, attr_name, args_dict):
        def on_btn():
            args_dict["show_vaule"] = True
            self.refesh_ui()
            
        return on_btn

    def set_entry_value(self, entry, value):
        if not entry:
            return
        
        entry.config(validate="none")
        entry.delete(0, tk.END)
        entry.insert(0, str(value))
        entry.config(validate="key")
        
        

    def update_value(self, obj, key, value):
        data = simobj.get_class_attr_to_data(obj.__class__, key)
        if data != None:
            att_cls = data["att_cls"]
            if att_cls != None and att_cls == "SimList":
                new_list = SimList()
                new_list.set_parent(obj)
                
                new_list.extend(value)
                value = new_list
            elif att_cls != None and att_cls == "SimDict":
                new_dict = SimDict()
                new_dict.set_ui_obj(self)
                new_dict.set_parent(obj)
                
                new_dict.update(value)
                value = new_dict
        
        
        self.update_ui(obj, key, value)
        return value

    def update_ui(self, obj, key, value):
        is_show = obj.__dict__.get("__showing", False)
        if not is_show:
            return
            
        show_dict = obj.__dict__.get("__show_dict", {})
        
        ui_dct = show_dict.get(key, None)
        if not ui_dct:
            return
        
        entry = ui_dct["entry"]
        
        if entry:
            self.set_entry_value(entry, value)
        
        entry_int_range = ui_dct["entry_int_range"]
        if entry_int_range:
            self.set_entry_value(entry_int_range, value)
            
        entry_float_range = ui_dct["entry_float_range"]
        if entry_int_range:
            self.set_entry_value(entry_float_range, value)

        checkbox = ui_dct["checkbox"]
        if checkbox:
            ui_dct["checkbox_var"].set(value)
        
        
        combo = ui_dct["combo"]
        if combo:
            dct = combo.value_to_show_dict
            combo.set(dct[value])
        
        text_node = ui_dct["text_node"]
        if text_node:
            # text_node.config(validate="none")
            text_node.delete("1.0", tk.END)
            text_node.insert(tk.END, value)
            # text_node.config(validate="key")

        img_label = ui_dct["img_label"]
        if img_label:
            self.refresh_import_by_path(img_label, value)
        
        
        scale_node = ui_dct["scale_node"]
        if scale_node:
            scale_node.set(value)
            
            
        list_num_label = ui_dct["list_num_label"]
        if list_num_label:
            list_val = value
            if not list_val:
                list_val = []
            str1 = str(len(list_val))
            list_num_label.set(str1)
        


    def make_btn_func(self, funcTemp, obj):
        def on_button_click():
            funcTemp(obj)
        return on_button_click




    def refesh_ui(self):
        iCurY = 0


        for widget in self.m_inner_frame.winfo_children():
            widget.destroy()

        
        add_all_func = False
        if self.m_root == simui.simmain.get_root_TK():#主UI才显示这个
            add_all_func = True
        for obj in self.m_show_obj_list:
            iCurY = self.make_TK_node(obj, iCurY, add_all_func)
            add_all_func = False
        
        self.m_inner_frame.config(width=g_win_width, height=iCurY+100)


    def add_show_obj(self, obj):
        if obj not in self.m_show_obj_list:
            self.m_show_obj_list.append(obj)
    
    def set_show_obj_list(self, list1):
        self.m_show_obj_list = list1


    def is_show_obj(self, obj):
        if obj in self.m_show_obj_list:
            return True
        return False
    
    def show(self, title=None):
        root = self.m_root
        self.initTK(root, title)
        self.refesh_ui()
        simui.simmain.add_show_win(self)
        root.mainloop()

















