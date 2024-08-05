import os
import random
import sys
import time
import tkinter as tk
from tkinter import ttk,simpledialog,messagebox
from quick_mgr import QuickCastManager
from sys import exit
import ctypes

def InputToShow(key):
    if key[0] == '`':
        return f"空{key[1:]}ms"
    elif len(key) >= 3 and (key[:2] == "lp" or key[:2] == "rp"):
        s = ""
        if key[:2] == "lp":
            s = "按住左键"
        else:
            s = "按住右键"
        return f"{s}({key[2:]}ms)"
    else:
        return key

# 界面管理
class ui:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("快速施法")

        # 窗口大小
        window_x = 300
        window_y = 300
        # 基础控件大小，为了避免官方的gird布局的问题，这里自己实现一份
        weight_base = 30
        height_base = 30
        # 禁止调整大小
        self.window.resizable(False, False)

        # DPI
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
        window_x = int(window_x * scale)
        window_y = int(window_y * scale)
        weight_base = int(weight_base * scale)
        height_base = int(height_base * scale)

        # 移动到屏幕中央
        self.window.geometry(f"{window_x}x{window_y}")
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        self.window.geometry("%dx%d+%d+%d" % (window_x, window_y, (screen_width - window_x) / 2, (screen_height - window_y) / 2))
        self.window.focus_force()
        
        left_space = weight_base * 0.5
        top_space = height_base * 0.5

        combobox_width = weight_base * 3
        combobox_height = height_base * 0.7
        combobox_x = 0 + left_space
        combobox_y = 0 + top_space

        start_button_width = weight_base * 2
        start_button_height = height_base * 1
        start_button_x = weight_base * 7 + left_space
        start_button_y = height_base * 1 + top_space

        stop_button_width = weight_base * 2
        stop_button_height = height_base * 1
        stop_button_x = weight_base * 7 + left_space
        stop_button_y = height_base * 2 + top_space

        clear_button_width = weight_base * 2
        clear_button_height = height_base * 1
        clear_button_x = weight_base * 7 + left_space
        clear_button_y = height_base * 4 + top_space

        self.combobox = ttk.Combobox(self.window)
        self.combobox.place(x=combobox_x, y=combobox_y, width=combobox_width, height=combobox_height)
        self.combobox["values"] = ("<新增方案>")
        self.combobox.current(0)

        listbox_width = weight_base * 6
        listbox_height = int(height_base * 8)
        listbox_x = 0 + left_space
        listbox_y = int(float(height_base) * 1 + top_space)

        self.listbox = tk.Listbox(self.window, width=listbox_width, height=listbox_height)
        self.listbox.place(x=listbox_x, y=listbox_y, width=listbox_width, height=listbox_height)

        self.start_button = tk.Button(self.window, text="开始")
        self.start_button.place(x=start_button_x, y=start_button_y, width=start_button_width, height=start_button_height)
        self.stop_button = tk.Button(self.window, text="停止")
        self.stop_button.place(x=stop_button_x, y=stop_button_y, width=stop_button_width, height=stop_button_height)
        self.clear_button = tk.Button(self.window, text="删除方案")
        self.clear_button.place(x=clear_button_x, y=clear_button_y,width=clear_button_width, height=clear_button_height)

        key_interval_label_width = weight_base * 2
        key_interval_label_height = height_base *  0.7
        key_interval_label_x = weight_base * 7 + left_space
        key_interval_label_y = height_base * 5 + top_space
        self.key_interval_label = tk.Label(self.window, text="按键间隔秒")
        self.key_interval_label.place(x=key_interval_label_x, y=key_interval_label_y, width=key_interval_label_width, height=key_interval_label_height)

        key_interval_entry_width = weight_base * 2
        key_interval_entry_height = height_base *  0.7
        key_interval_entry_x = weight_base * 7 + left_space
        key_interval_entry_y = height_base * 6 + top_space
        self.key_interval_entry = tk.Entry(self.window)
        self.key_interval_entry.place(x=key_interval_entry_x, y=key_interval_entry_y, width=key_interval_entry_width, height=key_interval_entry_height)

        key_up_interval_label_width = weight_base * 2
        key_up_interval_label_height = height_base * 0.7
        key_up_interval_label_x = weight_base * 7 + left_space
        key_up_interval_label_y = height_base * 7 + top_space
        self.key_up_interval_label = tk.Label(self.window, text="抬起间隔秒")
        self.key_up_interval_label.place(x=key_up_interval_label_x, y=key_up_interval_label_y, width=key_up_interval_label_width, height=key_up_interval_label_height)

        key_up_interval_entry_width = weight_base * 2
        key_up_interval_entry_height = height_base * 0.7
        key_up_interval_entry_x = weight_base * 7 + left_space
        key_up_interval_entry_y = height_base * 8 + top_space
        self.key_up_interval_entry = tk.Entry(self.window)
        self.key_up_interval_entry.place(x=key_up_interval_entry_x, y=key_up_interval_entry_y, width=key_up_interval_entry_width, height=key_up_interval_entry_height)
        #  主窗口销毁后，必须销毁整个程序，不然会出现已经开始的监听阻塞住主线程关闭，倒是无法真正关闭程序
        self.window.protocol("WM_DELETE_WINDOW", exit)

    def set_select(self, f):
        self.combobox.bind("<<ComboboxSelected>>", f)
    
    def set_click_start(self, f):
        # 如果绑定按下事件，切鼠标弹起时没有在按钮上，就会出现按键陷下去的情况……所以绑定鼠标弹起事件
        self.start_button.bind("<ButtonRelease-1>", f)

    def set_click_stop(self, f):
        self.stop_button.bind("<ButtonRelease-1>", f)

    def set_click_clear(self, f):
        self.clear_button.bind("<ButtonRelease-1>", f)

    def set_listbox_select(self, f):
        self.listbox.bind("<<ListboxSelect>>", f)
    
    def set_listbox_double_click(self, item_f):
        self.listbox.bind("<Double-Button-1>", item_f)
    
    # 一个方案开始监听后页面相关
    def on_start(self):
        self.key_interval_entry.config(state="disabled")
        self.key_up_interval_entry.config(state="disabled")
        self.combobox.config(state="disabled")
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.clear_button.config(state="disabled")
    
    # 一个方案停止监听后页面相关
    def on_stop(self):
        self.key_interval_entry.config(state="normal")
        self.key_up_interval_entry.config(state="normal")
        self.combobox.config(state="normal")
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.clear_button.config(state="normal")

    # 选择某个方案后页面相关
    def on_select_cast(self,cast_name,cast_list):
        self.key_interval_entry.config(state="normal")
        self.key_up_interval_entry.config(state="normal")
        self.combobox.config(state="normal")
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.clear_button.config(state="normal")
        str_list = []
        for combo in cast_list:
            trigger_key = combo["trigger_key"]
            if trigger_key == "x1":
                trigger_key = "鼠标侧后键"
            elif trigger_key == "x2":
                trigger_key = "鼠标侧前键"
            sequence = ""
            for key in combo["sequence"]:
                sequence += InputToShow(key)
                sequence += " "
            str_list.append(f"{trigger_key}: {sequence}")
        str_list.append("<双击新增|双击已有项删除>")
        self.combobox.set(cast_name)
        self.update_list(str_list)
    
    # 外部添加某个快捷键
    def on_add_combo(self,combo):
        trigger_key = combo["trigger_key"]
        if trigger_key == "x1":
            trigger_key = "鼠标侧后键"
        elif trigger_key == "x2":
            trigger_key = "鼠标侧前键"
        sequence = ""
        for key in combo["sequence"]:
            sequence += InputToShow(key)
        self.listbox.insert(tk.END, f"{trigger_key}: {sequence}")
        # 将<双击新增|双击已有项删除>移到最后
        for i in range(len(self.listbox.get(0, tk.END)) - 1):
            if self.listbox.get(i) == "<双击新增|双击已有项删除>":
                self.listbox.delete(i)
                self.listbox.insert(tk.END, "<双击新增|双击已有项删除>")

    # 外部删除某个快捷键
    def on_delete_combo(self,trigger_key):
        if trigger_key == "x1":
            trigger_key = "鼠标侧后键"
        elif trigger_key == "x2":
            trigger_key = "鼠标侧前键"
        self.listbox.delete(self.listbox.get(0, tk.END).index(trigger_key))

    # 外部设置按键间隔
    def on_set_key_interval(self, interval,up_interval):
        self.key_interval_entry.delete(0, tk.END)
        self.key_interval_entry.insert(0, interval)
        self.key_up_interval_entry.delete(0, tk.END)
        self.key_up_interval_entry.insert(0, up_interval)

    def add_select(self, text):
        values = self.combobox["values"]
        new_values = values + (text,)
        self.combobox["values"] = new_values
        self.combobox.current(len(new_values) - 1)
        # 将<新增方案>移到最后
        for i in range(len(self.combobox["values"]) - 1):
            if self.combobox["values"][i] == "<新增方案>":
                self.combobox["values"] = self.combobox["values"][:i] + self.combobox["values"][i+1:]
                self.combobox["values"] += ("<新增方案>",)
                break
    
    def remove_select(self, text):
        self.combobox["values"] = tuple(filter(lambda x: x != text, self.combobox["values"]))
    
    def update_list(self, list):
        self.listbox.delete(0, tk.END)
        for i in list:
            self.listbox.insert(tk.END, i)

    def start(self):
        self.window.mainloop()
    
class mgr:
    def __init__(self,ui_mgr:ui,quick_mgr:QuickCastManager) -> None:
        self.ui_mgr = ui_mgr
        self.quick_mgr = quick_mgr

        self.now_choose_cast = ""
        self.is_start = False

        for i,cast_name in enumerate(self.quick_mgr.quick_casts):
            self.ui_mgr.add_select(cast_name)
            if i == len(self.quick_mgr.quick_casts) - 1:
                self.select_cast(cast_name)
        # 添加一下事件
        self.ui_mgr.set_select(self.on_ui_select)
        self.ui_mgr.set_listbox_double_click(self.on_ui_item_double_click)
        self.ui_mgr.set_click_start(self.on_ui_start)
        self.ui_mgr.set_click_stop(self.on_ui_stop)
        self.ui_mgr.set_click_clear(self.on_ui_clear)

        # 如果没有方案则在打开时必须新增
        if len(self.quick_mgr.quick_casts) == 0:
            self.add_cast(True)
            self.select_default_cast()
        
        self.ui_mgr.on_set_key_interval(self.quick_mgr.settings["key_interval"],self.quick_mgr.settings["key_up_interval"])

    def select_default_cast(self):
        default = ""
        for i,cast_name in enumerate(self.quick_mgr.quick_casts):
            if i == 0:
                default = cast_name
        self.select_cast(default)

    def on_ui_start(self,event):
        if self.is_start:
            return
        self.is_start = True
        self.quick_mgr.settings["key_interval"] = float(self.ui_mgr.key_interval_entry.get())
        self.quick_mgr.settings["key_up_interval"] = float(self.ui_mgr.key_up_interval_entry.get())
        self.quick_mgr.save_settings()
        self.quick_mgr.run_listener(self.now_choose_cast)
        self.ui_mgr.on_start()

    def on_ui_stop(self,event):
        if not self.is_start:
            return
        self.is_start = False
        self.quick_mgr.stop_listener()
        self.ui_mgr.on_stop()

    def on_ui_clear(self,event):
        if self.is_start:
            tk.messagebox.showerror("删除方案", "请先停止监听")
            return
        self.on_del_casts()

    def select_cast(self,cast_name):
        self.ui_mgr.on_select_cast(cast_name,self.quick_mgr.quick_casts[cast_name])
        self.now_choose_cast = cast_name

    def on_ui_select(self,event:tk.Event):
        if self.ui_mgr.combobox.get() == "<新增方案>":
            self.add_cast()
            return
        self.select_cast(self.ui_mgr.combobox.get())

    def on_ui_item_double_click(self,event):
        selected_index = self.ui_mgr.listbox.nearest(event.y)
        if selected_index >= 0:
            selected_item = self.ui_mgr.listbox.get(selected_index)
            if selected_item == "<双击新增|双击已有项删除>":
                self.add_combo()
            else:
                yes = simpledialog.messagebox.askyesno("删除快捷键", "是否删除该快捷键")
                if yes:
                    # 提取出触发键
                    trigger_key = selected_item.split(":")[0]
                    if trigger_key == "鼠标侧后键":
                        trigger_key = "x1"
                    elif trigger_key == "鼠标侧前键":
                        trigger_key = "x2"
                    self.del_comb(trigger_key,selected_item)

    def del_comb(self, trigger_key,selected_item):
        self.quick_mgr.delete_combo_from_cast(self.now_choose_cast,trigger_key)
        self.ui_mgr.on_delete_combo(selected_item)
        self.quick_mgr.stop_listener()
        self.quick_mgr.run_listener(self.now_choose_cast)

    def on_del_casts(self):
        yes = simpledialog.messagebox.askyesno("删除方案", "是否删除该方案")
        if yes:
            # 删除当前方案，选择第一个方案，如果没有方案则触发新增
            self.quick_mgr.delete_cast(self.now_choose_cast)
            self.ui_mgr.remove_select(self.now_choose_cast)
            if len(self.quick_mgr.quick_casts) == 0:
                self.add_cast(True)
            self.select_default_cast()

    def add_combo(self):
        # 以下非常扭曲的写法是为了避免第二个窗口不能获得焦点的bug，具体见https://stackoverflow.com/questions/54043323/tkinter-simpledialog-boxes-not-getting-focus-in-windows-10-with-python3
        root = tk.Tk()
        root.withdraw()
        root.update_idletasks()
        trigger_key = simpledialog.askstring("新增快捷键", "请输入触发键,额外支持鼠标侧键后x1,前x2",parent=root)
        if trigger_key == None or trigger_key == "":
            return
        # 检测是否包含空格
        if " " in trigger_key:
            simpledialog.messagebox.showerror("新增方案", "触发键不能包含空格")
            return
        # 检测是否已经存在
        for combo in self.quick_mgr.quick_casts[self.now_choose_cast]:
            if combo["trigger_key"] == trigger_key:
                simpledialog.messagebox.showerror("新增方案", "已经存在该快捷键，请先删除旧的")
                return
        root.update_idletasks()
        sequence = simpledialog.askstring("新增快捷键", "请输入按键序列,支持方向键（上下左右）,按住左键（lpxx 按住xx毫秒），按住右键（rptxx 按住xx毫秒）,额外支持`n表示空n ms,对应技能后摇等,以空格分隔",parent=root)
        if sequence == None or sequence == "":
            return
        
        sequence = sequence.split()
        # 检查是否存在是否存在大于一个键的
        for key in sequence:
            if key[0] == "`":
                # 遍历下是否都为数字，且小于10000
                if not key[1:].isdigit() or int(key[1:]) > 10000:
                    simpledialog.messagebox.showerror("新增方案", "无效的按键序列")
                    return
            elif len(key) >= 3 and (key[:2] == "lp" or key[:2] == "rp"):
                if not key[2:].isdigit() or int(key[2:]) > 10000:
                    simpledialog.messagebox.showerror("新增方案", "无效的按键序列")
                    return
            elif len(key) > 1:
                simpledialog.messagebox.showerror("新增方案", "无效的按键序列")
                return
        self.quick_mgr.add_combo_to_cast(self.now_choose_cast,trigger_key,sequence,self.is_start)
        self.ui_mgr.on_add_combo({"trigger_key":trigger_key,"sequence":sequence})

    def add_cast(self,first=False):
        # 弹窗，用户输入方案名
        last_casts_name = self.now_choose_cast
        self.ui_mgr.combobox.set("")
        self.ui_mgr.listbox.delete(0, tk.END)
        if not first:
            cast_name = simpledialog.askstring("新增方案", "请输入方案名")
        else:
            cast_name = simpledialog.askstring("新增方案", "没有方案，请输入方案名",initialvalue="默认方案")
        if cast_name == None or cast_name == "":
            if last_casts_name == "":
                tk.messagebox.showerror("新增方案", "没有方案时不能拒绝新增，程序即将推出")
                exit()
            self.select_cast(last_casts_name)
            return
        self.quick_mgr.create_new_cast(cast_name,[])
        self.select_cast(cast_name)
        self.ui_mgr.add_select(cast_name)
        self.ui_mgr.on_select_cast(cast_name,[])
        
    def start(self):
        self.ui_mgr.start()
    

def main():
    # 打包的程序的实际路径是一个临时目录，所以注释
    # os.chdir(os.path.dirname(os.path.realpath(__file__)))
    # 如果没有管理员权限就以管理员权限重新启动
    if ctypes.windll.shell32.IsUserAnAdmin() == 0:
        messagebox.showwarning("警告", "未以管理员身份运行，将以管理员权限重新启动")
        # ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1 | 0x40)
        return
    random.seed()
    m = mgr(ui(),QuickCastManager())
    m.start()

if __name__ == "__main__":
    main()