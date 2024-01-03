import json
import keyboard
import random
import time
import os
from pynput import mouse
from ctypes import windll

user32 = windll.user32
kernel32 = windll.kernel32
psapi = windll.psapi
# TODO: 后面加一个自动识别游戏窗口的功能，避免在外面误触

class QuickCastManager:
    def __init__(self):
        self.settings = self.load_settings()
        self.quick_casts = self.load_quick_casts()
        self.save_settings()
        self.save_quick_casts()
        self.select_cast = None
        self.x1 = None
        self.x2 = None
        self.lock = False
        mouse_listener = mouse.Listener(on_click=self.on_click)
        # 启动监听器
        mouse_listener.start()

    def on_click(self, x,y, button, pressed):
        if not pressed:
            return
        if button == mouse.Button.x1:
            if self.x1 is None:
                # print("请先设置x1的快捷施法方案")
                return
            self.run_combo(self.x1)
        elif button == mouse.Button.x2:
            if self.x2 is None:
                # print("请先设置x2的快捷施法方案")
                return
            self.run_combo(self.x2)

    def load_settings(self):
        try:
            with open("setting.json", "r") as file:
                return json.load(file)
        except FileNotFoundError:
            # print("setting.json 文件未找到，默认使用默认设置。")
            return {"key_up_interval": 0.01, "key_interval": 0.08}

    def load_quick_casts(self):
        try:
            with open("quick.json", "r") as file:
                return json.load(file)
        except FileNotFoundError:
            # print("quick.json 文件未找到，默认使用空快捷施法方案。")
            return {}

    def save_quick_casts(self):
        with open("quick.json", "w+") as file:
            json.dump(self.quick_casts, file, indent=2)

    def save_settings(self):
        with open("setting.json", "w+") as file:
            json.dump(self.settings, file, indent=2)

    def change_settings(self, key_interval, key_up_interval):
        self.settings["key_interval"] = key_interval
        self.settings["key_up_interval"] = key_up_interval
        self.save_settings()

    def change_settings(self):
        print("当前设置:")
        print(f"按键间隔: {self.settings['key_interval']}秒")
        print(f"按键松开间隔: {self.settings['key_up_interval']}秒")
        new_key_interval = input("请输入新的按键间隔(秒): ")
        new_key_up_interval = input("请输入新的按键松开间隔(秒): ")
        try:
            self.settings["key_interval"] = float(new_key_interval)
            self.settings["key_up_interval"] = float(new_key_up_interval)
        except ValueError:
            print("无效的输入，请重新输入。")
            return
        self.save_settings()

    def display_quick_casts(self):
        print("当前可用的快捷施法方案:")
        for name, combos in self.quick_casts.items():
            print(f"{name}:{len(combos)}个快捷键")
            # for combo in combos:
            #     print(f"  {combo['trigger_key']}: {combo['sequence']}")

    def create_new_cast(self):
        name = input("请输入新快捷施法方案的名称: ")
        combos = []
        self.quick_casts[name] = combos
        while True:
            self.add_combo_to_cast(name)
            another = input("是否要添加另一个组合？(y/n): ").lower()
            if another != 'y':
                break
        self.save_quick_casts()

    def create_new_cast(self, name, combos):
        self.quick_casts[name] = combos
        self.save_quick_casts()

    def delete_cast(self):
        self.display_quick_casts()
        name = input("请输入要删除的快捷施法方案: ")
        if name not in self.quick_casts:
            print(f"找不到名为 {name} 的快捷施法方案。")
            return
        del self.quick_casts[name]
        self.save_quick_casts()

    def delete_cast(self, name):
        if name not in self.quick_casts:
            return False
        del self.quick_casts[name]
        self.save_quick_casts()
        return True

    def delete_combo_from_cast(self, cast_name,trigger_key):
        if cast_name not in self.quick_casts:
            return False
        find = False
        for combo in self.quick_casts[cast_name]:
            if combo["trigger_key"] == trigger_key:
                self.quick_casts[cast_name].remove(combo)
                find = True
                break
        self.save_quick_casts()
        return find

    def add_combo_to_cast(self, cast_name, hotkey=False):
        trigger_key = input("请输入触发键: a?(alt+?)或者f?或者x1(鼠标后)、x2(鼠标前):")
        if len(trigger_key) == 2 and trigger_key[0] == "a":
            trigger_key = "alt+" + trigger_key[1]
        elif len(trigger_key) == 2 and trigger_key[0] == "f" and trigger_key[1] in ["1", "2", "3"]:
            pass
        if len(trigger_key) == 2 and trigger_key[0] == "p": 
            pass
        elif trigger_key == "x1" or trigger_key == "x2":
            pass
        else:
            print(f"无效的触发键 {trigger_key}，请重新输入。")
            return
        sequence = input("请输入按键顺序，用空格分隔，`?(n*50ms延迟)").split()
        # 检查是否存在是否存在大于一个键的
        for key in sequence:
            if key[0] == "`" and len(key) == 2:
                pass
            elif len(key) > 1:
                print(f"无效的按键 {key}，请重新输入。")
                return
        if cast_name in self.quick_casts:
            # 如果存在就覆盖
            find = False
            for combo in self.quick_casts[cast_name]:
                if combo["trigger_key"] == trigger_key:
                    combo["sequence"] = sequence
                    find = True
                    break        
            if not find:
                self.quick_casts[cast_name].append({"trigger_key": trigger_key, "sequence": sequence})
        else:
            print(f"找不到名为 {cast_name} 的快捷施法方案。")
        if hotkey:
            if trigger_key in ["x1","x2"]:
                if trigger_key == "x1":
                    self.x1 = {"trigger_key": "x1", "sequence": sequence}
                if trigger_key == "x2":
                    self.x2 = {"trigger_key": "x2", "sequence": sequence}
            else:
                keyboard.add_hotkey("alt+" + trigger_key, self.run_combo, args=({"trigger_key": trigger_key, "sequence": sequence},))
        self.save_quick_casts()

    def add_combo_to_cast(self, cast_name,trigger_key,sequence, hotkey=False):
        if cast_name in self.quick_casts:
            # 如果存在就覆盖
            find = False
            for combo in self.quick_casts[cast_name]:
                if combo["trigger_key"] == trigger_key:
                    combo["sequence"] = sequence
                    find = True
                    break        
            if not find:
                self.quick_casts[cast_name].append({"trigger_key": trigger_key, "sequence": sequence})
        else:
            return
        if hotkey:
            if trigger_key in ["x1","x2"]:
                if trigger_key == "x1":
                    self.x1 = {"trigger_key": "x1", "sequence": sequence}
                if trigger_key == "x2":
                    self.x2 = {"trigger_key": "x2", "sequence": sequence}
            else:
                keyboard.add_hotkey("alt+" + trigger_key, self.run_combo, args=({"trigger_key": trigger_key, "sequence": sequence},))
        self.save_quick_casts()


    def run_listener(self):
        # 选择方案
        self.display_quick_casts()
        cast_name = input("请输入要使用的快捷施法方案: ")
        if cast_name not in self.quick_casts:
            print(f"找不到名为 {cast_name} 的快捷施法方案。")
            return
        self.select_cast = cast_name
        for combo in self.quick_casts[cast_name]:
            if combo["trigger_key"] in ["x1","x2"]:
                if combo["trigger_key"] == "x1":
                    self.x1 = combo
                if combo["trigger_key"] == "x2":
                    self.x2 = combo
            else:
                keyboard.add_hotkey(combo["trigger_key"], self.run_combo, args=(combo,))
        for combo in self.quick_casts[cast_name]:
            print(f"{combo['trigger_key']}:{combo['sequence']}")
        print(f"已开始监听 {cast_name} 快捷施法方案。")
        while True:
            print("1. 添加新组合")
            print("2. 停止监听")
            choice = input("请输入选项 (1-2): ")
            if choice == "1":
                self.add_combo_to_cast(cast_name,hotkey=True)
                # 清屏重新输出
                print("--------------------")
                for combo in self.quick_casts[cast_name]:
                    print(f"{combo['trigger_key']}:{combo['sequence']}")
            elif choice == "2":
                try:
                    keyboard.unhook_all_hotkeys()
                except Exception as e:
                    find = False
                    for combo in self.quick_casts[cast_name]:
                        if combo["trigger_key"] not in ["x1","x2"]:
                            find = True
                            break
                    if find:
                        print("内部错误:停止监听失败:",e)
                        exit()
                self.x1 = None
                self.x2 = None
                break
            else:
                print("无效的选项，请重新输入。")
    
    def run_listener(self,cast_name):
        # 选择方案
        if cast_name not in self.quick_casts:
            return False
        self.select_cast = cast_name
        for combo in self.quick_casts[cast_name]:
            if combo["trigger_key"] in ["x1","x2"]:
                if combo["trigger_key"] == "x1":
                    self.x1 = combo
                if combo["trigger_key"] == "x2":
                    self.x2 = combo
            else:
                keyboard.add_hotkey(combo["trigger_key"], self.run_combo, args=(combo,))
        self.cast_name = cast_name
        return True

    def stop_listener(self):
        self.x1 = None
        self.x2 = None
        try:
            keyboard.unhook_all_hotkeys()
        except Exception as e:
            find = False
            for combo in self.quick_casts[self.cast_name]:
                if combo["trigger_key"] not in ["x1","x2"]:
                    find = True
                    break
            if find:
                return False
        return True
        

    def run_combo(self, combo):
        if self.lock:
            return
        self.lock = True
        # 如果当前存在alt按键被按下，等待按键释放
        while keyboard.is_pressed('alt'):
            time.sleep(0.001)
        # for key in combo['sequence']:
        #     keyboard.press(key)
        #     delay = self.settings["key_interval"] * random.uniform(0.66, 1.33)
        #     time.sleep(delay)
        #     keyboard.release(key)
        #     delay = self.settings["key_up_interval"] * random.uniform(0.66, 1.33)
        #     time.sleep(delay)
        timeline_now = 0
        keys = []
        for key in combo['sequence']:
            if key[0] == "`":
                delay = float(0)
                # 为了提高准确度并且操作更加自然，随机延迟设定如下
                delay += random.uniform(0.95, 1.051) * 0.001 * int(key[1:])
                timeline_now += delay
                continue
            keys.append((key,True,timeline_now))
            delay = self.settings["key_up_interval"] * random.uniform(0.952, 1.05)
            keys.append((key,False,timeline_now+delay))
            timeline_now += self.settings["key_interval"] * random.uniform(0.82, 1.19)
        keys.sort(key=lambda x: x[2])
        for i,key in enumerate(keys):
            if i != 0:
                time.sleep(key[2]-keys[i-1][2])
            if key[1]:
                keyboard.press(key[0])
            else:
                keyboard.release(key[0])
        self.lock = False

def main():
    os.chdir(os.path.dirname(os.path.realpath(__file__)))
    manager = QuickCastManager()
    random.seed()
    while True:
        print("1. 新建快捷施法方案")
        print("2. 选择已有方案")
        print("3. 删除已有方案")
        print("4. 设置")
        print("5. 退出程序")

        choice = input("请输入选项 (1-3): ")

        if choice == "1":
            manager.create_new_cast()
        elif choice == "2":
            manager.run_listener()
        elif choice == "3":
            manager.delete_cast()
        elif choice == "4":
            manager.change_settings()
        elif choice == "5":
            exit()
        else:
            print("无效的选项，请重新输入。")

if __name__ == "__main__":
    main()
