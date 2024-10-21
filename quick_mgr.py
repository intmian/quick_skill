import json
import keyboard
import random
import time
import os
from pynput import mouse
from ctypes import windll
mouseController = mouse.Controller()
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

    # 鼠标按键监听
    def on_click(self, x,y, button, pressed):
        if not pressed:
            return
        if button == mouse.Button.x1:
            if self.x1 is None:
                return
            self.run_combo(self.x1)
        elif button == mouse.Button.x2:
            if self.x2 is None:
                return
            self.run_combo(self.x2)

    def load_settings(self):
        try:
            with open("setting.json", "r") as file:
                return json.load(file)
        except FileNotFoundError:
            return {"key_up_interval": 0.01, "key_interval": 0.08}

    def load_quick_casts(self):
        try:
            with open("quick.json", "r") as file:
                return json.load(file)
        except FileNotFoundError:
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

    def create_new_cast(self, name, combos):
        self.quick_casts[name] = combos
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
            if len(key) >= 3 and (key[:2] == "lp" or key[:2] == "rp"):
                delay = float(0)
                delay += random.uniform(0.99, 1.01) * 0.001 * int(key[2:])
                if key[0] == "l":
                    keys.append(("MLeft",True,timeline_now))
                    keys.append(("MLeft",False,timeline_now+delay))
                if key[0] == "r":
                    keys.append(("MRight",True,timeline_now))
                    keys.append(("MRight",False,timeline_now+delay))
                timeline_now += delay
                continue
            real_key = key
            if key == "上":
                real_key = "up"
            elif key == "下":
                real_key = "down"
            elif key == "左":
                real_key = "left"
            elif key == "右":
                real_key = "right"
            keys.append((real_key,True,timeline_now))
            delay = self.settings["key_up_interval"] * random.uniform(0.952, 1.05)
            keys.append((real_key,False,timeline_now+delay))
            timeline_now += self.settings["key_interval"] * random.uniform(0.82, 1.19)
        keys.sort(key=lambda x: x[2])
        for i,key in enumerate(keys):
            print(key)
            if i != 0:
                time.sleep(key[2]-keys[i-1][2])
            if key[0] == "MLeft":
                if key[1]:
                    mouseController.press(mouse.Button.left)
                else:
                    mouseController.release(mouse.Button.left)
                continue
            elif key[0] == "MRight":
                if key[1]:
                    mouseController.press(mouse.Button.right)
                else:
                    mouseController.release(mouse.Button.right)
                continue
            if key[1]:
                keyboard.press(key[0])
            else:
                keyboard.release(key[0])
        self.lock = False