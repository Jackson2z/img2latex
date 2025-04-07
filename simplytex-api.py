import requests
import json  # 添加 json 模块导入语句
import pyautogui  # 新增截图库
import pyperclip  # 新增剪贴板库
import io
import win32clipboard
from PIL import ImageGrab
import time
import keyboard  # 新增热键监听库
import threading
import os  # 添加在文件顶部导入区域
import win32api  # 添加在现有win32导入附近
import sys
import win32con
import win32gui  # 新增Win32 GUI模块
from win32api import GetModuleHandle



SIMPLETEX_UAT="XXXXXX申请密钥后填入XXXXXXX" #! 申请密钥后填入

api_url="https://server.simpletex.cn/api/latex_ocr_turbo"  #! 接口地址
"""
    轻量公式识别模型API:https://server.simpletex.cn/api/latex_ocr_turbo
    标准公式识别模型API https://server.simpletex.cn/api/latex_ocr
    通用图片识别-轻量模型接口地址：https://server.simpletex.cn/api/simpletex_ocrc
	
"""


data = { } # 请求参数数据（非文件型参数），视情况填入，可以参考各个接口的参数说明
header={ "token": SIMPLETEX_UAT } # 鉴权信息，此处使用UAT方式
# inline_formula_wrapper=["$","$"] # 行内公式包裹符，默认为["$","$"]，即$...$
# isolated_formula_wrapper=["$$","$$"] # 行间公式包裹符，默认为["$$","$$"]，即$$...$$




# 新增后台服务功能
def create_tray_icon():
    """创建系统托盘图标"""
    WM_APP = win32con.WM_APP + 1
    IDI_APPLICATION = 32512

    class TrayIcon:
        def __init__(self):
            self.hinst = GetModuleHandle(None)
            message_map = {
                win32con.WM_DESTROY: self.on_destroy,
                WM_APP: self.on_app,
                win32con.WM_COMMAND: self.on_command,
            }
            wc = win32gui.WNDCLASS()
            wc.hInstance = self.hinst
            wc.lpszClassName = "SimplyTexTray"
            wc.lpfnWndProc = message_map
            self.classAtom = win32gui.RegisterClass(wc)
            
            # 创建托盘图标
            self.hwnd = win32gui.CreateWindow(self.classAtom, "SimplyTex", 
                win32con.WS_OVERLAPPED, 0, 0, 
                win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT,
                0, 0, self.hinst, None)
            
            hicon = win32gui.LoadIcon(0, IDI_APPLICATION)
            flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
            nid = (self.hwnd, 0, flags, WM_APP, hicon, "SimplyTex公式识别")
            win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)

        def on_app(self, hwnd, msg, wparam, lparam):
            if lparam == win32con.WM_RBUTTONUP:
                self.show_menu()
            return True

        def on_command(self, hwnd, msg, wparam, lparam):
            id = win32api.LOWORD(wparam)
            if id == 1000:
                threading.Thread(target=process_screenshot).start()
            elif id == 1001:
                os.startfile(sys.argv[0])  # 重新启动程序
            elif id == 1002:
                win32gui.DestroyWindow(hwnd)
            return True

        def on_destroy(self, hwnd, msg, wparam, lparam):
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (self.hwnd, 0))
            win32gui.PostQuitMessage(0)
            sys.exit(0)

        def show_menu(self):
            menu = win32gui.CreatePopupMenu()
            win32gui.AppendMenu(menu, win32con.MF_STRING, 1000, "立即识别 (Ctrl+Alt+Q)")
            win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
            win32gui.AppendMenu(menu, win32con.MF_STRING, 1001, "重新启动")
            win32gui.AppendMenu(menu, win32con.MF_STRING, 1002, "退出")
            pos = win32gui.GetCursorPos()
            win32gui.SetForegroundWindow(self.hwnd)
            win32gui.TrackPopupMenu(menu, 
                win32con.TPM_LEFTALIGN, 
                pos[0], pos[1], 
                0, 
                self.hwnd, None)
            win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)

    tray = TrayIcon()
    win32gui.PumpMessages()


# 将主逻辑封装成函数
def process_screenshot():
    try:
        # 新增剪贴板图片检测
        image = ImageGrab.grabclipboard()
        if not image:
            print("剪贴板中没有图片")
            return
            
        # 原有处理流程保持不变
        buffer = io.BytesIO()# 触发Windows截图
        image.save(buffer, format='PNG')
        buffer.seek(0)
        file = [("file", ("screenshot.png", buffer, "image/png"))]
        
        res = requests.post(api_url, files=file, data=data, headers=header)
        response_data = json.loads(res.text)
        latex_value = response_data['res']['latex'] # 解析 JSON 数据并提取 'latex' 的值
        pyperclip.copy(latex_value)
        print("公式已复制到剪贴板")

        res = requests.post(api_url, files=file, data=data, headers=header) # 使用requests库上传文件
        
    except Exception as e:
        print(f"错误: {str(e)}")

# 注册全局热键 (Ctrl+Alt+Q)
keyboard.add_hotkey('ctrl+alt+q', process_screenshot)

# 主循环保持程序运行
if __name__ == '__main__':
    threading.Thread(target=create_tray_icon, daemon=True).start()
    print("程序已后台运行，使用 Ctrl+Alt+Q 触发识别")
    keyboard.wait()  # 保持程序持续运行

