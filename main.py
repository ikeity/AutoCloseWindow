import json
import time
import win32gui
import win32con
import winshell
import os
import configparser
import sys


def get_executable_path():
    if getattr(sys, 'frozen', False):
        # 获取编译后的可执行文件的路径
        executable_path = sys.executable
    elif __file__:
        # 在脚本模式下获取脚本文件的路径
        executable_path = __file__
    return os.path.abspath(executable_path)


def set_autorun():
    # 获取当前可执行文件的路径
    executable_path = get_executable_path()

    # 获取启动目录
    startup_folder = winshell.startup()

    # 生成快捷方式路径
    shortcut_path = os.path.join(startup_folder, "MyApp.lnk")

    # 创建快捷方式
    with winshell.shortcut(shortcut_path) as shortcut:
        shortcut.path = executable_path
        shortcut.working_directory = os.path.dirname(executable_path)


def read_config():
    # 获取当前可执行文件所在目录
    if getattr(sys, 'frozen', False):
        # 获取编译后的可执行文件的路径
        directory = os.path.dirname(sys.executable)
    elif __file__:
        # 在脚本模式下获取脚本文件的路径
        directory = os.path.dirname(__file__)
    # 读取配置文件
    config_file = os.path.join(directory, 'config.ini')
    config = configparser.ConfigParser()
    config.read(config_file, encoding='utf-8')
    return config['AutoClose']['List'].split(',')


def get_visible_windows():
    # 获取所有可见窗口
    windows = []
    def enum_window_callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            windows.append(hwnd)
    win32gui.EnumWindows(enum_window_callback, None)
    return windows

def close_window(hwnd):
    # 关闭窗口
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

def check_and_close_windows():
    # 检查并关闭包含配置文件内字符串的窗口
    config = read_config()
    windows = get_visible_windows()
    for hwnd in windows:
        window_name = win32gui.GetWindowText(hwnd)
        for keyword in config:
            if keyword in window_name:
                close_window(hwnd)

if __name__ == "__main__":
    set_autorun()
    while True:
        check_and_close_windows()
        time.sleep(2)  # 每隔2秒检查一次