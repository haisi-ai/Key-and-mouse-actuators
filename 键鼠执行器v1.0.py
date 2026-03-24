import sys
import pyautogui
import pygetwindow as gw
import time
import threading
import json
import subprocess
import os
import psutil
import requests
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QIcon, QKeySequence
from PyQt5.QtWidgets import QApplication, QLabel, QVBoxLayout, QWidget, QPushButton, QListWidget, QLineEdit, QComboBox, \
    QHBoxLayout, QFileDialog, QMessageBox, QStatusBar, QMainWindow, QTextEdit, QDialog, QAction, \
    QTextBrowser, QMenu, QDoubleSpinBox
from pynput import keyboard


# 全局变量，用于控制操作执行状态

actions = []        # 操作列表，每个操作都包括一个描述和对应的动作

loop_interval = 1       # 控制循环的时间间隔（秒）

execution_count = 1     # 默认执行次数为 0，表示无限循环

shortcut_key = "F8"     # 默认快捷键

stop_shortcut_key = "F10"   #停止程序

is_executing = False         # 默认不执行

app_instance = None         # 用于保存 MousePositionApp 的全局实例 调用类里面的函数。

def execute_actions():
    """
    执行操作列表
    """
    global is_executing #默认不执行
    global execution_count #默认执行次数为 0，表示无限循环

    current_execution = 0  # 当前已执行次数

    if not actions:  # 检查操作列表是否为空
        if app_instance:
            app_instance.log("操作列表为空，请先添加操作~", level="error")
            is_executing = False
            app_instance.update_toggle_button_text()  # 更新按钮文本
            return

    while is_executing:
        if execution_count > 0 and current_execution >= execution_count:
            # 达到指定次数，停止执行
            if app_instance:
                app_instance.log(f"已完成指定的执行次数，共执行 {current_execution} 次", level="warning")
                time.sleep(loop_interval)
                app_instance.log("停止执行操作...", level="warning")
                is_executing = False
                app_instance.update_toggle_button_text()  # 更新按钮文本
            break

        for action in actions:
            if not is_executing:
                break
            try:
                action["action"]()  # 执行对应的动作
            except Exception as e:
                if app_instance:
                    app_instance.log(f"执行操作时发生错误: {str(e)}", level="error")
            time.sleep(loop_interval)

        current_execution += 1  # 每次循环完成后增加计数

        # 输出当前已执行次数的日志
        if app_instance:
            app_instance.log(f"当前已执行次数: {current_execution}", level="alter")

    is_executing = False  # 确保最终状态重置为停止
    app_instance.update_toggle_button_text()  # 恢复按钮状态
def toggle_execution():
    """
    切换执行状态
    """
    global is_executing # 让全局变量引用该实例

    if not actions:  # 检查操作列表是否为空
        if app_instance:
            app_instance.log("操作列表为空，请先添加操作！", level="error")
        return

    is_executing = not is_executing # 切换状态
    if is_executing:
        if app_instance:
            app_instance.log("开始执行操作...", level="warning")
        threading.Thread(target=execute_actions).start() # 创建一个线程来执行操作
    else:
        if app_instance:
            app_instance.log("停止执行操作...", level="warning")


class MousePositionApp(QMainWindow):
    # 鼠标坐标显示、时间间隔输入框、按钮、操作列表显示、操作类型选择下拉框、输入框、按钮、保存和加载按钮
    def __init__(self):
        """
        初始化鼠标坐标显示、时间间隔输入框、按钮、操作列表显示、操作类型选择下拉框、输入框、按钮、保存和加载按钮
        """
        super().__init__()

        self.current_version = "v1.0"  # 当前程序版本
        self.version_url = "https://raw.githubusercontent.com/haisi-ai/Key-and-mouse-actuators/refs/heads/main/version.txt"  # 版本文件的远程地址

        # 设置窗口标题和图标
        self.setWindowTitle("键鼠执行器")
        self.setWindowIcon(QIcon("logo.ico"))
        self.setGeometry(1300, 100, 500, 800)

        # 初始化全局变量
        global app_instance # 让全局变量引用该实例
        app_instance = self # 赋值当前实例

        global loop_interval, execution_count , shortcut_key

        # loop_interval = int(loop_interval)  # 确保是整数
        # execution_count = int(execution_count)  # 确保是整数

        # 初始化变量
        self.is_editing = False # 是否正在编辑
        self.editing_index = None # 当前正在编辑的操作的索引

        # 设置中央控件 主框架
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget) # 将主控件设置为垂直布局

        # 创建状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        # 设置状态栏样式
        self.status_bar.setStyleSheet("color: green; font-size: 16px;")  # 设置字体颜色和大小

        # 在状态栏中显示消息
        self.status_bar.showMessage("加载中...!", 5000)  # 显示 5 秒

        # 初始化菜单栏
        self.init_menu_bar()

        # 初始化主布局
        layout = QVBoxLayout()
        central_widget.setLayout(layout) # 将主控件设置为垂直布局
        self.setLayout(layout) # 将布局设置为垂直布局

        # 第一行###########################################################
        first_row_layout = QHBoxLayout() # 第一行

        # 鼠标坐标显示
        self.label = QLabel("坐标: [0, 0]", self)
        self.label.setStyleSheet("font-size: 18px;")
        self.label.setAlignment(Qt.AlignCenter)

        # 时间间隔输入框
        self.interval_input = QLineEdit(self)
        self.interval_input.setPlaceholderText("循环间隔时间 (秒)")
        self.interval_input.setText(str(loop_interval))  # 设置默认值
        self.interval_input.textChanged.connect(self.update_interval)

        # 执行次数输入框
        self.execution_count_input = QLineEdit(self)
        self.execution_count_input.setPlaceholderText("默认1 / 无限0")
        self.execution_count_input.setText(str(execution_count)) # 设置默认值
        self.execution_count_input.textChanged.connect(self.update_execution_count)

        # 按钮：启动/停止操作
        self.toggle_button = QPushButton(self)
        self.toggle_button.setStyleSheet("font-size: 16px;")
        self.update_toggle_button_text()  # 设置初始文本
        self.toggle_button.clicked.connect(self.start_toggle_stop)

        first_row_layout.addWidget(self.label)
        first_row_layout.addWidget(QLabel("间隔:", self))
        first_row_layout.addWidget(self.interval_input)
        first_row_layout.addWidget(QLabel("次数:", self))
        first_row_layout.addWidget(self.execution_count_input)
        first_row_layout.addWidget(self.toggle_button)

        # 第二行：################################操作类型选择下拉框和参数输入框，添加按钮
        second_row_layout = QHBoxLayout()
        # 操作类型选择下拉框
        self.operation_type_combo = QComboBox(self)
        self.operation_type_combo.addItems([
            "选择类型",
            "键盘输入",
            "模拟按键",
            "按快捷键",
            "鼠标点击",
            "鼠标移动",
            "鼠标拖动",
            "鼠标滚动",
            "停留时间",
            "移动窗口",
            "窗口大小",
            "启动程序",
            "关闭程序",

            # "获取像素颜色",
        ])

        # 输入框用于添加/修改操作的参数
        self.operation_input = QLineEdit(self)
        self.operation_input.setPlaceholderText("输入执行操作类型的参数...")

        # 按钮：添加操作
        self.add_button = QPushButton("添加", self)
        self.add_button.clicked.connect(self.add_or_edit_action)

        second_row_layout.addWidget(self.operation_type_combo) # 添加操作类型选择下拉框
        second_row_layout.addWidget(self.operation_input)
        second_row_layout.addWidget(self.add_button)

        # 第三行：############################## 按钮：删除、修改、清空，上移、下移
        third_row_layout = QHBoxLayout()

        # 按钮：删除、修改、清空，上移、下移
        self.delete_button = QPushButton("删除", self)
        self.delete_button.clicked.connect(self.delete_action)

        self.edit_button = QPushButton("修改", self)
        self.edit_button.clicked.connect(self.edit_action)

        self.move_up_button = QPushButton("上移", self)
        self.move_up_button.clicked.connect(self.move_up_action)

        self.move_down_button = QPushButton("下移", self)
        self.move_down_button.clicked.connect(self.move_down_action)

        third_row_layout.addWidget(self.delete_button)
        third_row_layout.addWidget(self.edit_button)
        third_row_layout.addWidget(self.move_up_button)
        third_row_layout.addWidget(self.move_down_button)

        # 操作列表显示
        self.actions_list = QListWidget(self) # 操作列表
        self.update_action_list() # 更新操作列表

        # 第四行############################ 按钮：清空  保存  加载
        save_load_layout = QHBoxLayout() # 第四行
        # 按钮：保存和加载
        self.clear_button = QPushButton("清空", self)
        self.clear_button.clicked.connect(self.clear_actions)

        self.save_button = QPushButton("保存", self)
        self.save_button.clicked.connect(self.save_actions)

        self.load_button = QPushButton("加载", self)
        self.load_button.clicked.connect(self.load_actions)
        save_load_layout.addWidget(self.clear_button) # 添加清空按钮
        save_load_layout.addWidget(self.save_button) # 添加保存和加载按钮
        save_load_layout.addWidget(self.load_button) # 添加保存和加载按钮

        # 日志显示框
        self.log_display = QTextEdit(self)
        self.log_display.setReadOnly(True)  # 设置为只读模式
        self.log_display.setMinimumSize(500, 50)
        self.log_display.setStyleSheet("background-color: #2e2e2e;color: white; font-family: Consolas; font-size: 12px;")
        self.log_display.setPlaceholderText("日志输出显示框")

        # 第五行####################################################
        five_layout = QHBoxLayout() # 第五行 目前为空

        #主窗口上下位置布局
        layout.addLayout(first_row_layout) # 添加第一行
        layout.addLayout(second_row_layout) # 添加第二行
        layout.addLayout(third_row_layout) # 添加第三行
        layout.addWidget(self.actions_list) # 添加 操作列表
        layout.addLayout(save_load_layout) # 添加第四行
        layout.addWidget(self.log_display) # 添加 日志显示
        layout.addLayout(five_layout) # 添加第五行

        # 定时器用于更新鼠标坐标
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_mouse_position)
        self.timer.start(100)

        # 注册快捷键
        self.register_shortcuts()

        # 连接操作类型下拉框的选择改变事件
        self.operation_type_combo.currentIndexChanged.connect(self.update_parameters)
    # 初始化菜单栏
    def init_menu_bar(self):
        """
        初始化菜单栏
        """
        menu_bar = self.menuBar()

        # 添加“文件”菜单
        file_menu = menu_bar.addMenu("菜单")

        # 创建并添加“设置”菜单项
        setting_action = QAction("设置", self)
        # setting_action.setShortcut(QKeySequence("Ctrl+S"))  # 设置快捷键Ctrl+S
        setting_action.triggered.connect(self.show_settings)  # 绑定“设置”菜单项的触发事件
        file_menu.addAction(setting_action) # 添加“设置”菜单项

        load_actions = QAction("打开", self)
        load_actions.triggered.connect(self.load_actions)
        file_menu.addAction(load_actions)

        save_action = QMenu("保存", self)
        # save_action.triggered.connect(self.save_actions)
        file_menu.addMenu(save_action)

        #创建保存子菜单项
        save_list_action = QAction("列表", self)
        save_list_action.triggered.connect(self.save_actions)
        save_action.addAction(save_list_action)

        save_log_action = QAction("日志", self)
        save_log_action.triggered.connect(self.save_log)
        save_action.addAction(save_log_action)


        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # 添加“文件”菜单
        tool_menu = menu_bar.addMenu("工具")

        list_all_processes_action = QAction("进程", self)
        list_all_processes_action.triggered.connect(self.list_all_processes)
        tool_menu.addAction(list_all_processes_action)

        list_all_active_windows_action = QAction("活窗", self)
        list_all_active_windows_action.triggered.connect(self.list_all_active_windows)
        tool_menu.addAction(list_all_active_windows_action)


        # 添加“帮助”菜单
        help_menu = menu_bar.addMenu("帮助")

        about_action = QAction("说明", self)
        about_action.triggered.connect(self.show_help_dialog)
        help_menu.addAction(about_action)

        about_action = QAction("更新", self)
        about_action.triggered.connect(self.show_update_dialog)
        help_menu.addAction(about_action)

        about_action = QAction("关于", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)
    # 设置窗口
    def show_settings(self):
        """
        显示设置窗口
        """
        try:
            settings_dialog = SettingsDialog(self)
            settings_dialog.exec_()
        except Exception as e:
            self.log(f"打开设置窗口时出错: {e}", level="error")
    # 快捷键
    def register_shortcuts(self):
        """
        注册全局快捷键
        """
        global shortcut_key, stop_shortcut_key  # 使用全局变量

        # 设置全局快捷键，使用动态的快捷键
        hotkeys = {
            f"<{shortcut_key.lower()}>": self.start_toggle_stop,
            f"<{stop_shortcut_key.lower()}>": self.stop_execution
        }

        listener = keyboard.GlobalHotKeys(hotkeys)
        listener.start()

        # 更新按钮文本以显示最新快捷键
        self.update_toggle_button_text()
    # 更新按钮文本
    def update_toggle_button_text(self):
        """
        更新启动/停止按钮文本，显示快捷键信息
        """
        global shortcut_key, stop_shortcut_key  # 引用全局变量
        if is_executing:
            self.toggle_button.setText(f"停止 ({shortcut_key})")
        else:
            self.toggle_button.setText(f"启动 ({shortcut_key})")

    def start_toggle_stop(self):
        """
        切换执行状态并更新按钮文本
        """
        toggle_execution()  # 切换执行状态
        self.update_toggle_button_text()  # 更新按钮文本

    # 停止窗口
    def stop_execution(self):
        """
        停止操作
        """
        global is_executing # 引用全局变量
        is_executing = False # 停止执行
        self.toggle_button.setText("启动") # 恢复按钮状态
    # 帮助窗口
    def show_help_dialog(self):
        """
        显示非模态说明窗口
        """
        # 检查是否需要重新创建窗口
        if not hasattr(self, 'help_dialog') or self.help_dialog is None:
            self.help_dialog = HelpDialog(self)
            self.help_dialog.setWindowModality(Qt.NonModal)  # 设置为非模态
            self.help_dialog.setAttribute(Qt.WA_DeleteOnClose)  # 窗口关闭时释放资源
        elif not self.help_dialog.isVisible():
            # 如果窗口已关闭但对象仍在，则重新显示
            self.help_dialog = HelpDialog(self)

        # 显示窗口
        self.help_dialog.show()
        self.help_dialog.raise_()  # 确保窗口显示在最前面
        self.help_dialog.activateWindow()
    # 关于窗口
    def show_about_dialog(self, event=None):  # 去掉 event 或设置为可选参数
        QMessageBox.about(
            self,
            "关于",
            f"键鼠执行器：{self.current_version}\n"
            "作者：海斯\n"
            "邮箱：haisi@mail.com"
        )
    # 检查更新功能
    def show_update_dialog(self):
        """
        检查更新功能
        """
        try:
            # 获取远程版本
            response = requests.get(self.version_url)
            response.raise_for_status()

            remote_version = response.text.strip()

            if remote_version > self.current_version:
                # 获取更新日志
                changelog_url = "https://raw.githubusercontent.com/Haisi-1536/-/refs/heads/main/changelog.txt"
                changelog_response = requests.get(changelog_url)
                changelog_response.raise_for_status()

                changelog = changelog_response.text.strip()

            # 创建富文本消息框
                message_box = QMessageBox(self)
                message_box.setWindowTitle("检查更新")
                message_box.setTextFormat(Qt.RichText)  # 启用富文本格式
                message_box.setText(
                    f"发现新版本: <b>{remote_version}</b>！<br><br>"
                    f"<b>更新内容:</b><br>{changelog}<br><br>"
                    f"请前往 <a href='https://github.com/Haisi-1536/-'>官网下载更新</a>。"
                )
                message_box.setStandardButtons(QMessageBox.Ok)
                message_box.exec_()

            else:
                QMessageBox.information(
                    self,
                    "检查更新",
                    "当前已是最新版本！"
                )
        except requests.RequestException as e:
            QMessageBox.warning(
                self,
                "检查更新",
                f"无法连接到更新服务器: {e}"
            )
        except Exception as e:
            QMessageBox.warning(
                self,
                "检查更新",
                f"更新检查失败: {str(e)}"
            )
    # 下载更新文件
    def download_update(self, download_url, save_path):
        """
        下载更新文件
        """
        try:
            response = requests.get(download_url, stream=True)
            with open(save_path, "wb") as file:
                for chunk in response.iter_content(chunk_size=1024):
                    file.write(chunk)
            QMessageBox.information(self, "更新成功", "更新文件已下载！")
        except Exception as e:
            QMessageBox.warning(self, "下载失败", f"更新文件下载失败: {str(e)}")
    # 更新操作列表
    def update_action_list(self):
        """
        更新操作列表显示
        """
        current_row = self.actions_list.currentRow() # 保存当前选中行
        self.actions_list.clear() # 清空列表
        self.actions_list.addItems([action["description"] for action in actions]) # 添加操作描述
        if 0 <= current_row < len(actions): # 如果当前选中行有效，则选中它
            self.actions_list.setCurrentRow(current_row) # 设置选中行
    # 更新鼠标坐标显示
    def update_mouse_position(self):
        """
        更新鼠标坐标显示
        """
        x, y = pyautogui.position()
        self.label.setText(f"坐标: [{x}, {y}]")
    # 更新循环时间间隔
    def update_interval(self):
        """
        更新循环时间间隔
        """
        global loop_interval
        try:
            value = float(self.interval_input.text())
            if value > 0:
                loop_interval = value
                self.log(f"循环时间间隔已更新为: {loop_interval} 秒",level="alter")
            else:
                self.log("时间间隔必须大于 0 秒",level="warning")
        except ValueError:
            self.log("请输入有效的数值",level="error")
    # 更新执行次数
    def update_execution_count(self):
        """
        更新执行次数
        """
        global execution_count
        try:
            value = int(self.execution_count_input.text())
            if value > 0:
                execution_count = value
                self.log(f"执行次数已更新为: {execution_count} 次",level = "alter")
            else:
                execution_count = 0 # 0 表示无限循环
                self.log("执行次数已设置为无限",level = "alter")
        except ValueError:
            self.log("请输入有效的整数",level = "error")
    # 根据选择的操作类型和参数创建一个操作对象
    def create_action(self, selected_type, parameter):
        """
        根据选择的操作类型和参数创建一个操作对象
        """
        try:
            # 映射操作类型到对应的处理函数
            action_map = {
                "键盘输入": self.action_keyboard_input,
                "模拟按键": self.action_simulate_key,
                "按快捷键": self.action_hotkey,
                "鼠标点击": self.action_mouse_click,
                "鼠标移动": self.action_mouse_move,
                "鼠标拖动": self.action_mouse_drag,
                "鼠标滚动": self.action_mouse_scroll,
                "停留时间": self.action_wait,
                "移动窗口": self.action_move_window,
                "窗口大小": self.action_resize_window,
                "启动程序": self.action_start_program_wrapper,
                "关闭程序": self.action_close_program_wrapper,
            }

        # 检查类型是否支持
            if selected_type not in action_map:
                raise ValueError(f"不支持的操作类型: {selected_type}")

            # 调用对应的处理函数
            action = action_map[selected_type](parameter)
            if not action:
                raise ValueError("创建操作失败，参数不正确或为空！")
            return action

        except ValueError as e:
            # 不记录非关键错误日志
            return None

        except Exception as e:
            # 捕获错误并输出日志
            self.log(f"创建操作失败: {e}", level="warning")
            return None

    # ====================各操作的独立方法============================#
    # 键盘输入操作
    def action_keyboard_input(self, parameter):
        """
        处理键盘输入操作
        """
        if parameter:
            return {
                "description": f"键盘输入: {parameter}",
                "action": lambda: pyautogui.write(parameter, interval=0.5)
            }
    # 模拟按键操作
    def action_simulate_key(self, parameter):
        """
        处理模拟按键操作
        """
        if parameter:
            parts = parameter.split(",")
            key = parts[0].strip()
            presses = int(parts[1].strip()) if len(parts) > 1 else 1
            interval = float(parts[2].strip()) if len(parts) > 2 else 0.0
            return {
                "description": f"模拟按键: {key},{presses},{interval}",
                "action": lambda: pyautogui.press(key, presses=presses, interval=interval)
            }
    # 快捷键操作
    def action_hotkey(self, parameter):
        """
        处理快捷键操作
        """
        if parameter:
            keys = parameter.lower().split("+")
            return {
                "description": f"按快捷键: {parameter}",
                "action": lambda: pyautogui.hotkey(*keys)
            }
    # 鼠标点击操作
    def action_mouse_click(self, parameter):
        """
        处理鼠标点击操作
        """
        if parameter:
            parts = parameter.split(",")
            button = parts[0].strip().lower()
            clicks = int(parts[1].strip()) if len(parts) > 1 else 1
            duration = float(parts[2].strip()) if len(parts) > 2 else 0.0
            return {
                "description": f"鼠标点击: {button}, {clicks}, {duration}",
                "action": lambda: pyautogui.click(button=button, clicks=clicks, duration=duration)
            }
    # 鼠标移动操作
    def action_mouse_move(self, parameter):
        """
        处理鼠标移动操作
        """
        if parameter:
            x, y = map(int, parameter.split(","))
            return {
                "description": f"鼠标移动: {x}, {y}",
                "action": lambda: pyautogui.moveTo(x, y)
            }
    # 鼠标拖动操作
    def action_mouse_drag(self, parameter):
        """
        处理鼠标拖动操作
        """
        if parameter:
            parts = parameter.split(",")
            x, y = map(int, parts[:2])
            duration = float(parts[2].strip()) if len(parts) > 2 else 0.0
            return {
                "description": f"鼠标拖动: {x}, {y}, {duration}",
                "action": lambda: pyautogui.dragTo(x, y, duration=duration)
            }
    # 鼠标滚动操作
    def action_mouse_scroll(self, parameter):
        """
        处理鼠标滚动操作
        """
        if parameter:
            scroll_amount = int(parameter)
            return {
                "description": f"鼠标滚动: {scroll_amount}",
                "action": lambda: pyautogui.scroll(scroll_amount)
            }
    # 停留时间操作
    def action_wait(self, parameter):
        """
        处理停留时间操作
        """
        if parameter:
            delay_time = float(parameter)
            return {
                "description": f"停留时间: {delay_time}",
                "action": lambda: time.sleep(delay_time)
            }
    # 窗口位置操作
    def action_move_window(self, parameter):
        """
        处理移动窗口操作
        """
        if parameter:
            x, y = map(int, parameter.split(","))
            return {
                "description": f"移动窗口: {x}, {y}",
                "action": lambda: self.move_window(x, y)
            }
    # 移动窗口位置的独立方法
    def move_window(self, x, y):
        """
        执行移动当前活动窗口到指定位置
        """
        try:
            # 获取当前活动窗口
            active_window = gw.getActiveWindow()
            if active_window:
                active_window.moveTo(x, y)
                self.log(f"当前窗口已移动到 ({x}, {y})", level = "info")
            else:
                self.log("未找到当前活动窗口", level = "warning")
        except Exception as e:
            self.log(f"移动窗口失败: {e}", level = "error")
    # 调整窗口大小的独立方法
    def action_resize_window(self, parameter):
        """
        处理调整窗口大小操作
        """
        if parameter:
            parts = parameter.split(",")
            if len(parts) == 2:
                width, height = map(int, parts)
                return {
                    "description": f"窗口大小: {width},{height}",
                    "action": lambda: self.resize_window(width, height)
                }
    # 调整窗口大小的包装方法
    def resize_window(self, width, height):
        """
        执行调整窗口大小操作
        """
        try:
            # 获取当前活动窗口
            active_window = gw.getActiveWindow()
            if active_window:
                active_window.resizeTo(width, height)
                self.log(f"当前窗口已调整大小为 {width}x{height}", level = "info")
            else:
                self.log("未找到当前活动窗口", level = "warning")
        except Exception as e:
            self.log(f"调整窗口大小失败: {e}", level = "error")
    # 启动程序的包装方法
    def action_start_program_wrapper(self, program_path):
        """
        包装启动程序的操作对象生成逻辑
        """
        if program_path:
            return {
                "description": f"启动程序: {program_path}",
                "action": lambda: self.action_start_program(program_path)
            }

    # 启动程序函数
    def action_start_program(self, program_path):
        """
        启动程序并打印其PID和进程名称。若程序已启动，则不再重新启动，而是直接返回PID和进程名称。
        """
        try:
            # 获取程序的名称（假设程序路径包含名称），去除扩展名
            program_name = os.path.splitext(os.path.basename(program_path))[0].lower() # 将文件名转换为小写

            for proc in psutil.process_iter(['pid', 'name']):
                if program_name in proc.info['name'].lower():
                    self.log(f"程序 '{program_path}' 已经在运行！", level="warning")
                    self.log(f"PID: {proc.info['pid']}, 进程名称: {proc.info['name']}", level="warning")
                    return  # 程序已在运行，直接返回

            # 如果未找到进程，则启动程序
            process = subprocess.Popen(program_path)

            # 等待一会儿，确保程序启动
            time.sleep(2)

            # 获取进程的PID
            pid = process.pid

            # 使用psutil获取进程的名称
            process_info = psutil.Process(pid)
            process_name = process_info.name()

            self.log(f"程序 '{program_path}' 启动成功！", level="Hint")
            self.log(f"PID: {pid}, 进程名称: {process_name}", level="Hint")

        except psutil.NoSuchProcess:
            self.log(f"启动程序失败: 找不到进程", level="error")
        except psutil.AccessDenied:
            self.log(f"启动程序失败: 没有权限访问进程", level="error")
        except Exception as e:
            self.log(f"启动程序失败: {e}", level="error")

    # 关闭程序的包装方法
    def action_close_program_wrapper(self, process_name):
        """
        包装关闭程序的操作对象生成逻辑
        """
        if process_name:
            return {
                "description": f"关闭程序: {process_name}",
                "action": lambda: self.action_close_program(process_name)
            }

    # 关闭程序函数
    def action_close_program(self, process_name):
        """
        关闭指定名称的程序，如果程序正在运行。
        """
        try:
            # 遍历所有进程，查找匹配的程序
            for proc in psutil.process_iter(['pid', 'name']):
                if process_name.lower() in proc.info['name'].lower():
                    # 获取匹配进程的PID并关闭它
                    os.kill(proc.info['pid'], 9)
                    self.log(f"程序 '{process_name}' 已成功关闭！", level="Hint")
                    return  # 找到并关闭后直接返回
            self.log(f"未找到匹配的程序 '{process_name}'", level="warning")
        except Exception as e:
            self.log(f"关闭程序失败: {e}", level="error")
#=======================================================================#
    # 列出所有进程
    def list_all_processes(self):
        """
        获取并显示系统当前所有进程的名称和PID。
        """
        try:
            # 获取当前所有的进程
            processes = []
            for proc in psutil.process_iter(['pid', 'name']):
                processes.append(proc.info)
            # 打印所有进程信息
            if processes:
                self.log(f"{'PID':<10} {'Process Name'}", level = "Hint")
                self.log("-" * 30, level = "Hint")
                for process in processes:
                    self.log(f"{process['pid']:<10} {process['name']}", level = "processes")
            else:
                self.log("没有找到任何进程。", level = "warning")
        except Exception as e:
            self.log(f"获取进程信息失败: {e}", level = "error")
    # 列出所有活动的窗口
    def list_all_active_windows(self):
        """
        获取并显示所有当前活动窗口的标题和状态。
        """
        try:
            # 获取所有窗口
            windows = gw.getWindowsWithTitle('')

            # 如果有窗口，打印窗口标题和状态
            if windows:
                self.log(f"{'Window Title':<50} {'Status'}", level = "Hint")
                self.log("-" * 60, level = "Hint")
                for window in windows:
                    status = "Active" if window.isActive else "Inactive"
                    self.log(f"{window.title:<50} {status}", level = "processes")
            else:
                self.log("没有找到活动窗口。", level = "warning")
        except Exception as e:
            self.log(f"获取窗口信息失败: {e}", level = "error")



    # 新增添加操作
    def add_or_edit_action(self):
        """
        新增添加操作
        """
        selected_type = self.operation_type_combo.currentText()  # 获取当前选择的操作类型
        parameter = self.operation_input.text()  # 获取操作参数

        # 创建新操作
        action = self.create_action(selected_type, parameter)  # 创建操作对象
        if action is None:
            self.log("创建操作失败，请检查参数！", level="error")
            return  # 如果创建失败，直接退出

        else:
            # 新增模式
            actions.append(action)  # 添加到操作列表
            self.log(f"操作已添加: {action['description']}", level="Hint")

        # 更新界面
        self.operation_input.clear()  # 清空输入框
        self.update_action_list()  # 更新操作列表
    # 编辑操作
    def edit_action(self):
        """
        进入或退出修改模式
        """
        selected_row = self.actions_list.currentRow()  # 获取当前选中的行号

        # 检查是否选择了操作
        if selected_row < 0:
            QMessageBox.warning(self, "错误", "请先选择要修改的操作！")
            return

        # 如果当前不在修改模式，则进入修改模式
        if not self.is_editing:
            # 进入修改模式
            self.is_editing = True
            self.editing_index = selected_row

            # 禁用添加按钮，防止误操作
            self.add_button.setEnabled(False)
            self.clear_button.setEnabled(False)
            self.edit_button.setText("确认")  # 修改按钮显示为确认

            # 填充当前选中项的参数到输入框
            self._fill_action_to_inputs(selected_row)

            # 绑定实时更新逻辑
            self._bind_update_logic()

            self.log(f"进入修改模式: 第 {selected_row + 1} 行", level="Hint")
            return

        # 如果当前在修改模式，则退出修改模式
        else:
            # 验证当前修改的参数
            current_type = self.operation_type_combo.currentText()
            current_parameter = self.operation_input.text()

            if not current_type or not current_parameter:
                QMessageBox.warning(self, "错误", "当前修改的参数无效，请检查输入！")
                return

            # 更新操作列表并退出修改模式
            try:
                updated_action = self.create_action(current_type, current_parameter)
                if updated_action is not None:
                    actions[self.editing_index] = updated_action  # 更新操作
                    self.update_action_list()  # 刷新操作列表显示
            except Exception as e:
                QMessageBox.warning(self, "错误", f"更新操作失败: {e}")
                return

            # 恢复初始状态
            self.is_editing = False
            self.editing_index = None
            self.edit_button.setText("修改")  # 恢复按钮状态
            self.add_button.setEnabled(True)  # 启用添加按钮
            self.clear_button.setEnabled(True) # 启用清空按钮
            self._unbind_update_logic()  # 解绑实时更新逻辑

            self.log("已退出修改模式", level="Hint")
    # 修改模式时 实时更新类型及参数
    def _bind_update_logic(self):
        """
        绑定实时更新和即时预填充逻辑
        """
        # 延迟更新定时器
        self.update_timer = QTimer(self)
        self.update_timer.setSingleShot(True)  # 确保只触发一次
        self.update_timer.timeout.connect(self._perform_update_action)

        def handle_input_change():
            """
            延迟触发实时更新
            """
            if not self.is_editing or self.editing_index is None:
                return
            self.update_timer.start(2000)  # 2000 毫秒后执行更新

        def handle_selection_change():
            """
            用户切换选择时立即填充操作类型和参数
            """
            if not self.is_editing or self.editing_index is None:
                return

            selected_row = self.actions_list.currentRow()
            if selected_row >= 0:
                self.editing_index = selected_row
                self._fill_action_to_inputs(selected_row)  # 即时填充

        # 绑定信号
        self.operation_input.textChanged.connect(handle_input_change)
        self.operation_type_combo.currentTextChanged.connect(handle_input_change)
        self.actions_list.itemSelectionChanged.connect(handle_selection_change)
    # 实时更新
    def _perform_update_action(self):
        """
        执行实时更新操作，保持输入框和操作列表内容一致
        """
        if not self.is_editing or self.editing_index is None:
            return

        # 获取用户当前输入的类型和参数
        new_type = self.operation_type_combo.currentText()
        new_parameter = self.operation_input.text().strip()
        cursor_position = self.operation_input.cursorPosition()

        try:
            # 创建新的操作对象
            updated_action = self.create_action(new_type, new_parameter)
            if updated_action:
                # 仅当内容有变更时更新
                if actions[self.editing_index] != updated_action:
                    actions[self.editing_index] = updated_action
                    self.update_action_list()

                # 恢复光标位置
                self.operation_input.setCursorPosition(cursor_position)
        except Exception:
            pass  # 忽略异常，保持输入框内容不变
    # 选择变化
    def _on_action_selection_changed(self):
        """
        当操作列表选择变化时，实时更新输入框内容
        """
        if not self.is_editing:  # 如果不在修改模式，忽略选择变化
            return

        selected_row = self.actions_list.currentRow()
        if selected_row >= 0:
            self.editing_index = selected_row  # 更新当前修改的索引
            self._fill_action_to_inputs(selected_row)  # 填充新选中的操作
    # 解绑逻辑
    def _unbind_update_logic(self):
        """
        解绑实时更新逻辑
        """
        try:
            self.operation_type_combo.currentTextChanged.disconnect()
            self.operation_input.textChanged.disconnect()
            self.actions_list.itemSelectionChanged.disconnect()
        except TypeError:
            pass  # 信号未绑定时忽略错误
    # 填充操作类型和参数
    def _fill_action_to_inputs(self, row):
        """
        根据选中行号填充操作类型和参数到输入框
        """
        if row < 0 or row >= len(actions):
            return

        selected_action = actions[row]
        description_parts = selected_action["description"].split(": ", 1)

        # 分离操作类型和参数
        operation_type = description_parts[0] if len(description_parts) > 0 else ""
        parameter = description_parts[1] if len(description_parts) > 1 else ""

        # 暂时屏蔽信号，防止触发更新逻辑
        self.operation_type_combo.blockSignals(True)
        self.operation_input.blockSignals(True)

        # 填充内容
        self.operation_type_combo.setCurrentText(operation_type)
        self.operation_input.setText(parameter)

        # 恢复信号
        self.operation_type_combo.blockSignals(False)
        self.operation_input.blockSignals(False)
    # 删除已选操作
    def delete_action(self):
        """
        删除已选操作
        """
        selected_row = self.actions_list.currentRow() # 获取当前选中的行号
        if selected_row >= 0:
            del actions[selected_row]
            self.update_action_list()
        else:
            self.log("请先选择要删除的操作。", level = "warning")
    # 清空操作列表
    def clear_actions(self):
        """
        清空操作列表
        """
        global actions
        actions.clear()  # 清空操作列表
        self.update_action_list()  # 更新界面上的操作列表
        self.log("操作列表已清空", level = "Hint")
    # 向上移动选中的操作
    def move_up_action(self):
        """
        向上移动选中的操作
        """
        selected_row = self.actions_list.currentRow()
        if selected_row > 0:
            actions[selected_row], actions[selected_row - 1] = actions[selected_row - 1], actions[selected_row]
            self.update_action_list()
            self.actions_list.setCurrentRow(selected_row - 1)
    # 向下移动选中的操作
    def move_down_action(self):
        """
        向下移动选中的操作
        """
        selected_row = self.actions_list.currentRow()
        if selected_row < len(actions) - 1:
            actions[selected_row], actions[selected_row + 1] = actions[selected_row + 1], actions[selected_row]
            self.update_action_list()
            self.actions_list.setCurrentRow(selected_row + 1)
    # 保存操作列表
    def save_actions(self):
        """
        保存操作列表
        """
        filename, _ = QFileDialog.getSaveFileName(self, "保存操作", "", "JSON 文件 (*.json)")
        if filename:
            try:
                with open(filename, "w") as f:
                    json.dump([{"description": a["description"], "action": None} for a in actions], f)
                self.log(f"操作列表已保存到 {filename}", level = "Hint")
            except Exception as e:
                self.log(f"保存失败: {e}", level = "error")
    # 加载操作列表
    def load_actions(self):
        """
        加载操作列表
        """
        filename, _ = QFileDialog.getOpenFileName(self, "加载操作", "", "JSON 文件 (*.json)")
        if filename:
            try:
                with open(filename, "r") as f:
                    loaded_actions = json.load(f)
                    for action in loaded_actions:
                        self.add_action_from_description(action["description"])
                self.log(f"操作列表已从 {filename} 加载", level = "Hint")
            except Exception as e:
                self.log(f"加载失败: {e}", level = "error")
    # 重构操作列表
    def add_action_from_description(self, description):
        """
        根据描述重新创建动作
        """
        if description.startswith("键盘输入"):
            # 提取参数并重新创建动作
            parameter = description.split(": ")[1]
            actions.append({
                "description": description,
                "action": lambda: pyautogui.typewrite(parameter)
            })

        elif description.startswith("模拟按键"):
            # 提取按键模拟的参数
            try:
                _, param_str = description.split(": ")
                actions.append(self.action_simulate_key(param_str.strip()))
            except Exception as e:
                self.log(f"加载模拟按键失败: 无效参数 - {description}", level="error")

        elif description.startswith("按快捷键"):
            # 提取快捷键组合
            parameter = description.split(": ")[1]
            keys = parameter.split("+")
            actions.append({
                "description": description,
                "action": lambda: pyautogui.hotkey(*keys)
            })

        elif description.startswith("鼠标点击"):
            # 提取鼠标点击参数
            try:
                # 假设 description 格式为 "鼠标点击: 按键, 次数, 时间"
                _, param_str = description.split(": ")
                params = param_str.split(",")
                button = params[0].strip().lower()  # 按键类型：left/right/middle
                clicks = int(params[1].strip()) if len(params) > 1 else 1  # 点击次数，默认 1
                duration = float(params[2].strip()) if len(params) > 2 else 0.0  # 持续时间，默认 0.0
                actions.append({
                    "description": description,
                    "action": lambda: pyautogui.click(button=button, clicks=clicks, duration=duration)
                })
            except (ValueError, IndexError) as e:
                self.log(f"加载鼠标点击失败: 无效参数 - {description}", level = "error")

        elif description.startswith("鼠标移动"):
            try:
                # 假设 description 格式为 "鼠标移动: x, y"
                _, param_str = description.split(": ")
                x, y = map(int, param_str.split(","))
                actions.append({
                    "description": description,
                    "action": lambda: pyautogui.moveTo(x, y)
                })
            except (ValueError, IndexError):
                self.log(f"加载鼠标移动失败: 无效参数 - {description}", level = "error")

        elif description.startswith("鼠标拖动"):
            try:
                # 假设 description 格式为 "鼠标拖动: x, y, duration"
                _, param_str = description.split(": ")
                parts = param_str.split(",")
                x, y = map(int, parts[:2])  # 获取 x 和 y 坐标
                duration = float(parts[2].strip()) if len(parts) > 2 else 2.0  # 获取时间参数，默认为 0.0
                actions.append({
                    "description": description,
                    "action": lambda: pyautogui.dragTo(x, y, duration=duration)
                })
            except (ValueError, IndexError):
                self.log(f"加载鼠标拖动失败: 无效参数 - {description}", level="error")

        elif description.startswith("鼠标滚动"):
            try:
                # 假设 description 格式为 "鼠标滚轮: 步数"
                _, steps_str = description.split(": ")
                steps = int(steps_str.strip())  # 包括负数的转换
                actions.append({
                    "description": description,
                    "action": lambda: pyautogui.scroll(steps)
                })
            except (ValueError, IndexError) as e:
                self.log(f"加载鼠标滚轮失败: 无效参数 - {description}", level = "error")

        elif description.startswith("停留时间"):
            try:
                # 假设 description 格式为 "停留时间: x" (单位为秒)
                _, param_str = description.split(": ")
                delay_time = float(param_str)
                actions.append({
                    "description": description,
                    "action": lambda: time.sleep(delay_time)
                })
            except (ValueError, IndexError):
                self.log(f"加载停留时间失败: 无效参数 - {description}", level = "error")

        elif description.startswith("移动窗口"):
            # 处理窗口移动操作
            try:
                _, param_str = description.split(": ")
                actions.append(self.action_move_window(param_str.strip()))
            except Exception as e:
                self.log(f"加载移动窗口失败: 无效参数 - {description}", level="error")

        elif description.startswith("窗口大小"):
            # 处理调整窗口大小操作
            try:
                _, param_str = description.split(": ")
                actions.append(self.action_resize_window(param_str.strip()))
            except Exception as e:
                self.log(f"加载调整窗口大小失败: 无效参数 - {description}", level="error")

        elif description.startswith("启动程序"):
            _, param_str = description.split(": ")
            program_path = param_str.strip()

            # 使用包装函数创建操作对象
            actions.append({
                "description": description,
                "action": lambda: self.action_start_program_wrapper(program_path)["action"]()
            })

        elif description.startswith("关闭程序"):
            _, param_str = description.split(": ")
            process_name = param_str.strip()

            # 使用包装函数创建操作对象
            actions.append({
                "description": description,
                "action": lambda: self.action_close_program_wrapper(process_name)["action"]()
            })
        # 更新操作列表显示
        self.update_action_list()
    # 参数提示
    def update_parameters(self):
        """
        根据当前选择的操作类型更新输入框的提示信息
        """
        selected_type = self.operation_type_combo.currentText()
        if selected_type in ["键盘输入"]:
            self.operation_input.setPlaceholderText("输入拼音，英文。")
        elif selected_type in ["模拟按键"]:
            self.operation_input.setPlaceholderText("例如:a,2,3  /a按2次3秒间隔")
        elif selected_type in ["按快捷键"]:
            self.operation_input.setPlaceholderText("例如:ctrl+s,alt+a,space,win")
        elif selected_type in ["鼠标点击"]:
            self.operation_input.setText("left")
            self.operation_input.setPlaceholderText("例如:right,2,0.5。左右中:left/right/middle,点击次数,时间")
        elif selected_type in ["鼠标移动"]:
            self.operation_input.setPlaceholderText("例如:200, 200，1  /x,y,时间")
        elif selected_type in ["鼠标拖动"]:
            self.operation_input.setPlaceholderText("示例: 200,200,1 /拖动到200,200用时1秒" )
        elif selected_type in ["鼠标滚动"]:
            self.operation_input.setPlaceholderText("示例: 200 / 上滚 200 下滚 -200" )
        elif selected_type in ["移动窗口"]:
            self.operation_input.setPlaceholderText("100,100  / 程序左上角位置锁定要x,y 200处")
        elif selected_type in ["窗口大小"]:
            self.operation_input.setPlaceholderText("600,800 /调整当前活动窗口")
        elif selected_type in ["启动程序"]:
            self.operation_input.setPlaceholderText("例如: notepad.exe /或完整路径 C:\\Path\\to\\app.exe")
        elif selected_type in ["关闭程序"]:
            self.operation_input.setPlaceholderText("例如: notepad.exe /或 PID")
        elif selected_type in ["停留时间"]:
            self.operation_input.setPlaceholderText("请输入停留时间（秒）")
        elif selected_type in ["选择类型"]:
            self.operation_input.setPlaceholderText("参数输入框~~")
        else:
            self.operation_input.setPlaceholderText("错误的选择！")
    # 日志
    def log(self, message, level="info"):
        """
        在日志显示框中添加一条消息，支持不同颜色
        :param message: 日志内容
        :param level: 日志级别（info, warning, error）

            "info": "green", # 默认颜色为绿色
            "warning": "yellow", # 警告颜色为黄色
            "error": "red", # 错误颜色为红色
            "Hint": "cyan" ,# 提示颜色为青色
            "processes": "orange" # # 进程颜色为橙色
        """
        # 获取时间戳
        timestamp = time.strftime("[%Y-%m-%d %H:%M:%S]")

        # 根据日志级别设置颜色
        color_map = {
            "info": "green", # 默认颜色为绿色
            "warning": "yellow", # 警告颜色为黄色
            "error": "red", # 错误颜色为红色
            "Hint": "cyan" ,# 提示颜色为青色
            "processes": "orange" # # 进程颜色为橙色
        }
        color = color_map.get(level, "cyan")  # 默认颜色为白色white /cyan青色 /green绿 /gray灰色 /orange橙色

        # 格式化消息
        formatted_message = f'<font color="{color}">{timestamp} {message}</font>'
        self.log_display.append(formatted_message)
    #保存日志
    def save_log(self):
        """
        保存日志到本地文件
        """
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_path, _ = QFileDialog.getSaveFileName(self,
                                                   "保存日志",
                                                   os.getcwd(),
                                                   "日志文件 (*.txt);;所有文件 (*.*)",
                                                   options=options)
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as file:
                    # 获取日志内容（纯文本格式）
                    log_text = self.log_display.toPlainText()
                    file.write(log_text)
                self.log(f"日志已保存到 {file_path}", level="Hint")
            except Exception as e:
                self.log(f"保存日志失败: {e}", level="error")

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        global app_instance

        self.setWindowTitle("设置")
        self.setMinimumSize(400, 300)

        # 主布局
        layout = QVBoxLayout(self)

        # 修改快捷键
        shortcut_layout = QHBoxLayout()
        shortcut_label = QLabel("启动/停止快捷键: ")
        self.shortcut_input = QLineEdit(shortcut_key)  # 默认快捷键
        # self.shortcut_input.setEnabled(False)  # 禁用快捷键输入框
        shortcut_layout.addWidget(shortcut_label)
        shortcut_layout.addWidget(self.shortcut_input)
        layout.addLayout(shortcut_layout)

        # 时间间隔设置
        interval_layout = QHBoxLayout()
        interval_label = QLabel("默认时间间隔:  ")
        self.interval_spinbox = QDoubleSpinBox()
        self.interval_spinbox.setRange(0.1, 60.0)  # 范围从0.1秒到60秒
        self.interval_spinbox.setValue(loop_interval)  # 设置默认时间间隔
        self.interval_spinbox.setSingleStep(0.1)  # 设置步进值为0.1秒
        interval_layout.addWidget(interval_label)
        interval_layout.addWidget(self.interval_spinbox)
        layout.addLayout(interval_layout)

        # 循环次数设置
        count_layout = QHBoxLayout()
        count_label = QLabel("默认执行次数:    ")
        self.execution_count_input = QLineEdit(str(execution_count))  # 0表示无限执行
        count_layout.addWidget(count_label)
        count_layout.addWidget(self.execution_count_input)
        layout.addLayout(count_layout)

        # 按钮区
        button_layout = QHBoxLayout()
        self.save_button = QPushButton("确认")
        self.save_button.clicked.connect(self.save_settings)
        self.apply_button = QPushButton("应用")
        self.apply_button.clicked.connect(self.apply_settings)
        self.reset_button = QPushButton("恢复默认")
        self.reset_button.clicked.connect(self.reset_settings)
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.apply_button)
        button_layout.addWidget(self.reset_button)
        layout.addLayout(button_layout)

    def save_settings(self):
        """
        保存设置并直接写入主程序输入框
        """
        global shortcut_key, loop_interval, execution_count, app_instance

        self.save_button.clicked.disconnect()  # 先断开旧信号（如果已绑定）
        self.save_button.clicked.connect(self.save_settings)  # 再绑定新信号

        try:
            # 获取控件中的值
            shortcut_key = self.shortcut_input.text().strip()
            loop_interval = self.interval_spinbox.value()
            execution_count = int(self.execution_count_input.text())

            # 验证输入
            if not shortcut_key:
                raise ValueError("快捷键不能为空！")
            if loop_interval <= 0:
                raise ValueError("时间间隔必须大于0！")
            if execution_count < 0:
                raise ValueError("执行次数不能为负数！")

            # 直接写入主程序的输入框
            if hasattr(self.parent(), "interval_input"):
                self.parent().interval_input.setText(str(loop_interval))  # 写入时间间隔
            if hasattr(self.parent(), "execution_count_input"):
                self.parent().execution_count_input.setText(str(execution_count))  # 写入执行次数
            if hasattr(self.parent(), "register_shortcuts"):
                self.parent().register_shortcuts()  # 重新注册快捷键
            app_instance.log("设置已保存并更新！")
            self.close()  # 关闭设置窗口
        except ValueError as e:
            QMessageBox.warning(self, "输入错误", str(e))
        except Exception as e:
            QMessageBox.warning(self, "错误", f"保存设置失败: {str(e)}")
    def apply_settings(self):
        """
        应用设置
        """
        self.save_settings()
        app_instance.log("设置已应用！")

    def reset_settings(self):
        """
        恢复默认设置
        """
        self.shortcut_input.setText("F8")
        self.interval_spinbox.setValue(1)
        self.execution_count_input.setText("1")
        app_instance.log("设置已恢复默认值！")

class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("说明")
        self.setMinimumSize(1000, 500)

        # 布局
        layout = QVBoxLayout(self)

        # 使用 QTextBrowser 显示说明文本
        text_browser = QTextBrowser(self)
        text_browser.setOpenExternalLinks(True)  # 支持点击超链接
        text_browser.setHtml(self.get_help_text())  # 设置 HTML 内容

        # 添加关闭按钮
        close_button = QPushButton("关闭", self)
        close_button.clicked.connect(self.close)  # 点击按钮关闭窗口

        # 布局管理
        layout.addWidget(text_browser)
        layout.addWidget(close_button)

    def closeEvent(self, event):
        """
        当窗口关闭时，将父窗口中的 help_dialog 置为 None
        """
        if self.parent():
            self.parent().help_dialog = None
        super().closeEvent(event)

    def get_help_text(self):
        """
        返回帮助文本的 HTML 内容
        """
        return """
        <html>
            <body>
                <h1 style="color: blue; font-family: Arial; font-size: 24px; text-align: center;"><b>使用说明<b></h1>
                <p style="color: green; font-size: 16px;"><b>此程序用于模拟键鼠操作，以下是一些功能使用说明：<b></p>
                <hr>
                <ul>
                    <li><span style="color: green;"><b>坐标：<b></span> 当前鼠标实时坐标 x,y 。用于参考 编辑鼠标移动操作时输入对应参数。</li>
                    <li><span style="color: green;"><b>时隔：<b></span> 参数执行每个操作间隔时间 0.1秒就是10倍，1秒是默认倍数，数字越大执行越慢。</li>
                    <li><span style="color: green;"><b>次数：<b></span> 执行次数 默认1次， 0是无限循环执行。</li>
                    <li><span style="color: green;"><b>选择类型：<b></span>选择操作类型。</li>
                    <li><span style="color: green;"><b>输入操作参数：<b></span>选择不同的类型后，输入相应的参数</li>
                    <li><span style="color: green;"><b>键盘输入：<b></span>示例：hehehe 空格 //输出结果就是呵呵呵，不加空格就 加shift或切换英文输入 就是hehehe。</li>
                    <li><span style="color: green;"><b>鼠标点击：<b></span>示例：right,2,0.5  //按下鼠标 左/右/中：left/right/middle, 点击次数, 时间, 注意是英文逗号。</li>
                    <li><span style="color: green;"><b>鼠标移动：<b></span>示例：200,200 // 鼠标移动到坐标 x , y 为200 像素的地方。</li>
                    <li><span style="color: green;"><b>鼠标拖动：<b></span>示例：200,200,2 //鼠标拖动到 x , y 为200 像素的地方。一定要加入一个时间参数拖动太快会失效。</li>
                    <li><span style="color: green;"><b>鼠标滚动：<b></span>示例：200   //上滚 200 下滚 -200 像素单位。一般电脑屏幕 1919 * 1079像素。</li>
                    <li><span style="color: green;"><b>停留时间：<b></span>示例：10 //单位(秒) 程序运行到某个操作时停留的时间。</li>
                    <li><span style="color: green;"><b>移动窗口：<b></span>示例：200,200//移动当前窗口左上角到 x 200 y 200 的位置。</li>
                    <li><span style="color: green;"><b>窗口大小：<b></span>示例：1000,800 //设置当前窗口的大小。</li>
                    <li><span style="color: green;"><b>启动程序：<b></span>示例：键鼠执行器v1.0.exe  //或 C:\\script\\键鼠执行器v1.0.exe 使用完整目录。</li>
                    <li><span style="color: green;"><b>关闭程序：<b></span>示例：键鼠执行器v1.0.exe // 或 在任务管理器查看任务进程名。</li>
                </ul>
                <p style="font-size: 12px; text-align: right;">说明版本: <b>v1.0</b> | 作者: <i>海斯</i> | 时间: <i> / / / </i></p>
                <hr>
                <p style="font-size: 14px; font-family: Arial; color: gray;">感谢您的使用！<br>更多信息请访问
                <a href="https://github.com/haisi-ai" style="color: blue;">GitHub</a></p>
            </body>
        </html>
        """

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MousePositionApp()
    window.show()
    sys.exit(app.exec_())
