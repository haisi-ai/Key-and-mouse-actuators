"""
键鼠执行器 v2.0
功能强大的键鼠自动化操作工具
支持录制、脚本编辑、多种操作类型
"""

import sys
import json
import time
import threading
import os
import subprocess
from datetime import datetime
from typing import Optional, List, Dict, Any
from dataclasses import dataclass
from enum import Enum

import pyautogui
import pygetwindow as gw
import psutil
import requests
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal, QSize, QSettings
from PyQt5.QtGui import QFont, QIcon, QColor, QPalette, QBrush, QKeySequence
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QListWidget, QListWidgetItem, QLineEdit,
    QComboBox, QSpinBox, QDoubleSpinBox, QTextEdit, QFileDialog,
    QMessageBox, QTabWidget, QGroupBox, QCheckBox, QProgressBar,
    QSplitter, QFrame, QMenu, QAction, QSystemTrayIcon, QStyleFactory,
    QStyle, QTableWidget, QTableWidgetItem, QHeaderView, QDialog,
    QDialogButtonBox, QShortcut, QScrollArea
)
from pynput import keyboard


# ==================== 数据结构 ====================

class ActionType(Enum):
    """操作类型枚举"""
    KEYBOARD_INPUT = "键盘输入"
    KEY_PRESS = "按键按下"
    KEY_RELEASE = "按键释放"
    HOTKEY = "快捷键"
    MOUSE_CLICK = "鼠标点击"
    MOUSE_DOUBLE_CLICK = "鼠标双击"
    MOUSE_MOVE = "鼠标移动"
    MOUSE_DRAG = "鼠标拖动"
    MOUSE_SCROLL = "鼠标滚动"
    WAIT = "等待"
    MOVE_WINDOW = "移动窗口"
    RESIZE_WINDOW = "调整窗口"
    START_PROGRAM = "启动程序"
    CLOSE_PROGRAM = "关闭程序"
    ACTIVATE_WINDOW = "激活窗口"


@dataclass
class Action:
    """操作数据类"""
    type: ActionType
    params: Dict[str, Any]
    description: str
    enabled: bool = True

    def to_dict(self) -> dict:
        return {
            "type": self.type.value,
            "params": self.params,
            "description": self.description,
            "enabled": self.enabled
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'Action':
        return cls(
            type=ActionType(data["type"]),
            params=data["params"],
            description=data["description"],
            enabled=data.get("enabled", True)
        )


# ==================== 进程/窗口查看器对话框 ====================

class ProcessWindowViewer(QDialog):
    """进程和窗口查看器"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("系统信息查看器")
        self.setMinimumSize(800, 600)
        self.setModal(False)

        layout = QVBoxLayout(self)

        # 选项卡
        tab_widget = QTabWidget()

        # 进程选项卡
        process_tab = self.create_process_tab()
        tab_widget.addTab(process_tab, "📊 系统进程")

        # 窗口选项卡
        window_tab = self.create_window_tab()
        tab_widget.addTab(window_tab, "🪟 活动窗口")

        layout.addWidget(tab_widget)

        # 刷新按钮
        refresh_btn = QPushButton("🔄 刷新")
        refresh_btn.clicked.connect(self.refresh_all)
        layout.addWidget(refresh_btn)

        self.refresh_all()

    def create_process_tab(self):
        """创建进程选项卡"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # 搜索框
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("搜索:"))
        self.process_search = QLineEdit()
        self.process_search.setPlaceholderText("输入进程名过滤...")
        self.process_search.textChanged.connect(self.filter_processes)
        search_layout.addWidget(self.process_search)
        layout.addLayout(search_layout)

        # 进程表格
        self.process_table = QTableWidget()
        self.process_table.setColumnCount(4)
        self.process_table.setHorizontalHeaderLabels(["PID", "进程名称", "内存使用(MB)", "状态"])
        self.process_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.process_table.setAlternatingRowColors(True)
        self.process_table.itemDoubleClicked.connect(self.on_process_double_click)
        layout.addWidget(self.process_table)

        # 提示标签
        hint_label = QLabel("💡 提示：双击进程名可复制名称，用于关闭程序操作")
        hint_label.setStyleSheet("color: #89b4fa; padding: 5px;")
        layout.addWidget(hint_label)

        return widget

    def create_window_tab(self):
        """创建窗口选项卡"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # 搜索框
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("搜索:"))
        self.window_search = QLineEdit()
        self.window_search.setPlaceholderText("输入窗口标题过滤...")
        self.window_search.textChanged.connect(self.filter_windows)
        search_layout.addWidget(self.window_search)
        layout.addLayout(search_layout)

        # 窗口表格
        self.window_table = QTableWidget()
        self.window_table.setColumnCount(4)
        self.window_table.setHorizontalHeaderLabels(["窗口标题", "位置(X,Y)", "大小(W×H)", "状态"])
        self.window_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.window_table.setAlternatingRowColors(True)
        self.window_table.itemDoubleClicked.connect(self.on_window_double_click)
        layout.addWidget(self.window_table)

        # 操作按钮
        btn_layout = QHBoxLayout()
        activate_btn = QPushButton("激活选中窗口")
        activate_btn.clicked.connect(self.activate_selected_window)
        btn_layout.addWidget(activate_btn)

        bring_btn = QPushButton("置顶选中窗口")
        bring_btn.clicked.connect(self.bring_window_to_front)
        btn_layout.addWidget(bring_btn)

        layout.addLayout(btn_layout)

        # 提示标签
        hint_label = QLabel("💡 提示：双击窗口标题可复制，用于窗口操作")
        hint_label.setStyleSheet("color: #89b4fa; padding: 5px;")
        layout.addWidget(hint_label)

        return widget

    def refresh_all(self):
        """刷新所有数据"""
        self.refresh_processes()
        self.refresh_windows()

    def refresh_processes(self):
        """刷新进程列表"""
        self.process_table.setRowCount(0)
        processes = []

        for proc in psutil.process_iter(['pid', 'name', 'memory_info']):
            try:
                info = proc.info
                mem_mb = info['memory_info'].rss / 1024 / 1024 if info['memory_info'] else 0
                processes.append((info['pid'], info['name'], mem_mb))
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        # 按内存排序
        processes.sort(key=lambda x: x[2], reverse=True)

        for pid, name, mem_mb in processes[:100]:  # 只显示前100个
            row = self.process_table.rowCount()
            self.process_table.insertRow(row)
            self.process_table.setItem(row, 0, QTableWidgetItem(str(pid)))
            self.process_table.setItem(row, 1, QTableWidgetItem(name))
            self.process_table.setItem(row, 2, QTableWidgetItem(f"{mem_mb:.1f}"))
            self.process_table.setItem(row, 3, QTableWidgetItem("运行中"))

    def refresh_windows(self):
        """刷新窗口列表"""
        self.window_table.setRowCount(0)

        windows = gw.getAllWindows()

        for window in windows:
            if window.title and window.title.strip():  # 只显示有标题的窗口
                row = self.window_table.rowCount()
                self.window_table.insertRow(row)
                self.window_table.setItem(row, 0, QTableWidgetItem(window.title))
                self.window_table.setItem(row, 1, QTableWidgetItem(f"({window.left}, {window.top})"))
                self.window_table.setItem(row, 2, QTableWidgetItem(f"{window.width}×{window.height}"))
                status = "激活" if window.isActive else "正常"
                self.window_table.setItem(row, 3, QTableWidgetItem(status))

    def filter_processes(self):
        """过滤进程"""
        search_text = self.process_search.text().lower()
        for row in range(self.process_table.rowCount()):
            name_item = self.process_table.item(row, 1)
            if name_item:
                visible = search_text in name_item.text().lower()
                self.process_table.setRowHidden(row, not visible)

    def filter_windows(self):
        """过滤窗口"""
        search_text = self.window_search.text().lower()
        for row in range(self.window_table.rowCount()):
            title_item = self.window_table.item(row, 0)
            if title_item:
                visible = search_text in title_item.text().lower()
                self.window_table.setRowHidden(row, not visible)

    def on_process_double_click(self, item):
        """双击进程复制名称"""
        row = item.row()
        name_item = self.process_table.item(row, 1)
        if name_item:
            clipboard = QApplication.clipboard()
            clipboard.setText(name_item.text())
            QMessageBox.information(self, "已复制", f"进程名 '{name_item.text()}' 已复制到剪贴板")

    def on_window_double_click(self, item):
        """双击窗口复制标题"""
        row = item.row()
        title_item = self.window_table.item(row, 0)
        if title_item:
            clipboard = QApplication.clipboard()
            clipboard.setText(title_item.text())
            QMessageBox.information(self, "已复制", f"窗口标题 '{title_item.text()}' 已复制到剪贴板")

    def activate_selected_window(self):
        """激活选中的窗口"""
        current_row = self.window_table.currentRow()
        if current_row >= 0:
            title_item = self.window_table.item(current_row, 0)
            if title_item:
                windows = gw.getWindowsWithTitle(title_item.text())
                if windows:
                    windows[0].activate()
                    QMessageBox.information(self, "成功", f"已激活窗口: {title_item.text()}")

    def bring_window_to_front(self):
        """将窗口置顶"""
        current_row = self.window_table.currentRow()
        if current_row >= 0:
            title_item = self.window_table.item(current_row, 0)
            if title_item:
                windows = gw.getWindowsWithTitle(title_item.text())
                if windows:
                    windows[0].activate()
                    windows[0].raise_()
                    QMessageBox.information(self, "成功", f"已将窗口置顶: {title_item.text()}")


# ==================== 帮助对话框 ====================

class HelpDialog(QDialog):
    """帮助文档对话框"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("使用帮助")
        self.setMinimumSize(800, 600)
        self.setModal(False)

        layout = QVBoxLayout(self)

        # 创建滚动区域
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)

        # 帮助内容
        help_text = QTextEdit()
        help_text.setReadOnly(True)
        help_text.setHtml(self.get_help_html())
        content_layout.addWidget(help_text)

        scroll.setWidget(content_widget)
        layout.addWidget(scroll)

        # 关闭按钮
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)

    def get_help_html(self):
        """获取帮助HTML内容"""
        return """
        <html>
        <head>
            <style>
                body { font-family: 'Microsoft YaHei', sans-serif; margin: 20px; }
                h1 { color: #89b4fa; border-bottom: 2px solid #89b4fa; padding-bottom: 10px; }
                h2 { color: #a6e3a1; margin-top: 20px; }
                h3 { color: #f9e2af; }
                .shortcut { background-color: #313244; padding: 3px 8px; border-radius: 4px; font-family: monospace; }
                .note { background-color: #1e1e2e; border-left: 3px solid #89b4fa; padding: 10px; margin: 10px 0; }
                table { border-collapse: collapse; width: 100%; margin: 10px 0; }
                th, td { border: 1px solid #313244; padding: 8px; text-align: left; }
                th { background-color: #313244; }
            </style>
        </head>
        <body>
            <h1>📖 键鼠执行器 v2.0 使用帮助</h1>

            <h2>🎯 快速开始</h2>
            <p>键鼠执行器是一个强大的自动化工具，可以模拟键盘和鼠标操作，帮助您自动完成重复性任务。</p>

            <div class="note">
                <strong>💡 提示：</strong> 建议先使用"录制"功能快速创建操作序列，然后根据需要手动调整参数。
            </div>

            <h2>⌨️ 快捷键</h2>
            <table>
                <tr><th>快捷键</th><th>功能</th></tr>
                <tr><td class="shortcut">F8</td><td>开始/停止执行</td></tr>
                <tr><td class="shortcut">F9</td><td>暂停/继续执行</td></tr>
                <tr><td class="shortcut">F10</td><td>开始/停止录制</td></tr>
                <tr><td class="shortcut">F1</td><td>打开帮助文档</td></tr>
            </table>

            <h2>📝 操作类型详解</h2>

            <h3>⌨️ 键盘操作</h3>
            <ul>
                <li><strong>键盘输入</strong> - 模拟打字输入文本，参数：要输入的文本内容</li>
                <li><strong>按键按下</strong> - 按下指定按键，参数：按键名称（如 a, enter, space）</li>
                <li><strong>按键释放</strong> - 释放指定按键</li>
                <li><strong>快捷键</strong> - 组合键操作，参数：如 ctrl+c, alt+tab</li>
            </ul>

            <h3>🖱️ 鼠标操作</h3>
            <ul>
                <li><strong>鼠标点击</strong> - 单击指定位置，参数：坐标、按键(left/right/middle)、点击次数</li>
                <li><strong>鼠标双击</strong> - 双击指定位置</li>
                <li><strong>鼠标移动</strong> - 移动鼠标到指定位置，参数：坐标、持续时间</li>
                <li><strong>鼠标拖动</strong> - 拖动鼠标，参数：目标坐标、持续时间、按键</li>
                <li><strong>鼠标滚动</strong> - 滚动滚轮，参数：滚动量（正数向上，负数向下）</li>
            </ul>

            <h3>🪟 窗口操作</h3>
            <ul>
                <li><strong>移动窗口</strong> - 移动窗口到指定位置，参数：窗口标题、X坐标、Y坐标</li>
                <li><strong>调整窗口</strong> - 调整窗口大小，参数：窗口标题、宽度、高度</li>
                <li><strong>激活窗口</strong> - 激活指定窗口，参数：窗口标题</li>
            </ul>

            <h3>📦 程序操作</h3>
            <ul>
                <li><strong>启动程序</strong> - 启动外部程序，参数：程序完整路径</li>
                <li><strong>关闭程序</strong> - 关闭指定程序，参数：进程名称</li>
            </ul>

            <h3>⏱️ 流程控制</h3>
            <ul>
                <li><strong>等待</strong> - 暂停执行，参数：等待秒数</li>
            </ul>

            <h2>🎙️ 录制功能</h2>
            <p>录制功能可以捕获您的鼠标点击、滚动和键盘按键，自动生成操作序列：</p>
            <ol>
                <li>切换到"录制"选项卡</li>
                <li>选择要录制的操作类型（鼠标/键盘）</li>
                <li>点击"开始录制"按钮或按F10键</li>
                <li>执行您想要录制的操作</li>
                <li>点击"停止录制"按钮或按F10键</li>
                <li>录制的操作会自动添加到操作列表</li>
            </ol>

            <h2>💾 脚本管理</h2>
            <ul>
                <li><strong>保存脚本</strong> - 将当前操作列表保存为.kms文件</li>
                <li><strong>加载脚本</strong> - 加载之前保存的脚本文件</li>
                <li><strong>清空列表</strong> - 清空所有操作</li>
            </ul>

            <h2>🔧 高级技巧</h2>
            <div class="note">
                <strong>💡 技巧1：</strong> 使用"系统信息查看器"（工具菜单）可以查看所有运行的进程和窗口，双击即可复制名称用于操作参数。<br><br>
                <strong>💡 技巧2：</strong> 设置循环次数为0可以实现无限循环执行。<br><br>
                <strong>💡 技巧3：</strong> 可以在操作之间添加"等待"操作来控制执行节奏。<br><br>
                <strong>💡 技巧4：</strong> 使用主题切换功能（菜单栏→视图→主题）可以选择深色或浅色主题。
            </div>

            <h2>❓ 常见问题</h2>
            <p><strong>Q: 为什么有些操作没有生效？</strong><br>
            A: 请确保目标窗口处于激活状态，某些程序可能需要管理员权限。</p>

            <p><strong>Q: 如何获取窗口标题？</strong><br>
            A: 使用"工具"菜单中的"活窗"功能，可以查看所有当前窗口的标题。</p>

            <p><strong>Q: 如何获取进程名称？</strong><br>
            A: 使用"工具"菜单中的"进程"功能，可以查看所有运行中的进程名称。</p>

            <p><strong>Q: 录制的操作太多怎么办？</strong><br>
            A: 可以在操作列表中删除不需要的操作，或者手动编辑参数。</p>

            <hr>
            <p style="text-align: center; color: #6c7086;">键鼠执行器 v2.0 | 作者: 海斯 | 如有问题请查看GitHub项目页</p>
        </body>
        </html>
        """


# ==================== 主题管理器 ====================

class ThemeManager:
    """主题管理器"""

    DARK_THEME = """
        QMainWindow {
            background-color: #1e1e2e;
        }
        QWidget {
            background-color: #1e1e2e;
            color: #cdd6f4;
            font-family: 'Microsoft YaHei', 'Segoe UI', sans-serif;
        }
        QPushButton {
            background-color: #313244;
            border: none;
            border-radius: 6px;
            padding: 8px 16px;
            font-size: 13px;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #45475a;
        }
        QPushButton:pressed {
            background-color: #1e1e2e;
        }
        QPushButton#primary {
            background-color: #89b4fa;
            color: #1e1e2e;
        }
        QPushButton#primary:hover {
            background-color: #b4befe;
        }
        QPushButton#danger {
            background-color: #f38ba8;
            color: #1e1e2e;
        }
        QPushButton#danger:hover {
            background-color: #eba0ac;
        }
        QPushButton#success {
            background-color: #a6e3a1;
            color: #1e1e2e;
        }
        QLineEdit, QTextEdit, QComboBox, QSpinBox, QDoubleSpinBox {
            background-color: #313244;
            border: 1px solid #45475a;
            border-radius: 4px;
            padding: 6px 10px;
            font-size: 12px;
        }
        QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
            border-color: #89b4fa;
        }
        QListWidget, QTableWidget {
            background-color: #181825;
            border: 1px solid #313244;
            border-radius: 6px;
            outline: none;
        }
        QListWidget::item, QTableWidget::item {
            padding: 8px;
            border-bottom: 1px solid #313244;
        }
        QListWidget::item:selected, QTableWidget::item:selected {
            background-color: #89b4fa;
            color: #1e1e2e;
        }
        QListWidget::item:hover:!selected, QTableWidget::item:hover:!selected {
            background-color: #313244;
        }
        QTabWidget::pane {
            background-color: #1e1e2e;
            border: 1px solid #313244;
            border-radius: 6px;
        }
        QTabBar::tab {
            background-color: #313244;
            padding: 8px 20px;
            margin-right: 2px;
            border-top-left-radius: 4px;
            border-top-right-radius: 4px;
        }
        QTabBar::tab:selected {
            background-color: #89b4fa;
            color: #1e1e2e;
        }
        QTabBar::tab:hover:!selected {
            background-color: #45475a;
        }
        QGroupBox {
            border: 1px solid #313244;
            border-radius: 6px;
            margin-top: 10px;
            padding-top: 10px;
            font-weight: bold;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px;
        }
        QStatusBar {
            background-color: #181825;
            color: #a6e3a1;
        }
        QMenuBar {
            background-color: #181825;
            border-bottom: 1px solid #313244;
        }
        QMenuBar::item:selected {
            background-color: #89b4fa;
            color: #1e1e2e;
        }
        QMenu {
            background-color: #313244;
            border: 1px solid #45475a;
        }
        QMenu::item:selected {
            background-color: #89b4fa;
            color: #1e1e2e;
        }
        QScrollBar:vertical {
            background-color: #181825;
            width: 10px;
            border-radius: 5px;
        }
        QScrollBar::handle:vertical {
            background-color: #45475a;
            border-radius: 5px;
            min-height: 20px;
        }
        QScrollBar::handle:vertical:hover {
            background-color: #89b4fa;
        }
        QProgressBar {
            background-color: #313244;
            border-radius: 4px;
            text-align: center;
        }
        QProgressBar::chunk {
            background-color: #89b4fa;
            border-radius: 4px;
        }
        QHeaderView::section {
            background-color: #313244;
            padding: 5px;
            border: none;
        }
    """

    LIGHT_THEME = """
        QMainWindow {
            background-color: #f0f0f0;
        }
        QWidget {
            background-color: #f0f0f0;
            color: #2c3e50;
            font-family: 'Microsoft YaHei', 'Segoe UI', sans-serif;
        }
        QPushButton {
            background-color: #e0e0e0;
            border: 1px solid #bdc3c7;
            border-radius: 6px;
            padding: 8px 16px;
            font-size: 13px;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #d0d0d0;
        }
        QPushButton:pressed {
            background-color: #c0c0c0;
        }
        QPushButton#primary {
            background-color: #3498db;
            color: white;
            border: none;
        }
        QPushButton#primary:hover {
            background-color: #2980b9;
        }
        QPushButton#danger {
            background-color: #e74c3c;
            color: white;
            border: none;
        }
        QPushButton#danger:hover {
            background-color: #c0392b;
        }
        QPushButton#success {
            background-color: #27ae60;
            color: white;
            border: none;
        }
        QLineEdit, QTextEdit, QComboBox, QSpinBox, QDoubleSpinBox {
            background-color: white;
            border: 1px solid #bdc3c7;
            border-radius: 4px;
            padding: 6px 10px;
            font-size: 12px;
        }
        QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
            border-color: #3498db;
        }
        QListWidget, QTableWidget {
            background-color: white;
            border: 1px solid #bdc3c7;
            border-radius: 6px;
            outline: none;
        }
        QListWidget::item, QTableWidget::item {
            padding: 8px;
            border-bottom: 1px solid #ecf0f1;
        }
        QListWidget::item:selected, QTableWidget::item:selected {
            background-color: #3498db;
            color: white;
        }
        QListWidget::item:hover:!selected, QTableWidget::item:hover:!selected {
            background-color: #ecf0f1;
        }
        QTabWidget::pane {
            background-color: white;
            border: 1px solid #bdc3c7;
            border-radius: 6px;
        }
        QTabBar::tab {
            background-color: #e0e0e0;
            padding: 8px 20px;
            margin-right: 2px;
            border-top-left-radius: 4px;
            border-top-right-radius: 4px;
        }
        QTabBar::tab:selected {
            background-color: #3498db;
            color: white;
        }
        QGroupBox {
            border: 1px solid #bdc3c7;
            border-radius: 6px;
            margin-top: 10px;
            padding-top: 10px;
            font-weight: bold;
        }
        QStatusBar {
            background-color: #ecf0f1;
            color: #2c3e50;
        }
        QMenuBar {
            background-color: #ecf0f1;
            border-bottom: 1px solid #bdc3c7;
        }
        QMenuBar::item:selected {
            background-color: #3498db;
            color: white;
        }
        QMenu {
            background-color: white;
            border: 1px solid #bdc3c7;
        }
        QMenu::item:selected {
            background-color: #3498db;
            color: white;
        }
        QProgressBar {
            background-color: #ecf0f1;
            border-radius: 4px;
            text-align: center;
        }
        QProgressBar::chunk {
            background-color: #27ae60;
            border-radius: 4px;
        }
        QHeaderView::section {
            background-color: #ecf0f1;
            padding: 5px;
            border: none;
        }
    """

    @staticmethod
    def apply_dark_theme(app):
        app.setStyleSheet(ThemeManager.DARK_THEME)

    @staticmethod
    def apply_light_theme(app):
        app.setStyleSheet(ThemeManager.LIGHT_THEME)


# ==================== 执行引擎 ====================

class ExecutionEngine(QThread):
    """独立线程的执行引擎"""
    log_signal = pyqtSignal(str, str)
    progress_signal = pyqtSignal(int, int)
    finished_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.actions: List[Action] = []
        self.is_running = False
        self.is_paused = False
        self.loop_count = 1
        self.interval = 0.5

    def set_actions(self, actions: List[Action], loop_count: int, interval: float):
        self.actions = actions
        self.loop_count = loop_count
        self.interval = interval

    def run(self):
        self.is_running = True
        executed_loops = 0

        while self.is_running and (self.loop_count == 0 or executed_loops < self.loop_count):
            if self.is_paused:
                time.sleep(0.1)
                continue

            try:
                self._execute_actions()
            except Exception as e:
                self.log_signal.emit(f"执行错误: {str(e)}", "error")

            executed_loops += 1
            self.progress_signal.emit(executed_loops, self.loop_count if self.loop_count > 0 else 1)

            if self.loop_count == 0:
                self.log_signal.emit(f"完成第 {executed_loops} 次循环，继续执行...", "info")

        self.is_running = False
        self.finished_signal.emit()

    def _execute_actions(self):
        """执行操作列表"""
        for action in self.actions:
            if not self.is_running:
                break
            if self.is_paused:
                while self.is_paused and self.is_running:
                    time.sleep(0.1)
                if not self.is_running:
                    break

            if not action.enabled:
                continue

            try:
                self._execute_single_action(action)
                time.sleep(self.interval)
            except Exception as e:
                self.log_signal.emit(f"执行失败 [{action.description}]: {str(e)}", "error")

    def _execute_single_action(self, action: Action):
        """执行单个操作"""
        params = action.params

        if action.type == ActionType.KEYBOARD_INPUT:
            text = params.get("text", "")
            interval = params.get("interval", 0.05)
            pyautogui.write(text, interval=interval)
            self.log_signal.emit(f"键盘输入: {text[:50]}{'...' if len(text) > 50 else ''}", "info")

        elif action.type == ActionType.KEY_PRESS:
            key = params.get("key", "")
            pyautogui.keyDown(key)
            self.log_signal.emit(f"按键按下: {key}", "info")

        elif action.type == ActionType.KEY_RELEASE:
            key = params.get("key", "")
            pyautogui.keyUp(key)
            self.log_signal.emit(f"按键释放: {key}", "info")

        elif action.type == ActionType.HOTKEY:
            keys = params.get("keys", "").split("+")
            pyautogui.hotkey(*keys)
            self.log_signal.emit(f"快捷键: {'+'.join(keys)}", "info")

        elif action.type == ActionType.MOUSE_CLICK:
            x = params.get("x")
            y = params.get("y")
            button = params.get("button", "left")
            clicks = params.get("clicks", 1)

            if x is not None and y is not None:
                pyautogui.click(x, y, button=button, clicks=clicks)
            else:
                pyautogui.click(button=button, clicks=clicks)
            self.log_signal.emit(f"鼠标点击: {button}", "info")

        elif action.type == ActionType.MOUSE_DOUBLE_CLICK:
            x = params.get("x")
            y = params.get("y")
            button = params.get("button", "left")
            if x is not None and y is not None:
                pyautogui.doubleClick(x, y, button=button)
            else:
                pyautogui.doubleClick(button=button)
            self.log_signal.emit(f"鼠标双击: {button}", "info")

        elif action.type == ActionType.MOUSE_MOVE:
            x = params.get("x", 0)
            y = params.get("y", 0)
            duration = params.get("duration", 0.2)
            pyautogui.moveTo(x, y, duration=duration)
            self.log_signal.emit(f"鼠标移动: ({x}, {y})", "info")

        elif action.type == ActionType.MOUSE_DRAG:
            x = params.get("x", 0)
            y = params.get("y", 0)
            duration = params.get("duration", 0.5)
            button = params.get("button", "left")
            pyautogui.dragTo(x, y, duration=duration, button=button)
            self.log_signal.emit(f"鼠标拖动: ({x}, {y})", "info")

        elif action.type == ActionType.MOUSE_SCROLL:
            amount = params.get("amount", 0)
            pyautogui.scroll(amount)
            self.log_signal.emit(f"鼠标滚动: {amount}", "info")

        elif action.type == ActionType.WAIT:
            seconds = params.get("seconds", 1)
            for _ in range(int(seconds * 10)):
                if not self.is_running or self.is_paused:
                    break
                time.sleep(0.1)
            self.log_signal.emit(f"等待: {seconds} 秒", "info")

        elif action.type == ActionType.MOVE_WINDOW:
            title = params.get("title", "")
            x = params.get("x", 0)
            y = params.get("y", 0)
            windows = gw.getWindowsWithTitle(title)
            if windows:
                windows[0].moveTo(x, y)
                self.log_signal.emit(f"移动窗口: '{title}' 到 ({x}, {y})", "info")
            else:
                self.log_signal.emit(f"未找到窗口: {title}", "warning")

        elif action.type == ActionType.RESIZE_WINDOW:
            title = params.get("title", "")
            width = params.get("width", 800)
            height = params.get("height", 600)
            windows = gw.getWindowsWithTitle(title)
            if windows:
                windows[0].resizeTo(width, height)
                self.log_signal.emit(f"调整窗口: '{title}' 到 {width}x{height}", "info")
            else:
                self.log_signal.emit(f"未找到窗口: {title}", "warning")

        elif action.type == ActionType.ACTIVATE_WINDOW:
            title = params.get("title", "")
            windows = gw.getWindowsWithTitle(title)
            if windows:
                windows[0].activate()
                self.log_signal.emit(f"激活窗口: '{title}'", "info")
            else:
                self.log_signal.emit(f"未找到窗口: {title}", "warning")

        elif action.type == ActionType.START_PROGRAM:
            path = params.get("path", "")
            if os.path.exists(path):
                subprocess.Popen(path)
                self.log_signal.emit(f"启动程序: {path}", "info")
            else:
                self.log_signal.emit(f"程序不存在: {path}", "error")

        elif action.type == ActionType.CLOSE_PROGRAM:
            name = params.get("name", "")
            for proc in psutil.process_iter(['pid', 'name']):
                if name.lower() in proc.info['name'].lower():
                    proc.terminate()
                    self.log_signal.emit(f"关闭程序: {name}", "info")
                    return
            self.log_signal.emit(f"未找到程序: {name}", "warning")

    def stop(self):
        self.is_running = False

    def pause(self):
        self.is_paused = True

    def resume(self):
        self.is_paused = False


# ==================== 录制器 ====================

class Recorder(QThread):
    """鼠标键盘录制器"""
    log_signal = pyqtSignal(str, str)
    action_recorded = pyqtSignal(Action)

    def __init__(self):
        super().__init__()
        self.is_recording = False

    def run(self):
        self.is_recording = True

        def on_click(x, y, button, pressed):
            if self.is_recording and pressed:
                action = Action(
                    type=ActionType.MOUSE_CLICK,
                    params={"x": x, "y": y, "button": button.name, "clicks": 1},
                    description=f"鼠标点击: {button.name} at ({x}, {y})"
                )
                self.action_recorded.emit(action)
                self.log_signal.emit(f"录制点击: ({x}, {y})", "info")

        def on_scroll(x, y, dx, dy):
            if self.is_recording:
                action = Action(
                    type=ActionType.MOUSE_SCROLL,
                    params={"amount": dy * 100},
                    description=f"鼠标滚动: {dy * 100}"
                )
                self.action_recorded.emit(action)
                self.log_signal.emit(f"录制滚动: {dy * 100}", "info")

        def on_press(key):
            if self.is_recording:
                try:
                    key_name = key.char if hasattr(key, 'char') and key.char else str(key).replace("Key.", "")
                    action = Action(
                        type=ActionType.KEY_PRESS,
                        params={"key": key_name},
                        description=f"按键: {key_name}"
                    )
                    self.action_recorded.emit(action)
                    self.log_signal.emit(f"录制按键: {key_name}", "info")
                except:
                    pass

        from pynput import mouse, keyboard

        self.mouse_listener = mouse.Listener(on_click=on_click, on_scroll=on_scroll)
        self.keyboard_listener = keyboard.Listener(on_press=on_press)

        self.mouse_listener.start()
        self.keyboard_listener.start()

        self.mouse_listener.join()
        self.keyboard_listener.join()

    def stop(self):
        self.is_recording = False
        if hasattr(self, 'mouse_listener') and self.mouse_listener:
            self.mouse_listener.stop()
        if hasattr(self, 'keyboard_listener') and self.keyboard_listener:
            self.keyboard_listener.stop()


# ==================== 主窗口 ====================

class KeyMouseExecutor(QMainWindow):
    """主窗口类"""

    def __init__(self):
        super().__init__()

        # 初始化变量
        self.actions: List[Action] = []
        self.engine: Optional[ExecutionEngine] = None
        self.recorder: Optional[Recorder] = None
        self.is_editing = False
        self.editing_index = -1
        self.loop_count = 1
        self.interval = 0.5
        self.current_theme = "dark"  # dark/light

        # 设置窗口
        self.setWindowTitle("键鼠执行器 v2.0 - 智能自动化工具")
        self.setGeometry(100, 100, 1200, 800)
        self.setMinimumSize(1000, 600)

        # 应用深色主题
        ThemeManager.apply_dark_theme(QApplication.instance())

        # 初始化UI
        self.setup_ui()

        # 初始化菜单栏
        self.setup_menu()

        # 初始化系统托盘
        self.setup_tray()

        # 启动鼠标位置更新
        self.start_mouse_tracker()

        # 注册全局热键
        self.register_hotkeys()

    def setup_menu(self):
        """设置菜单栏"""
        menubar = self.menuBar()

        # 文件菜单
        file_menu = menubar.addMenu("文件")

        save_action = QAction("保存脚本", self)
        save_action.setShortcut(QKeySequence("Ctrl+S"))
        save_action.triggered.connect(self.save_script)
        file_menu.addAction(save_action)

        load_action = QAction("加载脚本", self)
        load_action.setShortcut(QKeySequence("Ctrl+O"))
        load_action.triggered.connect(self.load_script)
        file_menu.addAction(load_action)

        file_menu.addSeparator()

        exit_action = QAction("退出", self)
        exit_action.setShortcut(QKeySequence("Alt+F4"))
        exit_action.triggered.connect(self.quit_app)
        file_menu.addAction(exit_action)

        # 工具菜单
        tool_menu = menubar.addMenu("工具")

        process_action = QAction("系统进程", self)
        process_action.triggered.connect(self.show_process_viewer)
        tool_menu.addAction(process_action)

        window_action = QAction("活动窗口", self)
        window_action.triggered.connect(self.show_window_viewer)
        tool_menu.addAction(window_action)

        # 视图菜单
        view_menu = menubar.addMenu("视图")

        dark_theme_action = QAction("深色主题", self)
        dark_theme_action.triggered.connect(lambda: self.switch_theme("dark"))
        view_menu.addAction(dark_theme_action)

        light_theme_action = QAction("浅色主题", self)
        light_theme_action.triggered.connect(lambda: self.switch_theme("light"))
        view_menu.addAction(light_theme_action)

        # 帮助菜单
        help_menu = menubar.addMenu("帮助")

        help_action = QAction("使用帮助", self)
        help_action.setShortcut(QKeySequence("F1"))
        help_action.triggered.connect(self.show_help)
        help_menu.addAction(help_action)

        about_action = QAction("关于", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def switch_theme(self, theme: str):
        """切换主题"""
        self.current_theme = theme
        if theme == "dark":
            ThemeManager.apply_dark_theme(QApplication.instance())
        else:
            ThemeManager.apply_light_theme(QApplication.instance())
        self.log(f"已切换到{('深色' if theme == 'dark' else '浅色')}主题", "info")

    def show_process_viewer(self):
        """显示进程查看器"""
        self.process_viewer = ProcessWindowViewer(self)
        self.process_viewer.show()

    def show_window_viewer(self):
        """显示窗口查看器"""
        self.window_viewer = ProcessWindowViewer(self)
        self.window_viewer.show()

    def show_help(self):
        """显示帮助对话框"""
        self.help_dialog = HelpDialog(self)
        self.help_dialog.show()

    def show_about(self):
        """显示关于对话框"""
        QMessageBox.about(
            self,
            "关于键鼠执行器",
            f"键鼠执行器 v2.0\n\n"
            "功能强大的键鼠自动化工具\n"
            "支持录制、脚本编辑、多种操作类型\n\n"
            "作者: 海斯\n"
            "GitHub: github.com/haisi-ai\n\n"
            "快捷键:\n"
            "  F8 - 开始/停止执行\n"
            "  F9 - 暂停/继续执行\n"
            "  F10 - 开始/停止录制\n"
            "  F1 - 打开帮助"
        )

    def setup_ui(self):
        """设置用户界面"""
        # 中央控件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # 顶部工具栏
        self.setup_toolbar(main_layout)

        # 主分割区域
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)

        # 左侧：操作列表区域
        left_widget = self.setup_left_panel()
        splitter.addWidget(left_widget)

        # 右侧：操作编辑和日志区域
        right_widget = self.setup_right_panel()
        splitter.addWidget(right_widget)

        # 设置分割比例
        splitter.setSizes([500, 500])

        # 底部状态栏
        self.setup_status_bar()

    def setup_toolbar(self, parent_layout):
        """设置顶部工具栏"""
        toolbar = QHBoxLayout()
        toolbar.setSpacing(10)

        # 鼠标坐标显示
        self.coord_label = QLabel("坐标: [0, 0]")
        self.coord_label.setStyleSheet("""
            background-color: #313244;
            padding: 8px 15px;
            border-radius: 6px;
            font-size: 14px;
            font-weight: bold;
        """)
        toolbar.addWidget(self.coord_label)

        # 分隔
        toolbar.addWidget(self.create_separator())

        # 循环间隔
        toolbar.addWidget(QLabel("间隔:"))
        self.interval_spin = QDoubleSpinBox()
        self.interval_spin.setRange(0.01, 60)
        self.interval_spin.setValue(0.5)
        self.interval_spin.setSuffix(" 秒")
        self.interval_spin.valueChanged.connect(lambda v: setattr(self, 'interval', v))
        toolbar.addWidget(self.interval_spin)

        # 循环次数
        toolbar.addWidget(QLabel("次数:"))
        self.loop_spin = QSpinBox()
        self.loop_spin.setRange(0, 9999)
        self.loop_spin.setValue(1)
        self.loop_spin.setSpecialValueText("无限")
        self.loop_spin.valueChanged.connect(lambda v: setattr(self, 'loop_count', v))
        toolbar.addWidget(self.loop_spin)

        toolbar.addStretch()

        # 控制按钮
        self.start_btn = QPushButton("▶ 开始 (F8)")
        self.start_btn.setObjectName("success")
        self.start_btn.clicked.connect(self.start_execution)
        toolbar.addWidget(self.start_btn)

        self.pause_btn = QPushButton("⏸ 暂停 (F9)")
        self.pause_btn.setEnabled(False)
        self.pause_btn.clicked.connect(self.pause_execution)
        toolbar.addWidget(self.pause_btn)

        self.stop_btn = QPushButton("■ 停止")
        self.stop_btn.setObjectName("danger")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_execution)
        toolbar.addWidget(self.stop_btn)

        parent_layout.addLayout(toolbar)

    def setup_left_panel(self):
        """设置左侧面板"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)

        # 操作列表
        self.action_list = QListWidget()
        self.action_list.itemDoubleClicked.connect(self.edit_action_from_list)
        layout.addWidget(QLabel("📋 操作列表"))
        layout.addWidget(self.action_list)

        # 操作按钮
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        up_btn = QPushButton("↑ 上移")
        up_btn.clicked.connect(self.move_up)
        btn_layout.addWidget(up_btn)

        down_btn = QPushButton("↓ 下移")
        down_btn.clicked.connect(self.move_down)
        btn_layout.addWidget(down_btn)

        delete_btn = QPushButton("🗑 删除")
        delete_btn.clicked.connect(self.delete_action)
        btn_layout.addWidget(delete_btn)

        clear_btn = QPushButton("清空")
        clear_btn.clicked.connect(self.clear_actions)
        btn_layout.addWidget(clear_btn)

        layout.addLayout(btn_layout)

        # 文件操作
        file_layout = QHBoxLayout()
        save_btn = QPushButton("💾 保存脚本")
        save_btn.clicked.connect(self.save_script)
        file_layout.addWidget(save_btn)

        load_btn = QPushButton("📂 加载脚本")
        load_btn.clicked.connect(self.load_script)
        file_layout.addWidget(load_btn)

        layout.addLayout(file_layout)

        return widget

    def setup_right_panel(self):
        """设置右侧面板"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)

        # 选项卡
        tab_widget = QTabWidget()

        # 操作编辑选项卡
        edit_tab = self.setup_edit_tab()
        tab_widget.addTab(edit_tab, "✏ 操作编辑")

        # 录制选项卡
        record_tab = self.setup_record_tab()
        tab_widget.addTab(record_tab, "🎙 录制")

        # 日志选项卡
        log_tab = self.setup_log_tab()
        tab_widget.addTab(log_tab, "📝 日志")

        layout.addWidget(tab_widget)

        return widget

    def setup_edit_tab(self):
        """设置操作编辑选项卡"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # 操作类型选择
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("操作类型:"))

        self.action_type_combo = QComboBox()
        for action_type in ActionType:
            self.action_type_combo.addItem(action_type.value)
        self.action_type_combo.currentTextChanged.connect(self.on_action_type_changed)
        type_layout.addWidget(self.action_type_combo)
        type_layout.addStretch()
        layout.addLayout(type_layout)

        # 参数编辑区域
        self.params_group = QGroupBox("参数设置")
        self.params_layout = QVBoxLayout(self.params_group)
        layout.addWidget(self.params_group)

        # 描述
        layout.addWidget(QLabel("描述:"))
        self.desc_input = QLineEdit()
        layout.addWidget(self.desc_input)

        # 按钮
        btn_layout = QHBoxLayout()
        self.add_btn = QPushButton("➕ 添加操作")
        self.add_btn.setObjectName("primary")
        self.add_btn.clicked.connect(self.add_action)
        btn_layout.addWidget(self.add_btn)

        self.update_btn = QPushButton("🔄 更新操作")
        self.update_btn.clicked.connect(self.update_action)
        self.update_btn.setEnabled(False)
        btn_layout.addWidget(self.update_btn)

        self.cancel_btn = QPushButton("取消编辑")
        self.cancel_btn.clicked.connect(self.cancel_edit)
        self.cancel_btn.setEnabled(False)
        btn_layout.addWidget(self.cancel_btn)

        layout.addLayout(btn_layout)

        # 初始化参数编辑区域
        self.init_params_editor()

        return widget

    def init_params_editor(self):
        """初始化参数编辑器"""
        self.params_widgets = {}

        # 键盘输入参数
        self.add_param_row("text", "输入文本:", QLineEdit, placeholder="要输入的文本内容")
        self.add_param_row("interval", "输入间隔:", QDoubleSpinBox, 0.05, 0, 1)

        # 按键参数
        self.add_param_row("key", "按键名称:", QLineEdit, placeholder="例如: a, enter, space")

        # 快捷键参数
        self.add_param_row("keys", "快捷键组合:", QLineEdit, placeholder="例如: ctrl+c, alt+tab")

        # 鼠标参数
        self.add_param_row("x", "X坐标:", QSpinBox, 0, -9999, 9999)
        self.add_param_row("y", "Y坐标:", QSpinBox, 0, -9999, 9999)
        self.add_param_row("button", "鼠标按键:", QComboBox, items=["left", "right", "middle"])
        self.add_param_row("clicks", "点击次数:", QSpinBox, 1, 1, 10)
        self.add_param_row("duration", "持续时间:", QDoubleSpinBox, 0.2, 0, 10)

        # 滚动参数
        self.add_param_row("amount", "滚动量:", QSpinBox, 100, -9999, 9999)

        # 等待参数
        self.add_param_row("seconds", "等待秒数:", QDoubleSpinBox, 1, 0.1, 3600)

        # 窗口参数
        self.add_param_row("title", "窗口标题:", QLineEdit, placeholder="窗口标题关键词")
        self.add_param_row("width", "宽度:", QSpinBox, 800, 100, 4000)
        self.add_param_row("height", "高度:", QSpinBox, 600, 100, 4000)

        # 程序参数
        self.add_param_row("path", "程序路径:", QLineEdit, placeholder="exe文件路径")
        self.add_param_row("name", "程序名称:", QLineEdit, placeholder="进程名称")

        # 隐藏所有参数
        self.hide_all_params()

    def add_param_row(self, key: str, label: str, widget_type, default=None, min_val=None, max_val=None, items=None,
                      placeholder=None):
        """添加参数行"""
        row = QHBoxLayout()
        label_widget = QLabel(label)
        label_widget.setVisible(False)
        row.addWidget(label_widget)

        if widget_type == QLineEdit:
            widget = QLineEdit()
            if placeholder:
                widget.setPlaceholderText(placeholder)
        elif widget_type == QSpinBox:
            widget = QSpinBox()
            if min_val is not None:
                widget.setMinimum(min_val)
            if max_val is not None:
                widget.setMaximum(max_val)
            if default is not None:
                widget.setValue(default)
        elif widget_type == QDoubleSpinBox:
            widget = QDoubleSpinBox()
            if min_val is not None:
                widget.setMinimum(min_val)
            if max_val is not None:
                widget.setMaximum(max_val)
            if default is not None:
                widget.setValue(default)
            widget.setDecimals(2)
        elif widget_type == QComboBox:
            widget = QComboBox()
            if items:
                widget.addItems(items)
            if default is not None:
                widget.setCurrentText(str(default))
        else:
            return

        widget.setVisible(False)
        row.addWidget(widget)
        self.params_layout.addLayout(row)
        self.params_widgets[key] = (widget, label_widget)

    def hide_all_params(self):
        """隐藏所有参数"""
        for widget, label in self.params_widgets.values():
            widget.setVisible(False)
            label.setVisible(False)

    def show_params_for_type(self, action_type: ActionType):
        """根据操作类型显示对应的参数"""
        self.hide_all_params()

        param_keys = {
            ActionType.KEYBOARD_INPUT: ["text", "interval"],
            ActionType.KEY_PRESS: ["key"],
            ActionType.KEY_RELEASE: ["key"],
            ActionType.HOTKEY: ["keys"],
            ActionType.MOUSE_CLICK: ["x", "y", "button", "clicks"],
            ActionType.MOUSE_DOUBLE_CLICK: ["x", "y", "button"],
            ActionType.MOUSE_MOVE: ["x", "y", "duration"],
            ActionType.MOUSE_DRAG: ["x", "y", "duration", "button"],
            ActionType.MOUSE_SCROLL: ["amount"],
            ActionType.WAIT: ["seconds"],
            ActionType.MOVE_WINDOW: ["title", "x", "y"],
            ActionType.RESIZE_WINDOW: ["title", "width", "height"],
            ActionType.ACTIVATE_WINDOW: ["title"],
            ActionType.START_PROGRAM: ["path"],
            ActionType.CLOSE_PROGRAM: ["name"],
        }

        keys = param_keys.get(action_type, [])
        for key in keys:
            if key in self.params_widgets:
                widget, label = self.params_widgets[key]
                widget.setVisible(True)
                label.setVisible(True)

    def get_current_params(self) -> Dict[str, Any]:
        """获取当前参数"""
        params = {}
        for key, (widget, _) in self.params_widgets.items():
            if widget.isVisible():
                if isinstance(widget, QLineEdit):
                    value = widget.text()
                elif isinstance(widget, (QSpinBox, QDoubleSpinBox)):
                    value = widget.value()
                elif isinstance(widget, QComboBox):
                    value = widget.currentText()
                else:
                    continue
                if value:  # 只保存非空值
                    params[key] = value
        return params

    def on_action_type_changed(self, text: str):
        """操作类型改变时的处理"""
        for action_type in ActionType:
            if action_type.value == text:
                self.show_params_for_type(action_type)
                break

    def setup_record_tab(self):
        """设置录制选项卡"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # 录制控制
        record_group = QGroupBox("录制控制")
        record_layout = QVBoxLayout(record_group)

        self.record_btn = QPushButton("🔴 开始录制 (F10)")
        self.record_btn.setObjectName("danger")
        self.record_btn.clicked.connect(self.toggle_recording)
        record_layout.addWidget(self.record_btn)

        record_layout.addWidget(QLabel("提示: 录制将捕获鼠标点击、滚动和键盘按键"))

        layout.addWidget(record_group)

        # 录制选项
        options_group = QGroupBox("录制选项")
        options_layout = QVBoxLayout(options_group)

        self.record_mouse_cb = QCheckBox("录制鼠标操作")
        self.record_mouse_cb.setChecked(True)
        options_layout.addWidget(self.record_mouse_cb)

        self.record_keyboard_cb = QCheckBox("录制键盘操作")
        self.record_keyboard_cb.setChecked(True)
        options_layout.addWidget(self.record_keyboard_cb)

        layout.addWidget(options_group)

        layout.addStretch()

        return widget

    def setup_log_tab(self):
        """设置日志选项卡"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # 日志工具栏
        log_toolbar = QHBoxLayout()
        clear_log_btn = QPushButton("清空日志")
        clear_log_btn.clicked.connect(self.clear_log)
        log_toolbar.addWidget(clear_log_btn)

        export_log_btn = QPushButton("导出日志")
        export_log_btn.clicked.connect(self.export_log)
        log_toolbar.addWidget(export_log_btn)

        log_toolbar.addStretch()
        layout.addLayout(log_toolbar)

        # 日志显示
        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)
        layout.addWidget(self.log_display)

        return widget

    def setup_status_bar(self):
        """设置状态栏"""
        self.status_bar = self.statusBar()

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumWidth(200)
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)

        # 状态标签
        self.status_label = QLabel("就绪")
        self.status_bar.addPermanentWidget(self.status_label)

    def setup_tray(self):
        """设置系统托盘"""
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))

        tray_menu = QMenu()
        show_action = QAction("显示窗口", self)
        show_action.triggered.connect(self.show_normal)
        tray_menu.addAction(show_action)

        quit_action = QAction("退出", self)
        quit_action.triggered.connect(self.quit_app)
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)
        self.tray_icon.show()

    def start_mouse_tracker(self):
        """启动鼠标位置追踪"""
        self.mouse_timer = QTimer()
        self.mouse_timer.timeout.connect(self.update_mouse_position)
        self.mouse_timer.start(50)

    def update_mouse_position(self):
        """更新鼠标位置显示"""
        x, y = pyautogui.position()
        self.coord_label.setText(f"📍 坐标: [{x}, {y}]")

    def register_hotkeys(self):
        """注册全局热键"""

        def on_f8():
            if not self.engine or not self.engine.is_running:
                self.start_execution()
            else:
                self.stop_execution()

        def on_f9():
            if self.engine and self.engine.is_running:
                self.pause_execution()

        def on_f10():
            self.toggle_recording()

        hotkeys = {
            '<f8>': on_f8,
            '<f9>': on_f9,
            '<f10>': on_f10,
        }

        self.hotkey_listener = keyboard.GlobalHotKeys(hotkeys)
        self.hotkey_listener.start()

    # ==================== 操作列表管理 ====================

    def update_action_list(self):
        """更新操作列表显示"""
        self.action_list.clear()
        for i, action in enumerate(self.actions):
            item = QListWidgetItem(f"{i + 1}. {action.description}")
            if not action.enabled:
                item.setForeground(QColor("#6c7086"))
            self.action_list.addItem(item)

    def add_action(self):
        """添加操作"""
        action_type_text = self.action_type_combo.currentText()
        action_type = next(at for at in ActionType if at.value == action_type_text)
        params = self.get_current_params()
        description = self.desc_input.text() or f"{action_type_text}"

        action = Action(type=action_type, params=params, description=description)
        self.actions.append(action)
        self.update_action_list()
        self.log(f"添加操作: {description}", "info")

        # 清空输入
        self.desc_input.clear()

    def edit_action_from_list(self, item):
        """从列表编辑操作"""
        index = self.action_list.row(item)
        self.edit_action(index)

    def edit_action(self, index: int):
        """编辑操作"""
        if 0 <= index < len(self.actions):
            self.is_editing = True
            self.editing_index = index
            action = self.actions[index]

            # 设置类型
            self.action_type_combo.setCurrentText(action.type.value)
            # 设置参数
            for key, value in action.params.items():
                if key in self.params_widgets:
                    widget, _ = self.params_widgets[key]
                    if isinstance(widget, QLineEdit):
                        widget.setText(str(value))
                    elif isinstance(widget, (QSpinBox, QDoubleSpinBox)):
                        widget.setValue(value)
                    elif isinstance(widget, QComboBox):
                        widget.setCurrentText(str(value))
            # 设置描述
            self.desc_input.setText(action.description)

            self.add_btn.setEnabled(False)
            self.update_btn.setEnabled(True)
            self.cancel_btn.setEnabled(True)

    def update_action(self):
        """更新操作"""
        if self.is_editing and 0 <= self.editing_index < len(self.actions):
            action_type_text = self.action_type_combo.currentText()
            action_type = next(at for at in ActionType if at.value == action_type_text)
            params = self.get_current_params()
            description = self.desc_input.text() or f"{action_type_text}"

            self.actions[self.editing_index] = Action(
                type=action_type, params=params, description=description
            )
            self.update_action_list()
            self.log(f"更新操作: {description}", "info")

            self.cancel_edit()

    def cancel_edit(self):
        """取消编辑"""
        self.is_editing = False
        self.editing_index = -1
        self.add_btn.setEnabled(True)
        self.update_btn.setEnabled(False)
        self.cancel_btn.setEnabled(False)
        self.desc_input.clear()

    def delete_action(self):
        """删除操作"""
        current_row = self.action_list.currentRow()
        if current_row >= 0:
            action = self.actions.pop(current_row)
            self.update_action_list()
            self.log(f"删除操作: {action.description}", "info")

    def move_up(self):
        """上移操作"""
        current_row = self.action_list.currentRow()
        if current_row > 0:
            self.actions[current_row], self.actions[current_row - 1] = \
                self.actions[current_row - 1], self.actions[current_row]
            self.update_action_list()
            self.action_list.setCurrentRow(current_row - 1)

    def move_down(self):
        """下移操作"""
        current_row = self.action_list.currentRow()
        if 0 <= current_row < len(self.actions) - 1:
            self.actions[current_row], self.actions[current_row + 1] = \
                self.actions[current_row + 1], self.actions[current_row]
            self.update_action_list()
            self.action_list.setCurrentRow(current_row + 1)

    def clear_actions(self):
        """清空所有操作"""
        if QMessageBox.question(self, "确认", "确定要清空所有操作吗？") == QMessageBox.Yes:
            self.actions.clear()
            self.update_action_list()
            self.log("已清空所有操作", "info")

    # ==================== 脚本保存/加载 ====================

    def save_script(self):
        """保存脚本"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存脚本", "", "键鼠脚本 (*.kms);;JSON文件 (*.json)"
        )
        if file_path:
            try:
                data = {
                    "version": "2.0",
                    "actions": [action.to_dict() for action in self.actions],
                    "settings": {
                        "loop_count": self.loop_count,
                        "interval": self.interval
                    }
                }
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                self.log(f"脚本已保存: {file_path}", "success")
            except Exception as e:
                self.log(f"保存失败: {str(e)}", "error")

    def load_script(self):
        """加载脚本"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "加载脚本", "", "键鼠脚本 (*.kms);;JSON文件 (*.json)"
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                self.actions = [Action.from_dict(act) for act in data.get("actions", [])]
                self.update_action_list()

                if "settings" in data:
                    self.loop_count = data["settings"].get("loop_count", 1)
                    self.interval = data["settings"].get("interval", 0.5)
                    self.loop_spin.setValue(self.loop_count)
                    self.interval_spin.setValue(self.interval)

                self.log(f"脚本已加载: {file_path} (共{len(self.actions)}个操作)", "success")
            except Exception as e:
                self.log(f"加载失败: {str(e)}", "error")

    # ==================== 执行控制 ====================

    def start_execution(self):
        """开始执行"""
        if not self.actions:
            self.log("操作列表为空，请先添加操作！", "warning")
            return

        if self.engine and self.engine.is_running:
            self.log("已有任务在执行中", "warning")
            return

        self.engine = ExecutionEngine()
        self.engine.set_actions(self.actions, self.loop_count, self.interval)
        self.engine.log_signal.connect(self.log)
        self.engine.progress_signal.connect(self.update_progress)
        self.engine.finished_signal.connect(self.on_execution_finished)

        self.engine.start()

        self.start_btn.setEnabled(False)
        self.pause_btn.setEnabled(True)
        self.stop_btn.setEnabled(True)
        self.progress_bar.setVisible(True)
        self.status_label.setText("执行中...")
        self.log("开始执行操作序列", "success")

    def pause_execution(self):
        """暂停执行"""
        if self.engine and self.engine.is_running:
            if self.engine.is_paused:
                self.engine.resume()
                self.pause_btn.setText("⏸ 暂停 (F9)")
                self.log("恢复执行", "info")
            else:
                self.engine.pause()
                self.pause_btn.setText("▶ 继续 (F9)")
                self.log("已暂停执行", "warning")

    def stop_execution(self):
        """停止执行"""
        if self.engine and self.engine.is_running:
            self.engine.stop()
            self.log("正在停止执行...", "warning")

    def on_execution_finished(self):
        """执行完成"""
        self.start_btn.setEnabled(True)
        self.pause_btn.setEnabled(False)
        self.stop_btn.setEnabled(False)
        self.pause_btn.setText("⏸ 暂停 (F9)")
        self.progress_bar.setVisible(False)
        self.status_label.setText("就绪")
        self.log("执行完成", "success")

    def update_progress(self, current, total):
        """更新进度"""
        if total > 0:
            percent = int(current / total * 100)
            self.progress_bar.setValue(percent)
            self.status_label.setText(f"执行中: {current}/{total}")

    # ==================== 录制功能 ====================

    def toggle_recording(self):
        """切换录制状态"""
        if not self.recorder or not self.recorder.is_recording:
            self.start_recording()
        else:
            self.stop_recording()

    def start_recording(self):
        """开始录制"""
        self.recorder = Recorder()
        self.recorder.log_signal.connect(self.log)
        self.recorder.action_recorded.connect(self.on_action_recorded)
        self.recorder.start()

        self.record_btn.setText("⏹ 停止录制 (F10)")
        self.record_btn.setObjectName("")
        self.log("开始录制操作...", "info")

    def stop_recording(self):
        """停止录制"""
        if self.recorder:
            self.recorder.stop()
            self.recorder = None

        self.record_btn.setText("🔴 开始录制 (F10)")
        self.record_btn.setObjectName("danger")
        self.log("停止录制", "info")

    def on_action_recorded(self, action: Action):
        """录制到新操作"""
        if self.record_mouse_cb.isChecked() and action.type in [ActionType.MOUSE_CLICK, ActionType.MOUSE_SCROLL]:
            self.actions.append(action)
            self.update_action_list()
        elif self.record_keyboard_cb.isChecked() and action.type == ActionType.KEY_PRESS:
            self.actions.append(action)
            self.update_action_list()

    # ==================== 日志功能 ====================

    def log(self, message: str, level: str = "info"):
        """添加日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        level_colors = {
            "info": "#89b4fa",
            "success": "#a6e3a1",
            "warning": "#f9e2af",
            "error": "#f38ba8",
        }
        color = level_colors.get(level, "#cdd6f4")

        formatted = f'<span style="color: {color};">[{timestamp}] [{level.upper()}] {message}</span>'
        self.log_display.append(formatted)

        # 滚动到底部
        scrollbar = self.log_display.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def clear_log(self):
        """清空日志"""
        self.log_display.clear()

    def export_log(self):
        """导出日志"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出日志", "", "日志文件 (*.log);;文本文件 (*.txt)"
        )
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.log_display.toPlainText())
                self.log(f"日志已导出: {file_path}", "success")
            except Exception as e:
                self.log(f"导出失败: {str(e)}", "error")

    # ==================== 辅助功能 ====================

    def create_separator(self):
        """创建分隔线"""
        separator = QFrame()
        separator.setFrameShape(QFrame.VLine)
        separator.setStyleSheet("background-color: #313244;")
        return separator

    def show_normal(self):
        """显示窗口"""
        self.showNormal()
        self.activateWindow()

    def on_tray_activated(self, reason):
        """系统托盘激活"""
        if reason == QSystemTrayIcon.DoubleClick:
            self.show_normal()

    def quit_app(self):
        """退出应用"""
        if self.engine and self.engine.is_running:
            reply = QMessageBox.question(
                self, "确认退出",
                "有任务正在执行中，确定要退出吗？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return
            self.engine.stop()

        self.tray_icon.hide()
        QApplication.quit()

    def closeEvent(self, event):
        """关闭事件"""
        reply = QMessageBox.question(
            self, "确认",
            "是否最小化到系统托盘？",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
        )
        if reply == QMessageBox.Yes:
            self.hide()
            self.tray_icon.showMessage(
                "键鼠执行器",
                "程序已最小化到系统托盘",
                QSystemTrayIcon.Information,
                2000
            )
            event.ignore()
        elif reply == QMessageBox.No:
            self.quit_app()
            event.accept()
        else:
            event.ignore()


# ==================== 程序入口 ====================

def main():
    """主函数"""
    pyautogui.FAILSAFE = True

    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create("Fusion"))

    window = KeyMouseExecutor()
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
