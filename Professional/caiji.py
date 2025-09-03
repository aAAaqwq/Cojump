import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.animation as animation
import numpy as np
import socket
import threading
import queue
import time
import csv
import os
import sys
import win32com.client
import colorsys
import io
from collections import deque
import subprocess
import re
import struct


class NetworkDebugger:
    def __init__(self, root):
        # --- 修改：使用更可靠的方式获取程序路径 ---
        if getattr(sys, 'frozen', False):
            # 如果是打包后的exe文件
            self.program_dir = os.path.dirname(sys.executable)
        else:
            # 如果是Python脚本
            self.program_dir = os.path.dirname(os.path.abspath(__file__))

        # --- 修改：使用程序目录作为配置文件路径 ---
        self.config_path = os.path.join(self.program_dir, 'config.txt')
        self.num_channels = self.load_num_channels()
        # 新增：加载主界面y轴范围
        self.y_top_fixed, self.y_bottom_fixed = self.load_yaxis_range(mode="main")
        # 新增：加载实验参数
        self.cycle_total, self.contraction_time, self.relaxation_time, self.voice_broadcast_enabled = self.load_experiment_params()
        # 新增：加载IP地址和端口号
        self.ip_address, self.port_number = self.load_connection_params()

        # 新增：语音播报开关变量（与界面Checkbutton绑定）
        self.voice_broadcast_var = tk.BooleanVar(value=self.voice_broadcast_enabled)
        self.voice_broadcast_var.trace_add('write', self.on_voice_broadcast_toggle)

        self.root = root  # 保存Tkinter根窗口的引用
        self.root.title("肌电信号采集")  # 设置窗口标题
        self.root.geometry("1000x800")  # 设置窗口大小为1000x800像素
        self.root.minsize(700, 600)  # 设置最小窗口尺寸，防止过小

        # 连接状态
        self.connected = False  # 设置连接状态为未连接
        self.sock = None  # 设置套接字为None
        self.client_sock = None  # 设置客户端套接字为None
        self.gui_queue = queue.Queue()  # 创建GUI消息队列

        # 数据采集控制
        self.data_collection_active = False  # 数据采集状态，初始为未采集
        self.collection_started = False  # 新增：采集是否已开始
        self.current_data_file = None  # 当前数据文件
        self.data_counter = 0  # 数据计数器
        self.save_lock = threading.Lock()  # 文件写入锁
        # self.create_new_data_file()  # 创建初始数据文件（已移除，避免无用文件）

        # 新增：语音播报时间点记录
        self.voice_timestamps = []  # 记录语音播报的时间点
        self.voice_timestamp_lock = threading.Lock()  # 语音时间点记录锁

        # 新增：语音播报锁
        self.speech_lock = threading.Lock()

        # 默认通道数
        # self.num_channels = 1 # 原有行可删除或注释
        self.group_data_size = self.num_channels * 4  # 每个通道的数据大小为4字节
        self.packet_size = 2 + 5 * self.group_data_size  # 数据包的总大小

        # 颜色配置
        self.colors = self.generate_colors(16)  # 生成16个通道颜色
        self.channel_visibility = [True] * self.num_channels  # 初始化通道可见性列表，所有通道默认可见

        # 绘图优化配置
        self.plot_update_interval = 20  # 设置绘图更新间隔为20毫秒
        self.last_plot_update = 0  # 记录上次绘图更新时间
        self.plot_paused = False  # 绘图暂停状态


        # 多通道波形数据
        self.max_data_points = 10000  # 设置最大数据点数为10000
        self.plot_data = np.zeros((self.num_channels, self.max_data_points),
                                  dtype=np.float32)  # 创建存储波形数据的数组，16通道，每个通道10000个数据点
        self.time_data = np.zeros(self.max_data_points, dtype=np.float64)  # 初始化时间数据，用于存储每个数据点的时间戳
        self.data_index = 0  # 当前数据索引
        self.start_timestamp = None  # 记录第一个数据包时间戳

        # 暂停功能相关
        self.paused_plot_data = None  # 暂停时的波形数据
        self.paused_time_data = None  # 暂停时的时间数据
        self.paused_data_index = 0  # 暂停时的数据索引

        # 绘图区域大小
        self.plot_width = 1200  # 设置绘图区域宽度为1200像素
        self.plot_height = 500  # 设置绘图区域高度为500像素
        self.left_margin = 50  # 设置左侧边距为60像素
        self.top_margin = 20  # 新增：设置顶部边距为40像素，增大上边距
        self.bottom_margin = 30  # 设置底部边距为30像素
        self.effective_height = self.plot_height - self.bottom_margin - self.top_margin  # 计算有效高度
        self.y_scale = self.effective_height / (self.y_top_fixed - self.y_bottom_fixed)  # 计算y轴缩放比例
        self.plot_lock = threading.Lock()  # 创建绘图线程锁，用于线程安全的数据访问

        # 数据接收与处理
        self.data_queue = queue.Queue(maxsize=10000)  # 创建数据队列，用于存储接收到的数据
        self.processing_active = False  # 设置处理状态为未激活
        self.recv_buffer_size = 1048576  # 设置接收缓冲区大小为1048576字节，1MB

        # 批量写入数据
        self.batch_size = 10  # 设置批量写入数据的大小为10
        self.data_batch = []  # 创建数据批量列表，用于存储待写入的数据

        # 数据格式设置
        self.fixed_header = b'\xAA'  # 设置固定头部为0xAA
        self.start_time = time.time()  # 记录程序启动时间
        self.last_packet_time = 0  # 记录上次数据包时间
        self.packet_times = deque(maxlen=100)  # 创建一个固定大小的双端队列，用于存储数据包时间戳

        # 性能统计
        self.received_packets = 0  # 记录接收到的数据包数量
        self.processed_packets = 0  # 记录处理过的数据包数量
        self.last_stats_time = time.time()  # 记录上次统计时间
        self.stats_interval = 5.0  # 设置统计间隔为5秒

        # 创建界面
        self.create_widgets()  # 创建界面组件
        self.update_ui()  # 更新界面状态
        self.create_plot_area()  # 创建绘图区域

        # 初始化语音引擎
        self.init_speech_engine()

        # 启动数据处理线程
        self.processing_active = True
        self.processing_thread = threading.Thread(target=self.process_data_thread, daemon=True)
        self.processing_thread.start()

        # 启动GUI消息处理线程
        self.gui_thread = threading.Thread(target=self.process_gui_queue, daemon=True)
        self.gui_thread.start()

        # 启动语音播报线程
        self.speech_queue = queue.Queue()
        self.speech_thread = threading.Thread(target=self.speak_worker, daemon=True)
        self.speech_thread.start()

        # 循环语音提示相关
        self.cycle_active = False
        self.cycle_thread = None
        self.current_stage = 0  # 0: 放松, 1: 握紧, 2: 放松
        self.stage_start_time = None
        self.cycle_timer_id = None  # 新增：初始化cycle_timer_id为None

        # 设置窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def back_to_main(self):
        # 停止所有定时器和线程
        self.processing_active = False  # 停止数据处理线程
        self.stop_cycle_prompt()  # 停止循环提示


        # 销毁当前界面的所有组件
        for widget in self.root.winfo_children():
            widget.destroy()

        # 导入并显示主界面
        from kongzhi import WiFiHardwareController
        WiFiHardwareController(self.root)

    def on_close(self):
        self.processing_active = False
        if self.current_data_file:
            self.flush_data_batch()
            self.current_data_file.close()
        self.root.destroy()

    def generate_colors(self, n):
        colors = []  # 创建颜色列表存储颜色
        for i in range(n):
            h = i / n  # 计算色相值
            s = 0.8  # 设置饱和度
            v = 0.9  # 设置亮度
            r, g, b = colorsys.hsv_to_rgb(h, s, v)  # 将色相、饱和度和亮度转换为RGB颜色
            color = f'#{int(r * 255):02x}{int(g * 255):02x}{int(b * 255):02x}'  # 将RGB颜色转换为十六进制格式
            colors.append(color)  # 将颜色添加到列表中
        return colors  # 返回颜色列表

    def create_plot_area(self):
        self.plot_frame = ttk.LabelFrame(self.root, text="多通道数字波形监测")  # 创建带标题的框架
        self.plot_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=10)
        self.root.grid_rowconfigure(3, weight=2)
        self.root.grid_columnconfigure(0, weight=1)

        self.plot_frame.rowconfigure(0, weight=1)
        self.plot_frame.columnconfigure(0, weight=1)
        self.plot_frame.columnconfigure(1, weight=0)

        self.canvas_frame = tk.Frame(self.plot_frame)
        self.canvas_frame.grid(row=0, column=0, sticky="nsew")
        self.plot_frame.grid_rowconfigure(0, weight=1)
        self.plot_frame.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(self.canvas_frame, bg='white')  # 不再指定固定宽高
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", self.on_canvas_resize)
        # 新增：绑定双击事件
        self.canvas.bind("<Double-Button-1>", self.on_yaxis_double_click)

        self.legend_frame = ttk.Frame(self.plot_frame)
        self.legend_frame.grid(row=1, column=0, sticky="ew", pady=5)
        self.plot_frame.grid_rowconfigure(1, weight=0)
        self.update_legend()  # 更新图例

        self.root.after(self.plot_update_interval, self.update_plot)  # 设置定时器定期更新波形显示，每20毫秒更新一次绘图

    def on_canvas_resize(self, event):
        self.plot_width = event.width
        self.plot_height = event.height
        self.effective_height = self.plot_height - self.bottom_margin - self.top_margin  # 修改：考虑上边距
        self.y_scale = self.effective_height / (self.y_top_fixed - self.y_bottom_fixed)
        self.update_plot_once()

    def update_legend(self):
        # 清除现有图例
        for widget in self.legend_frame.winfo_children():
            widget.destroy()

        # 创建通道变量列表
        self.channel_vars = []

        # 创建通道复选框
        for i in range(self.num_channels):
            var = tk.BooleanVar(value=self.channel_visibility[i])
            self.channel_vars.append(var)
            row = i // 8
            col = i % 8 + 1  # 列位置+1，为按钮留出空间
            chk = tk.Checkbutton(
                self.legend_frame,
                text=f"通道 {i + 1}",
                variable=var,
                fg=self.colors[i],
                activeforeground=self.colors[i],
                command=lambda idx=i: self.toggle_channel(idx)
            )
            chk.grid(row=row, column=col, padx=5)

        # 创建全选/取消全选按钮
        self.toggle_select_btn = ttk.Button(
            self.legend_frame,
            text="取消全选",
            command=self.toggle_all_channels,
            width=8
        )
        self.toggle_select_btn.grid(row=0, column=0, padx=5, pady=5)

    def toggle_all_channels(self):
        # 检查当前是否所有通道都被选中
        all_selected = all(self.channel_visibility)

        # 根据当前状态切换所有通道的显示状态
        for i in range(self.num_channels):
            self.channel_visibility[i] = not all_selected
            self.channel_vars[i].set(not all_selected)

        # 更新按钮文本
        self.toggle_select_btn.configure(text="全选" if all_selected else "取消全选")

        # 如果图像暂停，更新显示
        if self.plot_paused:
            self.update_plot_once()

    def toggle_channel(self, idx):
        self.channel_visibility[idx] = not self.channel_visibility[idx]  # 切换通道的显示状态
        self.channel_vars[idx].set(self.channel_visibility[idx])  # 更新通道变量的显示状态
        # 发送状态消息到GUI队列
        self.gui_queue.put(("msg", f"通道 {idx + 1} 显示状态: {'显示' if self.channel_visibility[idx] else '隐藏'}",
                            'info'))  # 在消息队列中添加消息，提示通道的显示状态
        if self.plot_paused:  # 如果图像暂停，更新显示
            self.update_plot_once()

    def update_plot_once(self):
        self.canvas.delete("all")  # 清除画布
        if self.plot_paused and self.paused_plot_data is not None:  # 如果图像暂停且暂停数据不为空
            plot_data = self.paused_plot_data  # 使用暂停数据
            time_data = self.paused_time_data  # 使用暂停时间数据
            current_index = self.paused_data_index  # 使用暂停数据索引
        else:
            with self.plot_lock:  # 获取绘图锁
                plot_data = self.plot_data.copy()  # 复制当前波形数据
                time_data = self.time_data.copy()  # 复制当前时间数据
                current_index = self.data_index  # 使用当前数据索引

        # 根据数据时长决定x轴显示范围：
        # 前10秒固定显示 0~10s；10秒后显示 [elapsed-10, elapsed]
        if current_index == 0:
            min_time = 0
            max_time = 10
        else:
            if self.start_timestamp is None:
                self.start_timestamp = time_data[0]
            elapsed_time = time_data[current_index - 1] - self.start_timestamp
            if elapsed_time < 10:  # 如果数据时长小于10秒，显示0-10秒
                min_time = 0
                max_time = 10
            else:
                min_time = elapsed_time - 10  # 如果数据时长大于10秒，则显示[elapsed-10, elapsed]，显示最近10s
                max_time = elapsed_time

        self.draw_grid(min_time, max_time)  # 绘制网络

        if current_index >= 2:  # 绘制波形
            x_coords = np.linspace(self.left_margin, self.plot_width, self.max_data_points)
            for i in range(self.num_channels):
                if not self.channel_visibility[i]:
                    continue
                y_coords = self.top_margin + (self.y_top_fixed - plot_data[i]) * self.y_scale
                points = list(zip(x_coords, y_coords))
                self.canvas.create_line(points[:current_index], fill=self.colors[i], width=1)

    def update_plot(self):
        current_time = time.time() * 1000
        if current_time - self.last_plot_update < self.plot_update_interval:
            self.root.after(1, self.update_plot)
            return
        self.last_plot_update = current_time

        # 只有采集已开始才显示实时波形，否则显示空白或提示
        if not self.collection_started:
            self.canvas.delete("all")
            self.canvas.create_text(self.plot_width // 2, self.plot_height // 2, text="请点击'开始采集'按钮",
                                    fill="#888", font=("Arial", 18))
            self.root.after(self.plot_update_interval, self.update_plot)
            return

        if not self.data_collection_active:
            # 如果数据采集暂停，显示暂停时的数据
            self.canvas.delete("all")
            if self.paused_plot_data is not None:
                plot_data = self.paused_plot_data
                time_data = self.paused_time_data
                current_index = self.paused_data_index

                if current_index == 0:
                    min_time = 0
                    max_time = 10
                else:
                    if self.start_timestamp is None:
                        self.start_timestamp = time_data[0]
                    elapsed_time = time_data[current_index - 1] - self.start_timestamp
                    if elapsed_time < 10:
                        min_time = 0
                        max_time = 10
                    else:
                        min_time = elapsed_time - 10
                        max_time = elapsed_time

                self.draw_grid(min_time, max_time)

                if current_index >= 2:
                    x_coords = np.linspace(self.left_margin, self.plot_width, self.max_data_points)
                    for i in range(self.num_channels):
                        if not self.channel_visibility[i]:
                            continue
                        y_coords = self.top_margin + (self.y_top_fixed - plot_data[i]) * self.y_scale
                        points = list(zip(x_coords, y_coords))
                        self.canvas.create_line(points[:current_index], fill=self.colors[i], width=1)
        else:
            # 正常显示实时数据
            self.canvas.delete("all")
            with self.plot_lock:
                plot_data = self.plot_data.copy()
                time_data = self.time_data.copy()
                current_index = self.data_index

            if current_index == 0:
                min_time = 0
                max_time = 10
            else:
                if self.start_timestamp is None:
                    self.start_timestamp = time_data[0]
                elapsed_time = time_data[current_index - 1] - self.start_timestamp
                if elapsed_time < 10:
                    min_time = 0
                    max_time = 10
                else:
                    min_time = elapsed_time - 10
                    max_time = elapsed_time

            self.draw_grid(min_time, max_time)

            if current_index >= 2:
                x_coords = np.linspace(self.left_margin, self.plot_width, self.max_data_points)
                for i in range(self.num_channels):
                    if not self.channel_visibility[i]:
                        continue
                    y_coords = self.top_margin + (self.y_top_fixed - plot_data[i]) * self.y_scale
                    points = list(zip(x_coords, y_coords))
                    self.canvas.create_line(points[:current_index], fill=self.colors[i], width=1)

        self.root.after(self.plot_update_interval, self.update_plot)

    def draw_grid(self, min_time, max_time):
        y_bottom = self.y_bottom_fixed
        y_top = self.y_top_fixed
        # 均匀分布y轴刻度线：最大-0之间3条，最小-0之间3条
        y_ticks = [y_top]
        for i in range(1, 4):
            y_ticks.append(y_top - i * (y_top / 3))
        for i in range(1, 4):
            y_ticks.append(y_bottom + i * (-y_bottom / 3))
        y_ticks.append(y_bottom)
        y_ticks = sorted(set([round(v, 6) for v in y_ticks]), reverse=True)
        # 新增：记录最大/最小刻度的像素位置和数值
        self._yaxis_max_marker = None
        self._yaxis_min_marker = None
        # 绘制y轴刻度线
        for value in y_ticks:
            y_pos = self.top_margin + (y_top - value) * self.y_scale
            self.canvas.create_line(self.left_margin, y_pos, self.plot_width, y_pos, fill='#EEE')
            self.canvas.create_text(self.left_margin - 5, y_pos, text=f"{value:.3f}", anchor="e", fill="#666")
            # 记录最大/最小刻度
            if abs(value - y_top) < 1e-6:
                self._yaxis_max_marker = (y_pos, value)
            if abs(value - y_bottom) < 1e-6:
                self._yaxis_min_marker = (y_pos, value)
        # 0刻度线不再加粗，样式与其他刻度线一致
        # y0 = self.top_margin + (y_top - 0) * self.y_scale
        # self.canvas.create_line(self.left_margin, y0, self.plot_width, y0, fill='#888', width=2)
        # self.canvas.create_text(self.left_margin - 5, y0, text="0.000", anchor="e", fill="#666")
        # x轴
        x_axis_y = self.top_margin + self.effective_height
        self.canvas.create_line(self.left_margin, x_axis_y, self.plot_width, x_axis_y, fill='#EEE')
        # x轴刻度线
        time_range = max_time - min_time
        step = time_range / 10
        for i in range(11):
            t = min_time + i * step
            # 让最右侧刻度向左偏移2%宽度
            if i == 10:
                x = self.left_margin + 0.98 * (self.plot_width - self.left_margin)
            else:
                x = self.left_margin + (i / 10) * (self.plot_width - self.left_margin)
            # 纵向刻度线（延伸到顶部，样式与y轴横线一致）
            self.canvas.create_line(x, self.top_margin, x, x_axis_y, fill='#EEE')
            # 刻度文字
            self.canvas.create_text(x, x_axis_y + self.bottom_margin / 2, text=f"{t:.1f}s", anchor="n", fill="#666")

    def create_widgets(self):
        # 删除mode_frame及相关控件
        self.channel_frame = ttk.LabelFrame(self.root, text="通道设置")  # 通道设置区域
        self.channel_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=2)
        self.root.grid_rowconfigure(0, weight=0)
        self.root.grid_columnconfigure(0, weight=1)
        self.channel_frame.columnconfigure(0, weight=0)
        self.channel_frame.columnconfigure(1, weight=0)
        self.channel_frame.columnconfigure(2, weight=0)
        ttk.Label(self.channel_frame, text="通道数:").grid(row=0, column=0, padx=5, pady=5)  # 通道数标签
        self.channel_var = tk.IntVar(value=self.num_channels)  # 通道数变量
        self.channel_spinbox = tk.Spinbox(self.channel_frame, from_=1, to=16, width=5,
                                          textvariable=self.channel_var)  # 通道数输入框
        self.channel_spinbox.grid(row=0, column=1, padx=5, pady=5)
        self.channel_confirm_btn = ttk.Button(self.channel_frame, text="确定", command=self.change_channel_count)
        self.channel_confirm_btn.grid(row=0, column=2, padx=5, pady=5)
        self.channel_confirm_btn = ttk.Button(self.channel_frame, text="跳转至控制页面", command=self.back_to_main)
        self.channel_confirm_btn.grid(row=0, column=4, padx=5, pady=5)

        self.conn_frame = ttk.LabelFrame(self.root, text="连接设置")  # 连接设置区域
        self.conn_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=2)
        self.root.grid_rowconfigure(1, weight=0)
        self.conn_frame.columnconfigure(0, weight=0)
        self.conn_frame.columnconfigure(1, weight=0)
        self.conn_frame.columnconfigure(2, weight=0)
        self.conn_frame.columnconfigure(3, weight=0)
        self.conn_frame.columnconfigure(4, weight=0)
        self.conn_frame.columnconfigure(5, weight=0)
        self.conn_frame.columnconfigure(6, weight=0)
        ttk.Label(self.conn_frame, text="IP地址:").grid(row=0, column=0)  # IP地址标签
        self.ip_entry = ttk.Entry(self.conn_frame, width=15)  # IP地址输入框
        self.ip_entry.insert(0, self.ip_address)  # 使用从配置文件加载的IP地址
        self.ip_entry.grid(row=0, column=1)  # 将IP地址输入框放置在连接设置区域
        ttk.Label(self.conn_frame, text="端口:").grid(row=0, column=2)  # 端口标签
        self.port_entry = ttk.Entry(self.conn_frame, width=8)  # 端口输入框
        self.port_entry.insert(0, str(self.port_number))  # 使用从配置文件加载的端口号
        self.port_entry.grid(row=0, column=3)  # 将端口输入框放置在连接设置区域
        # 功能按钮
        self.conn_btn = ttk.Button(self.conn_frame, text="启动连接", command=self.toggle_connection)  # 启动连接按钮
        self.conn_btn.grid(row=0, column=4, padx=10)  # 将启动连接按钮放置在连接设置区域
        self.save_conn_btn = ttk.Button(self.conn_frame, text="保存连接参数", command=self.on_save_connection_params)
        self.save_conn_btn.grid(row=0, column=5, padx=10)
        self.pause_btn = ttk.Button(self.conn_frame, text="开始采集", command=self.toggle_data_collection)
        self.pause_btn.grid(row=0, column=6, padx=10)

        # 新增：实验参数设置框架
        self.experiment_frame = ttk.LabelFrame(self.root, text="实验参数设置")
        self.experiment_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=2)
        self.root.grid_rowconfigure(2, weight=0)
        self.experiment_frame.columnconfigure(0, weight=0)
        self.experiment_frame.columnconfigure(1, weight=0)
        self.experiment_frame.columnconfigure(2, weight=0)
        self.experiment_frame.columnconfigure(3, weight=0)
        self.experiment_frame.columnconfigure(4, weight=0)
        self.experiment_frame.columnconfigure(5, weight=0)

        # 收缩次数设置
        ttk.Label(self.experiment_frame, text="收缩次数:").grid(row=0, column=0, padx=5, pady=5)
        self.cycle_total_var = tk.IntVar(value=self.cycle_total)
        self.cycle_total_spinbox = tk.Spinbox(self.experiment_frame, from_=1, to=50, width=5,
                                              textvariable=self.cycle_total_var)
        self.cycle_total_spinbox.grid(row=0, column=1, padx=5, pady=5)

        # 收缩时间设置
        ttk.Label(self.experiment_frame, text="收缩时间(秒):").grid(row=0, column=2, padx=5, pady=5)
        self.contraction_time_var = tk.DoubleVar(value=self.contraction_time)
        self.contraction_time_spinbox = tk.Spinbox(self.experiment_frame, from_=1.0, to=60.0, increment=0.5, width=8,
                                                   textvariable=self.contraction_time_var)
        self.contraction_time_spinbox.grid(row=0, column=3, padx=5, pady=5)

        # 放松时间设置
        ttk.Label(self.experiment_frame, text="放松时间(秒):").grid(row=0, column=4, padx=5, pady=5)
        self.relaxation_time_var = tk.DoubleVar(value=self.relaxation_time)
        self.relaxation_time_spinbox = tk.Spinbox(self.experiment_frame, from_=1.0, to=60.0, increment=0.5, width=8,
                                                  textvariable=self.relaxation_time_var)
        self.relaxation_time_spinbox.grid(row=0, column=5, padx=5, pady=5)

        # 保存实验参数按钮
        self.save_params_btn = ttk.Button(self.experiment_frame, text="保存参数", command=self.save_experiment_params)
        self.save_params_btn.grid(row=0, column=6, padx=10, pady=5)
        # 新增：语音播报开关Checkbutton
        self.voice_broadcast_chk = tk.Checkbutton(self.experiment_frame, text="语音播报",
                                                  variable=self.voice_broadcast_var)
        self.voice_broadcast_chk.grid(row=0, column=7, padx=10, pady=5)

        self.msg_frame = ttk.LabelFrame(self.root, text="消息窗口")  # 消息窗口区域
        self.msg_frame.grid(row=4, column=0, sticky="nsew", padx=10, pady=5)
        self.root.grid_rowconfigure(4, weight=1)
        self.msg_text = tk.Text(self.msg_frame, height=2)  # 消息文本框
        self.msg_text.pack(fill="both", expand=True)
        self.msg_text.tag_config('send', foreground='blue')  # 发送消息标签
        self.msg_text.tag_config('recv', foreground='green')  # 接收消息标签
        self.msg_text.tag_config('error', foreground='red')  # 错误消息标签
        self.msg_text.tag_config('info', foreground='purple')  # 信息消息标签

    def change_channel_count(self):
        if self.connected:  # 检查设备是否已连接，如果已连接，则无法更改通道数
            messagebox.showwarning("警告", "连接建立后无法更改通道数！")
            self.channel_var.set(self.num_channels)
            return
        try:  # 通道验证和更新
            new_count = int(self.channel_spinbox.get())
            if new_count < 1 or new_count > 16:
                raise ValueError("通道数必须在1到16之间")
            self.num_channels = new_count  # 更新通道数
            self.save_num_channels()  # 新增：保存到配置文件
            self.group_data_size = self.num_channels * 4  # 更新组数据大小
            self.packet_size = 2 + 5 * self.group_data_size  # 更新包大小
            self.plot_data = np.zeros((self.num_channels, self.max_data_points), dtype=np.float32)  # 更新波形数据
            self.channel_visibility = [True] * self.num_channels  # 更新通道可见性
            self.update_legend()  # 更新图例
            self.gui_queue.put(("msg", f"通道数更改为 {self.num_channels}", 'info'))  # 在消息队列中添加消息，提示通道数已更改
        except Exception as e:
            messagebox.showerror("错误", str(e))  # 显示错误消息

    def toggle_pause_plot(self):
        self.plot_paused = not self.plot_paused  # 切换图像暂停状态
        if self.plot_paused:  # 如果图像暂停
            self.pause_btn.config(text="恢复图像")  # 更新按钮文本
            with self.plot_lock:  # 获取绘图锁
                self.paused_plot_data = self.plot_data.copy()  # 复制当前波形数据
                self.paused_time_data = self.time_data.copy()  # 复制当前时间数据
                self.paused_data_index = self.data_index  # 复制当前数据索引
            self.gui_queue.put(("msg", "图像显示已暂停", 'info'))  # 在消息队列中添加消息，提示图像显示已暂停
            self.update_plot_once()  # 更新波形显示
        else:  # 如果图像未暂停
            self.pause_btn.config(text="暂停图像")  # 更新按钮文本
            self.paused_plot_data = None  # 清空暂停数据
            self.paused_time_data = None  # 清空暂停时间数据
            self.paused_data_index = 0  # 清空暂停数据索引
            self.gui_queue.put(("msg", "图像显示已恢复", 'info'))  # 在消息队列中添加消息，提示图像显示已恢复

    def switch_mode(self):
        pass  # 删除工作模式切换功能

    def update_ui(self):
        state = "normal" if not self.connected else "disabled"  # 根据连接状态更新按钮状态
        # 删除主机/客户机按钮状态更新
        self.port_entry.config(state=state)  # 更新端口输入框状态
        self.conn_btn.config(text="断开连接" if self.connected else "启动连接")  # 更新连接按钮文本

    def toggle_connection(self):
        if not self.connected:  # 如果未连接
            self.start_connection()  # 启动连接
        else:  # 如果已连接
            self.stop_connection()  # 停止连接

    def start_connection(self):
        try:
            # 新增：连接前自动保存连接参数（静默模式）
            self.save_connection_params(silent=True)

            port = int(self.port_entry.get())  # 获取端口号
            if not (0 < port <= 65535):  # 验证端口号有效性
                raise ValueError("端口号无效")
            self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)  # 创建TCP套接字
            self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_RCVBUF, self.recv_buffer_size)  # 设置接收缓冲区大小
            self.sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)  # 设置TCP_NODELAY选项，禁用Nagle算法
            self.processing_active = True  # 设置数据处理线程状态
            threading.Thread(target=self.process_data_thread, daemon=True).start()  # 启动数据处理线程
            self.gui_queue.put(("msg", "数据处理线程已启动", 'info'))  # 在消息队列中添加消息，提示数据处理线程已启动
            # 仅主机模式
            self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.sock.bind(('0.0.0.0', port))
            self.sock.listen(1)
            threading.Thread(target=self.accept_connections, daemon=True).start()
            self.gui_queue.put(("status", "等待连接..."))
            self.connected = True  # 更新连接状态
        except Exception as e:  # 捕获异常
            messagebox.showerror("连接错误", str(e))  # 显示错误消息
            self.stop_connection()  # 停止连接
        finally:
            self.update_ui()  # 更新UI

    def stop_connection(self):
        try:
            if self.client_sock:  # 关闭客户端套接字
                self.client_sock.close()
                self.client_sock = None
            if self.sock:  # 关闭服务器套接字
                self.sock.close()
                self.sock = None
            self.connected = False  # 更新连接状态
            self.gui_queue.put(("status", "已断开"))  # 在消息队列中添加消息，提示已断开
        except Exception as e:  # 捕获异常
            messagebox.showerror("断开错误", str(e))  # 显示错误消息
        finally:
            self.update_ui()  # 更新UI

    def accept_connections(self):
        try:
            self.client_sock, addr = self.sock.accept()  # 接受客户端连接
            self.client_sock.setsockopt(socket.SOL_SOCKET, socket.SO_RCVBUF, self.recv_buffer_size)  # 设置接收缓冲区大小
            self.client_sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)  # 设置TCP_NODELAY选项，禁用Nagle算法
            self.gui_queue.put(("msg", f"已连接客户端: {addr}", 'recv'))  # 在消息队列中添加消息，提示已连接客户端
            self.speak_async("设备连接成功")
            # 新增：连接成功后，采集未开始，按钮显示"开始采集"
            self.collection_started = False
            self.data_collection_active = False
            self.pause_btn.config(text="开始采集")
            self.gui_queue.put(("msg", "设备连接成功，请点击'开始采集'按钮", 'info'))
            threading.Thread(target=self.receive_data, args=(self.client_sock,), daemon=True).start()  # 启动数据接收线程
            self.gui_queue.put(("status", "已连接"))  # 在消息队列中添加消息，提示已连接
        except Exception as e:  # 捕获异常
            if self.connected:  # 检查设备是否已连接，如果已连接，则无法接收数据
                self.gui_queue.put(("msg", f"连接异常: {str(e)}", 'error'))  # 在消息队列中添加消息，提示连接异常

    def receive_data(self, sock):
        data_buffer = b''  # 数据缓冲区
        while self.connected and sock:  # 数据接收循环，只要连接保持且套接字存在，就持续接收数据
            try:
                chunk = sock.recv(self.recv_buffer_size)  # 接收数据块
                if not chunk:
                    self.handle_disconnect()
                    break
                data_buffer += chunk  # 将接收到的数据块添加到数据缓冲区
                while len(data_buffer) >= self.packet_size:  # 如果数据缓冲区中的数据长度大于等于包大小
                    if data_buffer[0:1] != self.fixed_header:  # 检查数据缓冲区的第一个字节是否为固定头
                        idx = data_buffer.find(self.fixed_header)  # 查找固定头在数据缓冲区中的位置
                        if idx == -1:  # 如果固定头未找到
                            data_buffer = b''  # 清空数据缓冲区
                            break
                        else:
                            data_buffer = data_buffer[idx:]  # 更新数据缓冲区，只保留找到的固定头及其后的数据
                            if len(data_buffer) < self.packet_size:  # 如果数据缓冲区中的数据长度小于包大小
                                break
                    packet_data = data_buffer[:self.packet_size]  # 提取包数据
                    data_buffer = data_buffer[self.packet_size:]  # 更新数据缓冲区，只保留找到的固定头及其后的数据
                    recv_time = time.time() - self.start_time  # 计算接收时间
                    self.received_packets += 1  # 增加接收包计数
                    try:
                        self.data_queue.put_nowait((recv_time, packet_data))  # 将接收到的数据包添加到数据队列
                    except queue.Full:  # 如果数据队列已满
                        try:
                            self.data_queue.get_nowait()  # 从数据队列中移除一个数据包
                            self.data_queue.put_nowait((recv_time, packet_data))  # 将接收到的数据包添加到数据队列
                        except:
                            pass
            except ConnectionResetError:  # 如果连接被重置
                self.gui_queue.put(("msg", "连接已重置", 'error'))  # 在消息队列中添加消息，提示连接已重置
                break
            except Exception as e:  # 捕获异常
                self.handle_network_error(e)  # 处理网络错误
                break

    def process_data_thread(self):
        while self.processing_active:
            try:
                try:
                    current_time, binary_data = self.data_queue.get(timeout=0.1)
                except queue.Empty:
                    continue
                if len(binary_data) == self.packet_size:
                    packet_seq = binary_data[1]
                    payload = binary_data[2:]
                    for group_index in range(5):
                        start = group_index * self.group_data_size
                        group_bytes = payload[start:start + self.group_data_size]
                        values = struct.unpack(f'{self.num_channels}f', group_bytes)
                        group_timestamp = current_time - (5 - 1 - group_index) * 0.001

                        # 只在数据采集激活时保存数据
                        if self.data_collection_active:
                            self.data_batch.append((group_timestamp, np.array(values, dtype=np.float32), group_index))
                            if len(self.data_batch) >= self.batch_size:
                                self.flush_data_batch()

                        with self.plot_lock:
                            if self.data_index < self.max_data_points:
                                for i, val in enumerate(values):
                                    self.plot_data[i, self.data_index] = val
                                self.time_data[self.data_index] = group_timestamp
                                self.data_index += 1
                            else:
                                self.plot_data = np.roll(self.plot_data, -1, axis=1)
                                self.time_data = np.roll(self.time_data, -1)
                                for i, val in enumerate(values):
                                    self.plot_data[i, -1] = val
                                self.time_data[-1] = group_timestamp
                        self.processed_packets += 1
                else:
                    err_msg = f"数据包长度异常: 应为 {self.packet_size} 字节, 实际 {len(binary_data)} 字节"
                    self.gui_queue.put(("msg", err_msg, 'error'))
            except Exception as e:
                err_msg = f"数据处理错误: {str(e)}"
                self.gui_queue.put(("msg", err_msg, 'error'))

    def flush_data_batch(self):
        if not self.data_batch or not self.current_data_file:
            return
        try:
            with self.save_lock:
                buffer = io.StringIO()
                for timestamp, values, group in self.data_batch:
                    values_str = ','.join(f"{v:.3f}" for v in values)
                    buffer.write(f"{timestamp},{group},{values_str}\n")
                self.current_data_file.write(buffer.getvalue())
                if self.data_counter % 5 == 0:
                    self.current_data_file.flush()
                    os.fsync(self.current_data_file.fileno())
                self.data_counter += 1
            self.data_batch = []
        except Exception as e:
            self.gui_queue.put(("msg", f"文件写入错误: {e}", 'error'))

    def process_gui_queue(self):
        try:
            while not self.gui_queue.empty():  # 消息队列处理
                item = self.gui_queue.get()  # 获取消息队列中的消息
                if item[0] == "msg":  # 消息处理
                    tag = item[2] if len(item) > 2 else None  # 获取消息标签
                    self.msg_text.insert("end", item[1] + "\n", tag)  # 在消息文本框中插入消息
                    self.msg_text.see("end")  # 滚动到消息文本框的末尾
                elif item[0] == "status":  # 状态更新
                    self.root.title(f"肌电信号采集 - {item[1]}")  # 更新窗口标题
        except Exception as e:  # 捕获异常
            print(f"GUI更新错误: {str(e)}")  # 打印错误信息
        self.root.after(1, self.process_gui_queue)

    def handle_disconnect(self):
        self.gui_queue.put(("msg", "连接已断开", 'recv'))
        self.root.after(0, self.stop_connection)

    def handle_network_error(self, error):
        self.gui_queue.put(("msg", f"网络错误: {str(error)}", 'recv'))
        self.root.after(0, self.stop_connection)

    def create_new_data_file(self):
        """创建新的数据文件，保存在原始数据文件夹下"""
        if self.current_data_file:
            self.flush_data_batch()  # 确保之前的数据被写入
            self.current_data_file.close()
        # 确保原始数据文件夹存在
        raw_data_dir = os.path.join(self.program_dir, '原始数据')
        os.makedirs(raw_data_dir, exist_ok=True)
        filename = f"data_{time.strftime('%Y%m%d_%H%M%S')}.csv"
        file_path = os.path.join(raw_data_dir, filename)
        self.current_data_file = open(file_path, "a", buffering=8192)
        self.current_data_file.write("timestamp,group,values\n")
        self.data_counter = 0
        self.gui_queue.put(("msg", f"创建新的数据文件: {file_path}", 'info'))

    def toggle_data_collection(self):
        """切换数据采集状态"""
        # 未开始采集时，点击为"开始采集"
        if not self.collection_started:
            self.collection_started = True
            self.data_collection_active = True
            self.pause_btn.config(text="暂停采集")
            self.create_new_data_file()
            # 清空图像数据
            with self.plot_lock:
                self.plot_data = np.zeros((self.num_channels, self.max_data_points), dtype=np.float32)
                self.time_data = np.zeros(self.max_data_points, dtype=np.float64)
                self.data_index = 0
                self.start_timestamp = None
            # 新增：清空语音时间点记录
            with self.voice_timestamp_lock:
                self.voice_timestamps = []
            self.gui_queue.put(("msg", "数据采集已开始", 'info'))
            if self.voice_broadcast_enabled:
                self.speak_async("实验开始，请放松")
                self.start_cycle_prompt()  # 只有在语音播报开启时才启动循环语音提示
            else:
                self.speak_async("实验开始")
                self.gui_queue.put(("msg", "语音播报已关闭，请手动控制实验进程", 'info'))
            return
        # 已采集时，切换暂停/继续
        self.data_collection_active = not self.data_collection_active
        if self.data_collection_active:
            self.pause_btn.config(text="暂停采集")
            self.create_new_data_file()
            # 清空图像数据
            with self.plot_lock:
                self.plot_data = np.zeros((self.num_channels, self.max_data_points), dtype=np.float32)
                self.time_data = np.zeros(self.max_data_points, dtype=np.float64)
                self.data_index = 0
                self.start_timestamp = None
            # 新增：清空语音时间点记录
            with self.voice_timestamp_lock:
                self.voice_timestamps = []
            self.gui_queue.put(("msg", "数据采集已继续", 'info'))
            if self.voice_broadcast_enabled:
                self.speak_async("实验开始，请放松")
                self.start_cycle_prompt()  # 只有在语音播报开启时才启动循环语音提示
            else:
                self.speak_async("实验开始")
                self.gui_queue.put(("msg", "语音播报已关闭，请手动控制实验进程", 'info'))
        else:
            self.stop_cycle_prompt()  # 新增：采集暂停时停止循环
            self.pause_btn.config(text="继续采集")
            self.flush_data_batch()  # 保存当前数据
            # 保存当前图像状态
            with self.plot_lock:
                self.paused_plot_data = self.plot_data.copy()
                self.paused_time_data = self.time_data.copy()
                self.paused_data_index = self.data_index
            self.gui_queue.put(("msg", "数据采集已暂停", 'info'))
            if self.current_data_file:
                self.current_data_file.close()
                # 新增：弹出对话框让用户输入新文件名
                import tkinter.simpledialog
                old_filename = self.current_data_file.name
                new_filename = tkinter.simpledialog.askstring("重命名数据文件", "请输入新的数据文件名（无需扩展名）:")
                if new_filename:
                    if not new_filename.endswith('.csv'):
                        new_filename += '.csv'
                    # 新文件名也要在原始数据文件夹下
                    raw_data_dir = os.path.join(self.program_dir, '原始数据')
                    new_file_path = os.path.join(raw_data_dir, new_filename)
                    try:
                        os.rename(old_filename, new_file_path)
                        self.gui_queue.put(("msg", f"数据文件已重命名为: {new_file_path}", 'info'))
                        # 新增：弹出波形显示窗口，传递语音时间点
                        with self.voice_timestamp_lock:
                            voice_timestamps = self.voice_timestamps.copy()
                        self.show_waveform_window(new_file_path, voice_timestamps)
                        # 新增：清空语音时间点记录，为下次采集做准备
                        with self.voice_timestamp_lock:
                            self.voice_timestamps = []
                    except Exception as e:
                        self.gui_queue.put(("msg", f"重命名失败: {e}", 'error'))
                self.current_data_file = None

    def show_waveform_window(self, filename, voice_timestamps=None):
        import csv
        import numpy as np
        # 新增: 导入scipy.signal
        from scipy.signal import butter, filtfilt, iirnotch
        import os
        # 确保读取的是原始数据文件夹下的文件
        if not os.path.isabs(filename):
            raw_data_dir = os.path.join(self.program_dir, '原始数据')
            filename = os.path.join(raw_data_dir, filename)
        # 读取数据
        try:
            with open(filename, 'r') as f:
                reader = csv.reader(f)
                header = next(reader)
                data = [row for row in reader if len(row) >= 2 + self.num_channels]
            if not data:
                messagebox.showerror("错误", "数据文件为空或格式错误！")
                return
            times = np.array([float(row[0]) for row in data])
            values = np.array([[float(v) for v in row[2:2 + self.num_channels]] for row in data])
            raw_values = values.copy()  # 保存原始数据
        except Exception as e:
            messagebox.showerror("错误", f"读取数据文件失败: {e}")
            return

        # 独立通道数和颜色
        num_channels = values.shape[1]
        preview_colors = self.generate_colors(num_channels)

        # 独立y轴范围（预览界面独立）
        y_top, y_bottom = self.load_yaxis_range(mode="preview")

        # 新增: 对values进行滤波处理
        sample_rate = 1000  # 采样率
        nyq = 0.5 * sample_rate
        # 1. 50Hz及其倍数陷波（50, 100, 150, 200, 250, 300, 350, 400, 450Hz）
        notch_freqs = [50 * i for i in range(1, int(500 / 50) + 1)]
        Q = 30  # 品质因数
        filtered_values = values.copy()
        for freq in notch_freqs:
            b_notch, a_notch = iirnotch(w0=freq / nyq, Q=Q)
            for ch in range(filtered_values.shape[1]):
                filtered_values[:, ch] = filtfilt(b_notch, a_notch, filtered_values[:, ch])
        # 2. 10-500Hz带通滤波
        low = 10 / nyq
        high = 499 / nyq
        b_band, a_band = butter(N=4, Wn=[low, high], btype='band')
        for ch in range(filtered_values.shape[1]):
            filtered_values[:, ch] = filtfilt(b_band, a_band, filtered_values[:, ch])
        values = filtered_values

        # 新增: 保存滤波后的数据到 preprocess 文件夹下
        import os
        preprocess_dir = os.path.join(self.program_dir, '滤波后数据')
        os.makedirs(preprocess_dir, exist_ok=True)
        save_path = os.path.join(preprocess_dir, os.path.basename(filename))
        try:
            with open(save_path, 'w', newline='') as f:
                writer = csv.writer(f)
                # 写表头
                writer.writerow(header)
                # 写数据
                for i in range(len(times)):
                    row = [f"{times[i]:.6f}", data[i][1]]  # timestamp, group
                    row += [f"{v:.3f}" for v in values[i]]
                    writer.writerow(row)
        except Exception as e:
            messagebox.showerror("错误", f"保存滤波后数据失败: {e}")

        # 创建新窗口
        win = tk.Toplevel(self.root)
        win.title(f"数据波形预览 - {filename}")
        win.geometry("1200x650")
        # 画布
        canvas = tk.Canvas(win, bg='white', height=500)
        canvas.pack(fill="both", expand=True, side="top")
        # 滑块区域
        slider_frame = ttk.Frame(win)
        slider_frame.pack(fill="x", side="top")
        # 绘图参数
        left_margin = 45
        bottom_margin = 30
        sample_rate = 1000
        window_duration = 15.0
        t_min = times[0]
        t_max = times[-1]
        max_start = max(0.0, t_max - t_min - window_duration)
        slider_var = tk.DoubleVar(value=0.0)
        # 滑块
        slider = ttk.Scale(slider_frame, from_=0.0, to=max_start, orient="horizontal", variable=slider_var,
                           command=lambda v: redraw(), length=900)
        slider.pack(fill="x", padx=20, pady=5)
        slider_label = ttk.Label(slider_frame, text="起始时间: 0.0s")
        slider_label.pack(side="left", padx=10)

        # y轴刻度交互相关变量
        _yaxis_max_marker = [None]  # [像素位置, 数值]
        _yaxis_min_marker = [None]
        _yaxis_entry = [None]  # 只允许一个输入框

        def y_scale(h):
            return (h - bottom_margin - 20) / (y_top - y_bottom)

        # 新增: 保存原始header和data、通道小数位数
        original_header = header
        original_data = data
        channel_decimal_places = []
        if data:
            for v in data[0][2:]:
                if '.' in v:
                    channel_decimal_places.append(len(v.split('.')[-1]))
                else:
                    channel_decimal_places.append(0)

        # 区间标记相关变量
        contraction_ranges = []  # 存储所有区间，每个为[start, end]
        dragging = {'range_idx': None, 'endpoint': None, 'offset': 0.0, 'mode': None}  # 拖动状态，增加mode
        pending_start = [None]  # 临时存储未配对的起始点
        marker_radius = 8  # 可点击/拖动的像素半径

        # 新增：根据语音时间点自动生成区间标记
        if voice_timestamps:
            auto_ranges = []
            i = 0
            while i < len(voice_timestamps) - 1:
                current_time, current_text = voice_timestamps[i]
                next_time, next_text = voice_timestamps[i + 1]

                # 寻找"请握紧"和"请放松"的配对
                if current_text == "请握紧" and next_text == "请放松":
                    start_time = current_time
                    end_time = next_time
                    # 确保时间点在数据范围内
                    if start_time >= t_min and end_time <= t_max and start_time < end_time:
                        auto_ranges.append([start_time, end_time])
                    i += 2  # 跳过已处理的配对
                else:
                    i += 1  # 移动到下一个时间点
            contraction_ranges.extend(auto_ranges)
            # 在消息窗口显示自动生成的区间信息
            if auto_ranges:
                # 在消息窗口中显示信息
                self.gui_queue.put(("msg", f"根据语音播报自动生成了 {len(auto_ranges)} 个收缩区间", 'info'))

        # 支持端点和区间体检测
        def find_near_endpoint_or_body(event):
            w = canvas.winfo_width()
            start = slider_var.get()
            duration = window_duration
            x = event.x
            for idx, (s, e) in enumerate(contraction_ranges):
                # 端点检测
                for endpoint, t in [('start', s), ('end', e)]:
                    mx = left_margin + (w - left_margin - 10) * ((t - t_min - start) / duration)
                    if abs(mx - x) < marker_radius:
                        return idx, endpoint, 'endpoint'
                # 区间体检测（不在端点附近，且在区间内）
                sx = left_margin + (w - left_margin - 10) * ((s - t_min - start) / duration)
                ex = left_margin + (w - left_margin - 10) * ((e - t_min - start) / duration)
                if sx + marker_radius < x < ex - marker_radius:
                    return idx, None, 'body'
            return None, None, None

        # 保存按钮
        def save_marked_segments():
            if values is None or times is None or not contraction_ranges:
                messagebox.showinfo("提示", "没有可保存的标记区间！")
                return
            if original_header is None or original_data is None:
                messagebox.showerror("错误", "未找到原始数据，无法保存！")
                return
            if channel_decimal_places is None:
                messagebox.showerror("错误", "未检测到原始小数位数，无法保存！")
                return
            # 创建输出文件夹 - 修复：使用程序目录而不是__file__路径
            out_dir = os.path.join(self.program_dir, "收缩标记数据")
            os.makedirs(out_dir, exist_ok=True)
            # 原文件名（不含路径和扩展名）
            src_path = filename
            base_name = os.path.splitext(os.path.basename(src_path))[0]
            # 按区间起始时间排序
            sorted_ranges = sorted(contraction_ranges, key=lambda x: x[0])
            # 导出每个区间
            for idx, (s, e) in enumerate(sorted_ranges):
                # 找到区间内的索引
                seg_idx = np.where((times >= s) & (times <= e))[0]
                if len(seg_idx) == 0:
                    continue
                # 取原始行，替换通道数据为滤波后
                out_rows = []
                for i in seg_idx:
                    row = list(original_data[i])
                    # 替换通道数据（假设通道数据从第3列开始）
                    for ch in range(values.shape[1]):
                        dec = channel_decimal_places[ch] if ch < len(channel_decimal_places) else 6
                        fmt = f"{{:.{dec}f}}"
                        row[2 + ch] = fmt.format(values[i, ch])
                    out_rows.append(row)
                # 文件名
                out_name = f"{base_name}-第{idx + 1}次收缩.csv"
                out_path = os.path.join(out_dir, out_name)
                # 写入
                with open(out_path, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(original_header)
                    writer.writerows(out_rows)
            messagebox.showinfo("保存成功", f"已导出{len(sorted_ranges)}个区间到\n{out_dir}")

        # 新增：滤波显示开关
        show_filtered = tk.BooleanVar(value=True)

        def toggle_filter():
            show_filtered.set(not show_filtered.get())
            redraw()
            filter_btn.config(text="显示原始波形" if show_filtered.get() else "显示滤波后波形")

        # 鼠标按下事件，支持整体拖动
        def add_or_drag_marker(event):
            w = canvas.winfo_width()
            h = canvas.winfo_height()
            x = event.x
            y = event.y
            if x < left_margin or y < 20 or y > h - bottom_margin:
                return
            idx, endpoint, mode = find_near_endpoint_or_body(event)
            if idx is not None:
                dragging['range_idx'] = idx
                dragging['endpoint'] = endpoint
                dragging['mode'] = mode
                dragging['offset'] = 0.0
                # 记录拖动时鼠标与区间起点的距离（用于整体拖动）
                if mode == 'body':
                    s, e = contraction_ranges[idx]
                    sx = left_margin + (w - left_margin - 10) * ((s - t_min - slider_var.get()) / window_duration)
                    dragging['offset'] = x - sx
                return
            # 新增区间逻辑不变
            start = slider_var.get()
            duration = window_duration
            t = t_min + start + (x - left_margin) / (w - left_margin - 10) * duration
            t = max(t_min, min(t, t_max))
            t = round(t, 3)
            if pending_start[0] is None:
                pending_start[0] = t
            else:
                s, e = pending_start[0], t
                if s > e:
                    s, e = e, s
                if not any(abs(s - r[0]) < 0.002 and abs(e - r[1]) < 0.002 for r in contraction_ranges):
                    contraction_ranges.append([s, e])
                pending_start[0] = None
            redraw()

        # 鼠标拖动事件，支持整体拖动
        def on_mouse_move(event):
            if dragging['range_idx'] is not None:
                w = canvas.winfo_width()
                start = slider_var.get()
                duration = window_duration
                x = event.x
                t = t_min + start + (x - left_margin) / (w - left_margin - 10) * duration
                t = max(t_min, min(t, t_max))
                t = round(t, 3)
                idx = dragging['range_idx']
                mode = dragging['mode']
                s, e = contraction_ranges[idx]
                if mode == 'endpoint':
                    endpoint = dragging['endpoint']
                    if endpoint == 'start':
                        if t < e:
                            contraction_ranges[idx][0] = t
                    elif endpoint == 'end':
                        if t > s:
                            contraction_ranges[idx][1] = t
                elif mode == 'body':
                    # 计算区间长度
                    length = e - s
                    # 鼠标相对区间起点的偏移
                    sx = left_margin + (w - left_margin - 10) * ((s - t_min - start) / duration)
                    # 新的起点
                    new_s = t_min + start + (x - dragging['offset'] - left_margin) / (w - left_margin - 10) * duration
                    new_s = max(t_min, min(new_s, t_max - length))
                    new_e = new_s + length
                    # 边界保护
                    if new_s < t_min:
                        new_s = t_min
                        new_e = new_s + length
                    if new_e > t_max:
                        new_e = t_max
                        new_s = new_e - length
                    contraction_ranges[idx][0] = round(new_s, 3)
                    contraction_ranges[idx][1] = round(new_e, 3)
                redraw()

        def on_mouse_release(event):
            dragging['range_idx'] = None
            dragging['endpoint'] = None
            dragging['offset'] = 0.0
            dragging['mode'] = None

        def remove_range_menu(event):
            idx, endpoint, mode = find_near_endpoint_or_body(event)
            if idx is not None:
                menu = tk.Menu(canvas, tearoff=0)

                def do_remove():
                    contraction_ranges.pop(idx)
                    redraw()

                s, e = contraction_ranges[idx]
                menu.add_command(label=f"删除区间 {s:.3f}s ~ {e:.3f}s", command=do_remove)
                menu.tk_popup(event.x_root, event.y_root)

        canvas.bind("<Button-1>", add_or_drag_marker)
        canvas.bind("<B1-Motion>", on_mouse_move)
        canvas.bind("<ButtonRelease-1>", on_mouse_release)
        canvas.bind("<Button-3>", remove_range_menu)

        # 修改redraw，支持区间高亮和端点绘制，并记录y轴最大/最小刻度像素位置
        def redraw():
            canvas.delete("all")
            w = canvas.winfo_width()
            h = canvas.winfo_height()
            start = slider_var.get()
            slider_label.config(text=f"起始时间: {start:.2f}s")
            end = start + window_duration
            idx = np.where((times >= t_min + start) & (times <= t_min + end))[0]
            if len(idx) < 2:
                return
            t_window = times[idx] - t_min
            # 画y轴主刻度线：最大-0之间3条，最小-0之间3条
            y_ticks = []
            for i in range(4):
                y_ticks.append(y_top - i * (y_top / 3))
            for i in range(1, 4):
                y_ticks.append(y_bottom + i * (-y_bottom / 3))
            # 保证y_bottom一定在y_ticks中
            if not any(abs(v - y_bottom) < 1e-6 for v in y_ticks):
                y_ticks.append(y_bottom)
            y_ticks = sorted(set([round(v, 6) for v in y_ticks]), reverse=True)
            _yaxis_max_marker[0] = None
            _yaxis_min_marker[0] = None
            for value in y_ticks:
                y_pos = (y_top - value) * y_scale(h) + 20
                color = '#EEE'
                width = 1
                canvas.create_line(left_margin, y_pos, w, y_pos, fill=color, width=width)
                canvas.create_text(left_margin - 5, y_pos, text=f"{value:.3f}", anchor="e", fill="#666")
                if abs(value - y_top) < 1e-6:
                    _yaxis_max_marker[0] = (y_pos, value)
                if abs(value - y_bottom) < 1e-6:
                    _yaxis_min_marker[0] = (y_pos, value)
            # 画x轴网格线
            x_axis_y = h - bottom_margin
            duration = window_duration
            x_grid_step = 1.5
            x = start - (start % x_grid_step)
            while x <= end + 1e-6:
                x_pos = left_margin + (w - left_margin - 10) * ((x - start) / duration)
                color = '#EEE'
                width = 1
                canvas.create_line(x_pos, 20, x_pos, x_axis_y, fill=color, width=width)
                canvas.create_text(x_pos, x_axis_y + bottom_margin / 2, text=f"{x:.1f}s", anchor="n", fill="#666")
                x += x_grid_step
            canvas.create_line(left_margin, x_axis_y, w, x_axis_y, fill='#EEE')
            # 波形
            data_to_plot = values if show_filtered.get() else raw_values
            for ch in range(num_channels):
                if not channel_vars[ch].get():
                    continue
                yvals = data_to_plot[idx, ch]
                xvals = left_margin + (w - left_margin - 10) * (t_window - start) / window_duration
                yvals_plot = (y_top - yvals) * y_scale(h) + 20
                points = list(zip(xvals, yvals_plot))
                if len(points) > 1:
                    canvas.create_line(points, fill=preview_colors[ch], width=1)
            # 区间标记
            for s, e in contraction_ranges:
                if e < t_min + start or s > t_min + end:
                    continue
                sx = left_margin + (w - left_margin - 10) * ((max(s, t_min + start) - t_min - start) / duration)
                ex = left_margin + (w - left_margin - 10) * ((min(e, t_min + end) - t_min - start) / duration)
                canvas.create_rectangle(sx, 20, ex, x_axis_y, fill='#ffcccc', outline='', stipple='gray25')
                canvas.create_line(sx, 20, sx, x_axis_y, fill='red', width=2)
                canvas.create_line(ex, 20, ex, x_axis_y, fill='blue', width=2)
                canvas.create_oval(sx - marker_radius, x_axis_y - marker_radius, sx + marker_radius,
                                   x_axis_y + marker_radius, fill='red', outline='black')
                canvas.create_oval(ex - marker_radius, x_axis_y - marker_radius, ex + marker_radius,
                                   x_axis_y + marker_radius, fill='blue', outline='black')
                canvas.create_text(sx, x_axis_y + 18, text="起始", fill='red', font=("Arial", 10, "bold"))
                canvas.create_text(ex, x_axis_y + 18, text="结束", fill='blue', font=("Arial", 10, "bold"))
            # 更新全选按钮文本
            if all(var.get() for var in channel_vars):
                toggle_btn.config(text="取消全选")
                select_all_state[0] = True
            elif not any(var.get() for var in channel_vars):
                toggle_btn.config(text="全选")
                select_all_state[0] = False
            else:
                toggle_btn.config(text="全选")
                select_all_state[0] = False

        # 通道复选框和全选按钮区域
        frame = ttk.Frame(win)
        frame.pack(fill="x", side="bottom")
        channel_vars = []
        select_all_state = [True]

        def toggle_all():
            new_state = not all(var.get() for var in channel_vars)
            for var in channel_vars:
                var.set(new_state)
            select_all_state[0] = new_state
            toggle_btn.config(text="全选" if not new_state else "取消全选")
            redraw()

        toggle_btn = ttk.Button(frame, text="取消全选", command=toggle_all, width=8)
        toggle_btn.pack(side="left", padx=5)
        for i in range(num_channels):
            var = tk.BooleanVar(value=True)
            channel_vars.append(var)
            chk = tk.Checkbutton(frame, text=f"通道 {i + 1}", variable=var, fg=preview_colors[i],
                                 activeforeground=preview_colors[i], command=lambda idx=i: redraw())
            chk.pack(side="left", padx=5)

        # 新增：按钮区域，横向排列，放在通道复选框下方
        button_frame = ttk.Frame(win)
        button_frame.pack(fill="x", side="bottom", anchor="w", pady=(0, 10))
        save_btn = ttk.Button(button_frame, text="保存标记区间数据", command=save_marked_segments)
        save_btn.pack(side="left", padx=10, pady=5)
        filter_btn = ttk.Button(button_frame, text="显示原始波形", command=toggle_filter)
        filter_btn.pack(side="left", padx=10, pady=5)

        canvas.bind("<Configure>", lambda e: redraw())
        redraw()

        # y轴双击事件
        def on_yaxis_double_click(event):
            x = event.x
            y = event.y
            margin = 30
            max_marker = _yaxis_max_marker[0]
            min_marker = _yaxis_min_marker[0]
            # 检查是否靠近最大/最小刻度
            if max_marker and abs(y - max_marker[0]) < margin and x < left_margin:
                show_yaxis_entry(y=max_marker[0], value=y_top, is_max=True)
            elif min_marker and abs(y - min_marker[0]) < margin and x < left_margin:
                show_yaxis_entry(y=min_marker[0], value=y_bottom, is_max=False)

        def show_yaxis_entry(y, value, is_max):
            # 若已有输入框，先销毁
            if _yaxis_entry[0]:
                _yaxis_entry[0].destroy()
                _yaxis_entry[0] = None
            entry_x = left_margin - 45
            entry_y = int(y) - 12
            entry = tk.Entry(canvas, width=8, justify='right')
            entry.insert(0, str(value))
            entry.place(x=entry_x, y=entry_y)
            entry.focus_set()
            entry.select_range(0, tk.END)
            _yaxis_entry[0] = entry

            def confirm(event=None):
                nonlocal y_top, y_bottom
                try:
                    new_val = float(entry.get())
                    if is_max:
                        if new_val > y_bottom:
                            y_top = new_val
                    else:
                        if new_val < y_top:
                            y_bottom = new_val
                    # 保存到config.txt（预览模式）
                    self.save_yaxis_range(y_top, y_bottom, mode="preview")
                    entry.destroy()
                    _yaxis_entry[0] = None
                    redraw()
                except Exception:
                    entry.bell()

            def cancel(event=None):
                entry.destroy()
                _yaxis_entry[0] = None

            entry.bind('<Return>', confirm)
            entry.bind('<KP_Enter>', confirm)
            entry.bind('<Escape>', cancel)
            entry.bind('<FocusOut>', cancel)

        canvas.bind("<Double-Button-1>", on_yaxis_double_click)

    def load_num_channels(self):
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        if line.startswith('num_channels='):
                            val = int(line.strip().split('=')[1])
                            if 1 <= val <= 16:
                                return val
        except Exception:
            pass
        return 1  # 默认值

    def save_num_channels(self):
        try:
            # 读取现有配置
            lines = []
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()

            # 更新或添加通道数配置
            found = False
            for i, line in enumerate(lines):
                if line.startswith('num_channels='):
                    lines[i] = f'num_channels={self.num_channels}\n'
                    found = True
                    break

            # 如果参数不存在，则添加
            if not found:
                lines.append(f'num_channels={self.num_channels}\n')

            # 写入配置文件
            with open(self.config_path, 'w', encoding='utf-8') as f:
                f.writelines(lines)
        except Exception:
            pass

    def load_yaxis_range(self, mode="main"):
        # mode: "main" or "preview"
        if mode == "main":
            y_top_key = 'y_top_main'
            y_bottom_key = 'y_bottom_main'
            default_top = 1.5
            default_bottom = -1.5
        else:
            y_top_key = 'y_top_preview'
            y_bottom_key = 'y_bottom_preview'
            default_top = 1.5
            default_bottom = -1.5
        y_top = default_top
        y_bottom = default_bottom
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        if line.startswith(f'{y_top_key}='):
                            y_top = float(line.strip().split('=')[1])
                        elif line.startswith(f'{y_bottom_key}='):
                            y_bottom = float(line.strip().split('=')[1])
        except Exception:
            pass
        return y_top, y_bottom

    def save_yaxis_range(self, y_top, y_bottom, mode="main"):
        # mode: "main" or "preview"
        if mode == "main":
            y_top_key = 'y_top_main'
            y_bottom_key = 'y_bottom_main'
        else:
            y_top_key = 'y_top_preview'
            y_bottom_key = 'y_bottom_preview'
        lines = []
        if os.path.exists(self.config_path):
            with open(self.config_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
        found_top = found_bottom = False
        for i, line in enumerate(lines):
            if line.startswith(f'{y_top_key}='):
                lines[i] = f'{y_top_key}={y_top}\n'
                found_top = True
            elif line.startswith(f'{y_bottom_key}='):
                lines[i] = f'{y_bottom_key}={y_bottom}\n'
                found_bottom = True
        if not found_top:
            lines.append(f'{y_top_key}={y_top}\n')
        if not found_bottom:
            lines.append(f'{y_bottom_key}={y_bottom}\n')
        with open(self.config_path, 'w', encoding='utf-8') as f:
            f.writelines(lines)

    def on_yaxis_double_click(self, event):
        """双击y轴最大/最小刻度，在刻度处直接输入新值，回车确认"""
        x = event.x
        y = event.y
        margin = 30  # 允许的像素误差
        max_marker = getattr(self, '_yaxis_max_marker', None)
        min_marker = getattr(self, '_yaxis_min_marker', None)
        # 检查是否靠近最大/最小刻度
        if max_marker and abs(y - max_marker[0]) < margin and x < self.left_margin:
            # 编辑最大刻度
            self.show_yaxis_entry(y=max_marker[0], value=self.y_top_fixed, is_max=True)
        elif min_marker and abs(y - min_marker[0]) < margin and x < self.left_margin:
            # 编辑最小刻度
            self.show_yaxis_entry(y=min_marker[0], value=self.y_bottom_fixed, is_max=False)

    def show_yaxis_entry(self, y, value, is_max):
        """在画布y轴刻度处显示Entry输入框，回车确认，Esc或失焦取消"""
        # 若已有输入框，先销毁
        if hasattr(self, '_yaxis_entry') and self._yaxis_entry:
            self._yaxis_entry.destroy()
            self._yaxis_entry = None
        # 计算输入框位置（画布左侧margin-5，纵向y）
        entry_x = self.left_margin - 45
        entry_y = int(y) - 12  # 适当上移
        # 创建Entry
        entry = tk.Entry(self.canvas, width=8, justify='right')
        entry.insert(0, str(value))
        entry.place(x=entry_x, y=entry_y)
        entry.focus_set()
        entry.select_range(0, tk.END)  # 自动全选内容
        self._yaxis_entry = entry

        # 事件处理
        def confirm(event=None):
            try:
                new_val = float(entry.get())
                if is_max:
                    if new_val > self.y_bottom_fixed:
                        self.y_top_fixed = new_val
                        self.save_yaxis_range(self.y_top_fixed, self.y_bottom_fixed, mode="main")
                else:
                    if new_val < self.y_top_fixed:
                        self.y_bottom_fixed = new_val
                        self.save_yaxis_range(self.y_top_fixed, self.y_bottom_fixed, mode="main")
                self.effective_height = self.plot_height - self.bottom_margin - self.top_margin
                self.y_scale = self.effective_height / (self.y_top_fixed - self.y_bottom_fixed)
                entry.destroy()
                self._yaxis_entry = None
                self.update_plot_once()
            except Exception:
                entry.bell()

        def cancel(event=None):
            entry.destroy()
            self._yaxis_entry = None

        entry.bind('<Return>', confirm)
        entry.bind('<KP_Enter>', confirm)
        entry.bind('<Escape>', cancel)
        entry.bind('<FocusOut>', cancel)

    def init_speech_engine(self):
        """初始化语音播报引擎"""
        try:
            with self.speech_lock:
                self.tts_engine = win32com.client.Dispatch("SAPI.SpVoice")
                self.tts_engine.Rate = 0  # 语速 (范围: -10 到 10, 0为正常)
                self.tts_engine.Volume = 100  # 音量 (范围: 0 到 100)
                # 测试语音引擎是否正常工作
                self.tts_engine.Speak("", 0)  # 空字符串测试
                print("语音引擎初始化成功")
        except Exception as e:
            print(f"语音引擎初始化失败: {e}")
            self.tts_engine = None

    def speak(self, text):
        """语音播报方法"""
        if not text or not self.tts_engine:
            return

        # 新增：记录语音播报结束时间，并计算中点
        pending = getattr(self, '_pending_voice_time', None)
        is_voice_mark = text in ("请握紧", "请放松")
        start_time = None
        if is_voice_mark and pending and pending.get('text') == text:
            start_time = pending.get('start_time')
        try:
            with self.speech_lock:
                # 检查语音引擎状态
                if self.tts_engine:
                    self.tts_engine.Speak(text, 0)  # 0表示同步播报
            # 新增：播报结束后记录结束时间，计算中点
            if is_voice_mark and start_time is not None:
                end_time = time.time() - self.start_time
                mid_time = (start_time + end_time) / 2
                with self.voice_timestamp_lock:
                    self.voice_timestamps.append((mid_time, text))
                self._pending_voice_time = None
        except Exception as e:
            print(f"语音播报错误: {e}")
            # 尝试重新初始化语音引擎
            try:
                self.init_speech_engine()
            except:
                print("语音引擎重新初始化失败")

    def speak_async(self, text):
        # 新增：记录语音播报开始时间点（只记录"请握紧"和"请放松"）
        if (self.voice_broadcast_enabled and self.data_collection_active and
                (text == "请握紧" or text == "请放松")):
            # 使用与数据文件相同的时间基准（程序启动时间）
            current_time = time.time() - self.start_time
            # 暂存到实例变量，待speak播报完毕后计算中点
            self._pending_voice_time = {'start_time': current_time, 'text': text}
        self.speech_queue.put(text)

    def speak_worker(self):
        while True:
            text = self.speech_queue.get()
            try:
                self.speak(text)
            except Exception as e:
                print(f"语音播报错误: {e}")

    def start_cycle_prompt(self):
        """启动放松-握紧循环语音提示（第一次放松已由按钮播报，放松时间后进入第一次握紧）"""
        # 只有在语音播报开启时才启动循环
        if not self.voice_broadcast_enabled:
            self.gui_queue.put(("msg", "语音播报已关闭，不会启动自动循环提示", 'info'))
            return

        self.cycle_count = 1  # 第一次放松已播报
        self.cycle_stage = 1  # 1=握紧, 0=放松
        self.cycle_running = True
        # 放松时间后进入第一次"请握紧"
        self.cycle_timer_id = self.root.after(int(self.relaxation_time * 1000), self.cycle_next_stage)

    def cycle_next_stage(self):
        # 优先处理实验结束
        if self.cycle_stage == 3:
            if self.voice_broadcast_enabled:
                self.speak_async("实验结束")
                self.gui_queue.put(("msg", f"实验已完成{self.cycle_total}次循环，自动暂停采集", 'info'))
                self.cycle_running = False
                self.cycle_stage = 4  # 防止再次进入
                self.root.after(1000, self.toggle_data_collection)
            else:
                # 语音播报关闭时，不自动结束采集，只停止循环
                self.gui_queue.put(("msg", f"实验已完成{self.cycle_total}次循环，请手动点击'暂停采集'按钮结束", 'info'))
                self.cycle_running = False
                self.cycle_stage = 4  # 防止再次进入
            return
        # 只要不是最后一次实验结束阶段，且未运行，则阻断
        if not self.cycle_running:
            return
        # 1~9次正常循环
        if self.cycle_count < self.cycle_total:
            if self.cycle_stage == 1:
                if self.voice_broadcast_enabled:
                    self.speak_async("请握紧")
                    self.gui_queue.put(("msg", f"第{self.cycle_count}次：请握紧", 'info'))
                self.cycle_stage = 0
                # 使用收缩时间
                self.cycle_timer_id = self.root.after(int(self.contraction_time * 1000), self.cycle_next_stage)
            else:
                if self.voice_broadcast_enabled:
                    self.speak_async("请放松")
                    self.gui_queue.put(("msg", f"第{self.cycle_count + 1}次：请放松", 'info'))
                self.cycle_stage = 1
                self.cycle_count += 1
                # 使用放松时间
                self.cycle_timer_id = self.root.after(int(self.relaxation_time * 1000), self.cycle_next_stage)
        # 第10次特殊处理
        elif self.cycle_count == self.cycle_total:
            if self.cycle_stage == 1:
                if self.voice_broadcast_enabled:
                    self.speak_async("请握紧")
                    self.gui_queue.put(("msg", f"第{self.cycle_count}次：请握紧", 'info'))
                self.cycle_stage = 0
                # 使用收缩时间
                self.cycle_timer_id = self.root.after(int(self.contraction_time * 1000), self.cycle_next_stage)
            elif self.cycle_stage == 0:
                if self.voice_broadcast_enabled:
                    self.speak_async("请放松")
                    self.gui_queue.put(("msg", f"第{self.cycle_total + 1}次：请放松", 'info'))
                self.cycle_stage = 3  # 结束标志
                # 使用放松时间
                self.cycle_timer_id = self.root.after(int(self.relaxation_time * 1000), self.cycle_next_stage)

    def stop_cycle_prompt(self):
        """停止循环语音提示"""
        self.cycle_running = False
        if self.cycle_timer_id:
            self.root.after_cancel(self.cycle_timer_id)
            self.cycle_timer_id = None

    def load_experiment_params(self):
        """加载实验参数：收缩次数、收缩时间、放松时间、语音播报开关"""
        cycle_total = 10  # 默认收缩次数
        contraction_time = 10.0  # 默认收缩时间（秒）
        relaxation_time = 10.0  # 默认放松时间（秒）
        voice_broadcast_enabled = True  # 默认语音播报开启
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        if line.startswith('cycle_total='):
                            cycle_total = int(line.strip().split('=')[1])
                        elif line.startswith('contraction_time='):
                            contraction_time = float(line.strip().split('=')[1])
                        elif line.startswith('relaxation_time='):
                            relaxation_time = float(line.strip().split('=')[1])
                        elif line.startswith('voice_broadcast_enabled='):
                            v = line.strip().split('=')[1].strip()
                            voice_broadcast_enabled = (v == '1' or v.lower() == 'true')
        except Exception:
            pass
        return cycle_total, contraction_time, relaxation_time, voice_broadcast_enabled

    def save_experiment_params(self):
        """保存实验参数到配置文件"""
        try:
            # 获取当前输入框的值
            cycle_total = int(self.cycle_total_var.get())
            contraction_time = float(self.contraction_time_var.get())
            relaxation_time = float(self.relaxation_time_var.get())
            voice_broadcast_enabled = self.voice_broadcast_var.get()
            # 验证参数范围
            if cycle_total < 1 or cycle_total > 50:
                messagebox.showerror("错误", "收缩次数必须在1到50之间")
                return
            if contraction_time < 1.0 or contraction_time > 60.0:
                messagebox.showerror("错误", "收缩时间必须在1到60秒之间")
                return
            if relaxation_time < 1.0 or relaxation_time > 60.0:
                messagebox.showerror("错误", "放松时间必须在1到60秒之间")
                return
            # 更新实例变量
            self.cycle_total = cycle_total
            self.contraction_time = contraction_time
            self.relaxation_time = relaxation_time
            self.voice_broadcast_enabled = voice_broadcast_enabled
            # 读取现有配置
            lines = []
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
            # 更新或添加实验参数
            found_cycle = found_contraction = found_relaxation = found_voice = False
            for i, line in enumerate(lines):
                if line.startswith('cycle_total='):
                    lines[i] = f'cycle_total={cycle_total}\n'
                    found_cycle = True
                elif line.startswith('contraction_time='):
                    lines[i] = f'contraction_time={contraction_time}\n'
                    found_contraction = True
                elif line.startswith('relaxation_time='):
                    lines[i] = f'relaxation_time={relaxation_time}\n'
                    found_relaxation = True
                elif line.startswith('voice_broadcast_enabled='):
                    lines[i] = f'voice_broadcast_enabled={1 if voice_broadcast_enabled else 0}\n'
                    found_voice = True
            # 如果参数不存在，则添加
            if not found_cycle:
                lines.append(f'cycle_total={cycle_total}\n')
            if not found_contraction:
                lines.append(f'contraction_time={contraction_time}\n')
            if not found_relaxation:
                lines.append(f'relaxation_time={relaxation_time}\n')
            if not found_voice:
                lines.append(f'voice_broadcast_enabled={1 if voice_broadcast_enabled else 0}\n')
            # 写入配置文件
            with open(self.config_path, 'w', encoding='utf-8') as f:
                f.writelines(lines)
            self.gui_queue.put(("msg",
                                f"实验参数已保存：收缩{cycle_total}次，收缩时间{contraction_time}秒，放松时间{relaxation_time}秒，语音播报{'开启' if voice_broadcast_enabled else '关闭'}",
                                'info'))
        except ValueError as e:
            messagebox.showerror("错误", f"参数格式错误：{str(e)}")
        except Exception as e:
            messagebox.showerror("错误", f"保存参数失败：{str(e)}")

    def on_voice_broadcast_toggle(self, *args):
        self.voice_broadcast_enabled = self.voice_broadcast_var.get()

    def on_save_connection_params(self):
        """处理保存连接参数按钮点击事件"""
        self.save_connection_params(silent=False)

    def load_connection_params(self):
        """加载连接参数：IP地址和端口号"""
        ip_address = "192.168.31.22"  # 默认IP地址
        port_number = 8080  # 默认端口号
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        if line.startswith('ip_address='):
                            ip_address = line.strip().split('=')[1]
                        elif line.startswith('port_number='):
                            port_number = int(line.strip().split('=')[1])
        except Exception:
            pass
        return ip_address, port_number

    def save_connection_params(self, silent=False):
        """保存连接参数到配置文件"""
        try:
            # 获取当前输入框的值
            ip_address = self.ip_entry.get().strip()
            port_number = int(self.port_entry.get().strip())

            # 验证IP地址格式
            ip_pattern = r'^(\d{1,3}\.){3}\d{1,3}$'
            if not re.match(ip_pattern, ip_address):
                if not silent:
                    messagebox.showerror("错误", "IP地址格式不正确")
                return False

            # 验证IP地址每个数字段的范围
            try:
                ip_parts = ip_address.split('.')
                for part in ip_parts:
                    if not (0 <= int(part) <= 255):
                        if not silent:
                            messagebox.showerror("错误", "IP地址数字段必须在0-255范围内")
                        return False
            except ValueError:
                if not silent:
                    messagebox.showerror("错误", "IP地址格式不正确")
                return False

            # 验证端口号范围
            if not (1 <= port_number <= 65535):
                if not silent:
                    messagebox.showerror("错误", "端口号必须在1到65535之间")
                return False

            # 更新实例变量
            self.ip_address = ip_address
            self.port_number = port_number

            # 读取现有配置
            lines = []
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()

            # 更新或添加连接参数
            found_ip = found_port = False
            for i, line in enumerate(lines):
                if line.startswith('ip_address='):
                    lines[i] = f'ip_address={ip_address}\n'
                    found_ip = True
                elif line.startswith('port_number='):
                    lines[i] = f'port_number={port_number}\n'
                    found_port = True

            # 如果参数不存在，则添加
            if not found_ip:
                lines.append(f'ip_address={ip_address}\n')
            if not found_port:
                lines.append(f'port_number={port_number}\n')

            # 写入配置文件
            with open(self.config_path, 'w', encoding='utf-8') as f:
                f.writelines(lines)

            if not silent:
                self.gui_queue.put(("msg", f"连接参数已保存：IP地址 {ip_address}，端口号 {port_number}", 'info'))
            return True
        except ValueError as e:
            if not silent:
                messagebox.showerror("错误", f"端口号格式错误：{str(e)}")
            return False
        except Exception as e:
            if not silent:
                messagebox.showerror("错误", f"保存连接参数失败：{str(e)}")
            return False


if __name__ == "__main__":
    root = tk.Tk()
    app = NetworkDebugger(root)
    root.mainloop()