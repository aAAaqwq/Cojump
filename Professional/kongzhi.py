import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import socket
import threading
import time
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np

# 设置matplotlib支持中文显示
# 只保留系统中肯定存在的中文字体
plt.rcParams["font.family"] = ["SimHei", "Microsoft YaHei", "Arial"]
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题


class WiFiHardwareController:
    def __init__(self, root):
        self.root = root
        self.root.title("WiFi硬件控制器")
        self.root.geometry("1000x700")
        self.root.configure(bg="#f0f0f0")

        # 设置中文字体支持
        self.style = ttk.Style()
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TLabel", font=("SimHei", 10), background="#f0f0f0")
        self.style.configure("TFrame", background="#f0f0f0")
        self.style.configure("TLabelframe.Label", font=("SimHei", 11, "bold"))

        # 网络连接状态
        self.connected = False
        self.client_socket = None
        self.receive_thread = None
        self.running = True

        # EMG数据记录
        self.emg_data = []
        self.emg_time = []
        self.max_data_points = 100  # 最大数据点数量

        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # 创建左侧控制面板和右侧数据显示面板
        self.control_panel = ttk.Frame(self.main_frame)
        self.data_panel = ttk.Frame(self.main_frame)

        self.control_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.data_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)

        # 创建连接区域
        self.create_connection_frame()

        # 创建热身模式区域
        self.create_control_frame()

        # 创建肌力训练模式区域
        self.create_mode_frame()

        # 创建模式切换区域
        self.create_change_frame()

        # 创建EMG校准训练区域
        self.create_emg_calibration_frame()

        # 创建日志区域
        self.create_log_frame()

        # 创建数据可视化区域
        self.create_visualization_frame()

        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)


    def open_second_window(self):
        # 在需要时才导入第二个界面，避免循环导入
        from caiji import NetworkDebugger
        # 隐藏当前界面
        self.main_frame.destroy()
        # 创建并显示第二个界面
        NetworkDebugger(self.root)

    def create_connection_frame(self):
        frame = ttk.LabelFrame(self.control_panel, text="WiFi连接设置", padding="10")
        frame.pack(fill=tk.X, pady=5)

        ttk.Label(frame, text="IP地址:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.ip_entry = ttk.Entry(frame, width=20)
        self.ip_entry.grid(row=0, column=1, padx=5, pady=5)
        self.ip_entry.insert(0, "192.168.43.1")  # 默认IP地址

        ttk.Label(frame, text="端口:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.port_entry = ttk.Entry(frame, width=10)
        self.port_entry.grid(row=0, column=3, padx=5, pady=5)
        self.port_entry.insert(0, "80")  # 默认端口

        self.connect_btn = ttk.Button(frame, text="连接", command=self.connect)
        self.connect_btn.grid(row=0, column=4, padx=5, pady=5)

        self.disconnect_btn = ttk.Button(frame, text="断开连接", command=self.disconnect, state=tk.DISABLED)
        self.disconnect_btn.grid(row=0, column=5, padx=5, pady=5)

        self.connection_status = ttk.Label(frame, text="未连接", foreground="red")
        self.connection_status.grid(row=0, column=6, padx=5, pady=5)

    def create_control_frame(self):
        frame = ttk.LabelFrame(self.control_panel, text="热身模式", padding="10")
        frame.pack(fill=tk.X, pady=5)

        # 速度控制
        speed_frame = ttk.Frame(frame)
        speed_frame.pack(fill=tk.X, pady=5)

        ttk.Label(speed_frame, text="设备速度（1500-3000）:").pack(side=tk.LEFT, padx=5)
        self.speed_entry = ttk.Entry(speed_frame, width=10)
        self.speed_entry.pack(side=tk.LEFT, padx=5)
        self.speed_entry.insert(0, "1500")

        self.set_speed_btn = ttk.Button(speed_frame, text="开始热身",
                                        command=lambda: self.send_command(f"S{self.speed_entry.get()}"))
        self.set_speed_btn.pack(side=tk.LEFT, padx=5)

        self.pause_btn = ttk.Button(speed_frame, text="暂停", command=lambda: self.send_command("pause"), width=10)
        self.pause_btn.pack(side=tk.LEFT, padx=5)

    def create_mode_frame(self):
        frame = ttk.LabelFrame(self.control_panel, text="肌力训练模式", padding="10")
        frame.pack(fill=tk.X, pady=5)

        # 模式选择按钮
        mode_frame = ttk.Frame(frame)
        mode_frame.pack(fill=tk.X, pady=5)

        self.mode1_btn = ttk.Button(mode_frame, text="模式1", command=lambda: self.send_command("mode1"), width=10)
        self.mode1_btn.pack(side=tk.LEFT, padx=5)

        self.mode2_btn = ttk.Button(mode_frame, text="模式2", command=lambda: self.send_command("mode2"), width=10)
        self.mode2_btn.pack(side=tk.LEFT, padx=5)

        self.mode3_btn = ttk.Button(mode_frame, text="模式3", command=lambda: self.send_command("mode3"), width=10)
        self.mode3_btn.pack(side=tk.LEFT, padx=5)

        # 训练控制区域（包含时长设置和倒计时）
        control_frame = ttk.LabelFrame(frame, text="训练控制", padding="5")
        control_frame.pack(fill=tk.X, pady=5)

        # 训练时长设置
        ttk.Label(control_frame, text="训练时长(秒):").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.training_duration = ttk.Entry(control_frame, width=10)
        self.training_duration.grid(row=0, column=1, padx=5, pady=5)
        self.training_duration.insert(0, "60")  # 默认60秒

        # 倒计时显示
        ttk.Label(control_frame, text="剩余时间:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.countdown_label = ttk.Label(control_frame, text="--:--")
        self.countdown_label.grid(row=0, column=3, padx=5, pady=5)

        # 控制按钮
        self.up_btn = ttk.Button(control_frame, text="开始训练", command=self.start_training, width=10)
        self.up_btn.grid(row=0, column=4, padx=5, pady=5)

        self.resume_btn = ttk.Button(control_frame, text="继续训练", command=self.resume_training, width=10)
        self.resume_btn.grid(row=0, column=6, padx=5, pady=5)

        self.pause_btn = ttk.Button(control_frame, text="暂停训练", command=self.pause_training, width=10)
        self.pause_btn.grid(row=0, column=5, padx=5, pady=5)

        # 初始化倒计时相关变量
        self.countdown_running = False
        self.remaining_time = 0
        self.countdown_after_id = None

        # 模式速度设置
        speed_mode_frame = ttk.LabelFrame(frame, text="模式速度", padding="5")
        speed_mode_frame.pack(fill=tk.X, pady=5)

        ttk.Label(speed_mode_frame, text="目前上拉速度:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.speedup1_entry = ttk.Entry(speed_mode_frame, width=10)
        self.speedup1_entry.grid(row=0, column=1, padx=5, pady=5)
        self.speedup1_entry.insert(0, "2000")

        ttk.Label(speed_mode_frame, text="目前下拉速度:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.speeddown1_entry = ttk.Entry(speed_mode_frame, width=10)
        self.speeddown1_entry.grid(row=0, column=3, padx=5, pady=5)
        self.speeddown1_entry.insert(0, "1500")

        ttk.Label(speed_mode_frame, text="获取速度参数:").grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        self.get_speed_btn = ttk.Button(speed_mode_frame, text="获取",
                                        command=self.get_speed_parameters)
        self.get_speed_btn.grid(row=0, column=5, padx=5, pady=5)

        # 时间参数控制
        time_frame = ttk.LabelFrame(frame, text="时间参数", padding="5")
        time_frame.pack(fill=tk.X, pady=5)

        ttk.Label(time_frame, text="时间1:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.time2_entry = ttk.Entry(time_frame, width=10)
        self.time2_entry.grid(row=0, column=1, padx=5, pady=5)
        self.time2_entry.insert(0, "1000000")

        ttk.Label(time_frame, text="时间2:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.time3_entry = ttk.Entry(time_frame, width=10)
        self.time3_entry.grid(row=0, column=3, padx=5, pady=5)
        self.time3_entry.insert(0, "2000000")

        ttk.Label(time_frame, text="时间3:").grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        self.time4_entry = ttk.Entry(time_frame, width=10)
        self.time4_entry.grid(row=0, column=5, padx=5, pady=5)
        self.time4_entry.insert(0, "3000000")

        ttk.Label(time_frame, text="获取时间参数:").grid(row=0, column=6, padx=5, pady=5, sticky=tk.W)
        self.get_time_btn = ttk.Button(time_frame, text="获取",
                                       command=self.get_time_parameters)
        self.get_time_btn.grid(row=0, column=8, padx=5, pady=5)

    def start_training(self):
        """开始训练并启动倒计时"""
        # 发送开始命令
        self.send_command("start")

        # 如果不在运行中，获取训练时长并开始倒计时
        if not self.countdown_running:
            try:
                # 从输入框获取时长（秒）
                duration = int(self.training_duration.get())
                if duration <= 0:
                    raise ValueError
                self.remaining_time = duration
                self.countdown_running = True
                self.update_countdown()
            except ValueError:
                # 输入无效时显示错误消息并暂停训练
                tk.messagebox.showerror("输入错误", "请输入有效的训练时长（正整数）")
                self.send_command("pause")

    # 继续训练方法
    def resume_training(self):
        """继续训练并恢复倒计时"""
        # 发送继续命令
        self.send_command("start")

        # 如果处于暂停状态且还有剩余时间，则恢复倒计时
        if not self.countdown_running and self.remaining_time > 0:
            self.countdown_running = True
            self.update_countdown()

    def pause_training(self):
        """暂停训练并停止倒计时"""
        # 发送暂停命令
        self.send_command("pause")

        # 停止倒计时
        self.countdown_running = False
        if self.countdown_after_id:
            self.control_panel.after_cancel(self.countdown_after_id)
            self.countdown_after_id = None

    def update_countdown(self):
        """更新倒计时显示"""
        if self.countdown_running and self.remaining_time > 0:
            # 格式化剩余时间为分:秒
            minutes = self.remaining_time // 60
            seconds = self.remaining_time % 60
            self.countdown_label.config(text=f"{minutes:02d}:{seconds:02d}")

            # 减少一秒
            self.remaining_time -= 1

            # 一秒后再次更新
            self.countdown_after_id = self.control_panel.after(1000, self.update_countdown)
        elif self.remaining_time <= 0:
            # 倒计时结束，自动暂停
            self.countdown_label.config(text="00:00")
            self.pause_training()

    def create_change_frame(self):
        frame = ttk.LabelFrame(self.control_panel, text="模式切换", padding="10")
        frame.pack(fill=tk.X, pady=5)

        # 模式切换
        change_switch_frame = ttk.Frame(frame)
        change_switch_frame.pack(fill=tk.X, pady=5)

        # 修改按钮命令，在发送命令的同时更新显示
        self.emg_change_btn = ttk.Button(
            change_switch_frame,
            text="切换到EMG模式",
            # 先发送命令，再更新标签文本
            command=lambda: [self.send_command("emgmode"), self.current_change.config(text="当前模式: EMG模式")],
            width=15
        )
        self.emg_change_btn.pack(side=tk.LEFT, padx=5)

        self.wifi_change_btn = ttk.Button(
            change_switch_frame,
            text="切换到WiFi模式",
            # 同样更新WiFi模式的显示
            command=lambda: [self.send_command("wifimode"), self.current_change.config(text="当前模式: WiFi模式")],
            width=15
        )
        self.wifi_change_btn.pack(side=tk.LEFT, padx=5)

        # 当前模式显示
        self.current_change = ttk.Label(change_switch_frame, text="当前模式: WiFi模式", foreground="blue")
        self.current_change.pack(side=tk.LEFT, padx=20)


    def create_emg_calibration_frame(self):
        frame = ttk.LabelFrame(self.control_panel, text="EMG校准", padding="10")
        frame.pack(fill=tk.X, pady=5)

        calib_frame = ttk.Frame(frame)
        calib_frame.pack(fill=tk.X, pady=5)

        self.start_calib_btn = ttk.Button(calib_frame, text="开始EMG校准",
                                          command=lambda: self.send_command("emgcalibrate"), width=15)
        self.start_calib_btn.pack(side=tk.LEFT, padx=5)

        self.apply_calib_btn = ttk.Button(calib_frame, text="应用推荐阈值",
                                          command=lambda: self.send_command("applyrecommended"), width=15)
        self.apply_calib_btn.pack(side=tk.LEFT, padx=5)

        self.reset_calib_btn = ttk.Button(calib_frame, text="重置校准数据",
                                          command=self.reset_calibration, width=15)
        self.reset_calib_btn.pack(side=tk.LEFT, padx=5)

        # 校准状态
        self.calib_status = ttk.Label(frame, text="校准状态: 未开始", foreground="green")
        self.calib_status.pack(side=tk.LEFT, padx=20)

        # 校准结果
        result_frame = ttk.Frame(frame)
        result_frame.pack(fill=tk.X, pady=5)

        ttk.Label(result_frame, text="放松平均值:").pack(side=tk.LEFT, padx=5)
        self.relax_avg = ttk.Label(result_frame, text="--")
        self.relax_avg.pack(side=tk.LEFT, padx=5)

        ttk.Label(result_frame, text="用力平均值:").pack(side=tk.LEFT, padx=5)
        self.contract_avg = ttk.Label(result_frame, text="--")
        self.contract_avg.pack(side=tk.LEFT, padx=5)

        ttk.Label(result_frame, text="推荐阈值:").pack(side=tk.LEFT, padx=5)
        self.recommended_threshold = ttk.Label(result_frame, text="--")
        self.recommended_threshold.pack(side=tk.LEFT, padx=5)

        # EMG状态
        emg_status_frame = ttk.Frame(frame)
        emg_status_frame.pack(fill=tk.X, pady=5)

        ttk.Label(emg_status_frame, text="当前EMG值:").pack(side=tk.LEFT, padx=5)
        self.current_emg_value = ttk.Label(emg_status_frame, text="--", font=("SimHei", 10, "bold"))
        self.current_emg_value.pack(side=tk.LEFT, padx=5)

        ttk.Label(emg_status_frame, text="传感器状态:").pack(side=tk.LEFT, padx=5)
        self.sensor_status = ttk.Label(emg_status_frame, text="未连接", foreground="red")
        self.sensor_status.pack(side=tk.LEFT, padx=5)

        # 阈值控制
        threshold_frame = ttk.Frame(frame)
        threshold_frame.pack(fill=tk.X, pady=5)

        ttk.Label(threshold_frame, text="EMG阈值:").pack(side=tk.LEFT, padx=5)
        self.threshold_entry = ttk.Entry(threshold_frame, width=10)
        self.threshold_entry.pack(side=tk.LEFT, padx=5)
        self.threshold_entry.insert(0, "30")

        self.set_threshold_btn = ttk.Button(threshold_frame, text="设置阈值",
                                            command=lambda: self.send_command(f"T{self.threshold_entry.get()}"))
        self.set_threshold_btn.pack(side=tk.LEFT, padx=5)

        self.get_threshold_btn = ttk.Button(threshold_frame, text="获取当前阈值",
                                            command=lambda: self.send_command("getthreshold"))
        self.get_threshold_btn.pack(side=tk.LEFT, padx=5)

        self.get_threshold_btn = ttk.Button(threshold_frame, text="跳转至肌电信号采集页面",
                                            command=self.open_second_window)
        self.get_threshold_btn.pack(side=tk.LEFT, padx=5)



    def create_log_frame(self):
        frame = ttk.LabelFrame(self.control_panel, text="通信日志", padding="10")
        frame.pack(fill=tk.BOTH, expand=True, pady=5)

        log_controls = ttk.Frame(frame)
        log_controls.pack(fill=tk.X, pady=5)

        self.clear_log_btn = ttk.Button(log_controls, text="清空日志", command=self.clear_log)
        self.clear_log_btn.pack(side=tk.RIGHT, padx=5)

        self.log_text = scrolledtext.ScrolledText(frame, wrap=tk.WORD, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)

    def create_visualization_frame(self):
        # 创建数据可视化面板
        viz_frame = ttk.LabelFrame(self.data_panel, text="数据可视化", padding="10")
        viz_frame.pack(fill=tk.BOTH, expand=True)

        # 创建EMG数据图表
        self.emg_fig = plt.Figure(figsize=(5, 4), dpi=100)
        self.emg_ax = self.emg_fig.add_subplot(111)
        self.emg_ax.set_title("EMG信号实时监测")
        self.emg_ax.set_xlabel("时间")
        self.emg_ax.set_ylabel("EMG值")
        self.emg_ax.grid(True)
        self.emg_line, = self.emg_ax.plot([], [], 'b-')
        self.threshold_line = self.emg_ax.axhline(y=30, color='r', linestyle='--', label='阈值')
        self.emg_ax.legend()

        self.emg_canvas = FigureCanvasTkAgg(self.emg_fig, master=viz_frame)
        self.emg_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # 图表控制按钮
        viz_controls = ttk.Frame(viz_frame)
        viz_controls.pack(fill=tk.X, pady=5)

        self.start_plot_btn = ttk.Button(viz_controls, text="开始绘图", command=self.start_plotting)
        self.start_plot_btn.pack(side=tk.LEFT, padx=5)

        self.stop_plot_btn = ttk.Button(viz_controls, text="停止绘图", command=self.stop_plotting, state=tk.DISABLED)
        self.stop_plot_btn.pack(side=tk.LEFT, padx=5)

        self.clear_plot_btn = ttk.Button(viz_controls, text="清空图表", command=self.clear_plot)
        self.clear_plot_btn.pack(side=tk.LEFT, padx=5)

        self.plotting_active = False

    def log(self, message):
        """在日志区域显示消息"""
        self.log_text.config(state=tk.NORMAL)
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def clear_log(self):
        """清空日志区域"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.log("日志已清空")

    def connect(self):
        """连接到WiFi设备"""
        ip = self.ip_entry.get()
        port = self.port_entry.get()

        if not ip or not port:
            messagebox.showerror("错误", "请输入IP地址和端口")
            return

        try:
            port = int(port)
            self.client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.client_socket.connect((ip, port))
            self.connected = True

            # 更新UI状态
            self.connect_btn.config(state=tk.DISABLED)
            self.disconnect_btn.config(state=tk.NORMAL)
            self.connection_status.config(text="已连接", foreground="green")
            self.log(f"已连接到 {ip}:{port}")

            # 启动接收线程
            self.receive_thread = threading.Thread(target=self.receive_data, daemon=True)
            self.receive_thread.start()

        except Exception as e:
            self.log(f"连接失败: {str(e)}")
            messagebox.showerror("连接失败", f"无法连接到设备: {str(e)}")
            self.client_socket = None
            self.connected = False

    def disconnect(self):
        """断开与WiFi设备的连接"""
        if self.connected and self.client_socket:
            try:
                self.client_socket.close()
            except Exception as e:
                self.log(f"断开连接错误: {str(e)}")

            self.connected = False
            self.client_socket = None

            # 更新UI状态
            self.connect_btn.config(state=tk.NORMAL)
            self.disconnect_btn.config(state=tk.DISABLED)
            self.connection_status.config(text="未连接", foreground="red")
            self.log("已断开连接")

    def send_command(self, command):
        """发送命令到设备"""
        if not self.connected or not self.client_socket:
            messagebox.showwarning("未连接", "请先连接到设备")
            return

        try:
            self.client_socket.sendall((command + "\n").encode('utf-8'))
            self.log(f"发送命令: {command}")
        except Exception as e:
            self.log(f"发送命令失败: {str(e)}")
            messagebox.showerror("发送失败", f"无法发送命令: {str(e)}")
            self.disconnect()

    def get_time_parameters(self):
        """获取时间参数"""
        self.send_command("time2")
        self.send_command("time3")
        self.send_command("time4")

    def get_speed_parameters(self):
        """获取速度参数"""
        self.send_command("speedup")
        self.send_command("speeddown")

    def reset_calibration(self):
        """重置校准数据"""
        self.relax_avg.config(text="--")
        self.contract_avg.config(text="--")
        self.recommended_threshold.config(text="--")
        self.calib_status.config(text="校准状态: 已重置")
        self.log("已重置校准数据")

    def start_plotting(self):
        """开始绘制EMG数据"""
        self.plotting_active = True
        self.start_plot_btn.config(state=tk.DISABLED)
        self.stop_plot_btn.config(state=tk.NORMAL)
        self.log("开始绘制EMG数据")
        self.update_plot()

    def stop_plotting(self):
        """停止绘制EMG数据"""
        self.plotting_active = False
        self.start_plot_btn.config(state=tk.NORMAL)
        self.stop_plot_btn.config(state=tk.DISABLED)
        self.log("停止绘制EMG数据")

    def clear_plot(self):
        """清空图表数据"""
        self.emg_data = []
        self.emg_time = []
        self.emg_line.set_xdata([])
        self.emg_line.set_ydata([])
        self.emg_ax.relim()
        self.emg_ax.autoscale_view()
        self.emg_canvas.draw()
        self.log("已清空图表数据")

    def update_plot(self):
        """更新EMG数据图表"""
        if self.plotting_active and self.emg_data and self.emg_time:
            self.emg_line.set_xdata(self.emg_time)
            self.emg_line.set_ydata(self.emg_data)

            # 更新阈值线
            threshold = float(self.threshold_entry.get())
            self.threshold_line.set_ydata([threshold])

            # 自动调整坐标轴范围
            self.emg_ax.relim()
            self.emg_ax.autoscale_view()

            self.emg_canvas.draw()

        # 继续更新图表
        if self.plotting_active:
            self.root.after(100, self.update_plot)

    def receive_data(self):
        """接收设备发送的数据"""
        buffer = ""
        while self.running and self.connected and self.client_socket:
            try:
                data = self.client_socket.recv(1024).decode('utf-8')
                if not data:
                    self.log("连接已关闭")
                    self.root.after(0, self.disconnect)
                    break

                buffer += data
                # 处理完整的消息（以换行符分隔）
                while '\n' in buffer:
                    line, buffer = buffer.split('\n', 1)
                    line = line.strip()
                    if line:
                        self.log(f"收到数据: {line}")
                        self.process_received_data(line)

            except Exception as e:
                self.log(f"接收数据错误: {str(e)}")
                self.root.after(0, self.disconnect)
                break

    def process_received_data(self, data):
        """处理接收到的数据并更新UI"""
        if data.startswith("EMG_MODE_ACTIVE"):
            self.root.after(0, lambda: self.current_change.config(text="当前模式: EMG模式"))
        elif data.startswith("WIFI_MODE_ACTIVE"):
            self.root.after(0, lambda: self.current_change.config(text="当前模式: WiFi模式"))
        elif data.startswith("CALIB_START:RELAX"):
            self.root.after(0, lambda: self.calib_status.config(text="校准状态: 请放松肌肉10秒"))
        elif data.startswith("CALIB_STAGE:CONTRACT"):
            self.root.after(0, lambda: self.calib_status.config(text="校准状态: 请用力收缩肌肉5秒"))
        elif data.startswith("RELAX,"):
            remaining = data.split(",")[1]
            self.root.after(0, lambda: self.calib_status.config(text=f"校准状态: 放松阶段，剩余{remaining}秒"))
        elif data.startswith("CONTRACT,"):
            remaining = data.split(",")[1]
            self.root.after(0, lambda: self.calib_status.config(text=f"校准状态: 用力阶段，剩余{remaining}秒"))
        elif data.startswith("Relax="):
            value = data.split("=")[1]
            self.root.after(0, lambda: self.relax_avg.config(text=value))
        elif data.startswith("Contract="):
            value = data.split("=")[1]
            self.root.after(0, lambda: self.contract_avg.config(text=value))
        elif data.startswith("Recommend="):
            value = data.split("=")[1]
            self.root.after(0, lambda: self.recommended_threshold.config(text=value))
        elif data.startswith("THRESHOLD:"):
            value = data.split(":")[1]
            self.root.after(0, lambda: self.threshold_entry.delete(0, tk.END))
            self.root.after(0, lambda: self.threshold_entry.insert(0, value))
        elif data.startswith("THRESHOLD_APPLIED:"):
            value = data.split(":")[1]
            self.root.after(0, lambda: self.threshold_entry.delete(0, tk.END))
            self.root.after(0, lambda: self.threshold_entry.insert(0, value))
        elif data.startswith("CALIBRATION_STARTED"):
            self.root.after(0, lambda: self.calib_status.config(text="校准状态: 校准已开始"))
        elif data.startswith("SPEED_SET:"):
            value = data.split(":")[1]
            self.root.after(0, lambda: self.speed_entry.delete(0, tk.END))
            self.root.after(0, lambda: self.speed_entry.insert(0, value))
        elif data.startswith("SPEEDUP:"):
            value = data.split(":")[1]
            self.root.after(0, lambda: self.speedup1_entry.delete(0, tk.END))
            self.root.after(0, lambda: self.speedup1_entry.insert(0, value))
        elif data.startswith("SPEEDDOWN:"):
            value = data.split(":")[1]
            self.root.after(0, lambda: self.speeddown1_entry.delete(0, tk.END))
            self.root.after(0, lambda: self.speeddown1_entry.insert(0, value))
        elif data.startswith("TIME2:"):
            value = data.split(":")[1]
            self.root.after(0, lambda: self.time2_entry.delete(0, tk.END))
            self.root.after(0, lambda: self.time2_entry.insert(0, value))
        elif data.startswith("TIME3:"):
            value = data.split(":")[1]
            self.root.after(0, lambda: self.time3_entry.delete(0, tk.END))
            self.root.after(0, lambda: self.time3_entry.insert(0, value))
        elif data.startswith("TIME4:"):
            value = data.split(":")[1]
            self.root.after(0, lambda: self.time4_entry.delete(0, tk.END))
            self.root.after(0, lambda: self.time4_entry.insert(0, value))
        elif data.startswith("EMG值:"):
            # 提取EMG值并更新图表
            try:
                emg_value_str = data.split(": ")[1].split(",")[0]
                emg_value = float(emg_value_str)
                sensor_state = "已连接" if "已连接" in data else "未连接"
                sensor_color = "green" if sensor_state == "已连接" else "red"

                self.root.after(0, lambda: self.current_emg_value.config(text=f"{emg_value:.1f}"))
                self.root.after(0, lambda: self.sensor_status.config(text=sensor_state, foreground=sensor_color))

                # 添加到数据列表用于绘图
                current_time = time.time()
                self.emg_data.append(emg_value)
                self.emg_time.append(current_time)

                # 保持数据点数量在限制范围内
                if len(self.emg_data) > self.max_data_points:
                    self.emg_data.pop(0)
                    self.emg_time.pop(0)
            except:
                pass
        elif data.startswith("ERROR:NO_RECOMMENDED_THRESHOLD"):
            self.root.after(0, lambda: messagebox.showerror("错误", "没有可用的推荐阈值，请先进行校准"))

    def on_close(self):
        """窗口关闭时的处理"""
        self.running = False
        self.plotting_active = False
        self.disconnect()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = WiFiHardwareController(root)
    root.mainloop()
