# -*- coding: utf-8 -*-
# monitor.py
import psutil                                 # 系统/进程信息
import matplotlib.pyplot as plt               # 绘图
import matplotlib
from matplotlib.animation import FuncAnimation# 动态更新
import time,sys

# ---------- 1. 要监控的进程名（或 PID） ----------
MAX_POINTS = 200                              # 折线图最多保留多少个点

# ---------- 2. 数据容器 ----------
x_data, cpu_data, mem_data = [], [], []       # 时间轴、CPU、内存

# ---------- 3. 找到目标进程 ----------
if len(sys.argv) > 1:          # 如果传了参数就用它
    pid = int(sys.argv[1])
else:
    pid = 4
proc = psutil.Process(pid)
print(f"已连接到 PID={pid}  {proc.name()}")

# ---------- 4. 采样函数 ----------
def sample(_):
    try:
        # 进程是否还活着
        if not proc.is_running():
            raise psutil.NoSuchProcess(pid)

        now = time.time()
        cpu = proc.cpu_percent()                  # CPU 占用百分比
        mem = proc.memory_info().rss / 1024 / 1024  # RSS 物理内存 MB

        # 追加最新数据
        x_data.append(now)
        cpu_data.append(cpu)
        mem_data.append(mem)

        # 只保留最近 MAX_POINTS 个
        if len(x_data) > MAX_POINTS:
            x_data.pop(0); cpu_data.pop(0); mem_data.pop(0)

        # 更新折线
        line_cpu.set_data(x_data, cpu_data)
        line_mem.set_data(x_data, mem_data)

        # 自动缩放坐标轴
        ax1.relim(); ax1.autoscale_view()
        ax2.relim(); ax2.autoscale_view()

        return line_cpu, line_mem

    except psutil.NoSuchProcess:
        # 目标进程已退出：关闭窗口并结束脚本
        plt.close(fig)        # 关闭图形界面
        return                # 直接返回，动画循环将自动结束

# ---------- 5. 创建画布 ----------
matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['font.sans-serif'] = ['SimHei']   # 黑体，Windows 自带
matplotlib.rcParams['axes.unicode_minus'] = False     # 让负号正常显示
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(6, 4), sharex=True)
fig.canvas.manager.set_window_title(f"实时监控 - {'python.exe'}")
line_cpu, = ax1.plot([], [], color='tab:blue', label='CPU %')
line_mem, = ax2.plot([], [], color='tab:orange', label='内存 MB')
ax1.set_ylabel("CPU %"); ax1.legend(); ax1.grid(True)
ax2.set_ylabel("内存 MB"); ax2.set_xlabel("时间戳 (s)"); ax2.legend(); ax2.grid(True)

# ---------- 6. 动态更新 ----------
ani = FuncAnimation(fig, sample, interval=1000, blit=False)
plt.show()