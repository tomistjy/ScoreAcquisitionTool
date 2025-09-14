# -*- coding: utf-8 -*-                       # 声明源文件编码为 UTF-8，防止中文乱码
"""
prepare.py  湛江中考成绩抓取                # 模块文档字符串：一句话描述本文件用途
修正：关闭验证码窗口=中断；空验证码+确认=未输入验证码
"""

# ---------- 1. 标准库 & 第三方库导入 ----------
import re                                     # 正则表达式，提取网页中的关键字符串
import requests                               # HTTP 客户端库，负责和服务器通信
import csv                                    # CSV 读写库，保存成绩数据
import tkinter as tk                          # Tk GUI 框架，做桌面窗口
from tkinter import messagebox, ttk, Menu     # Tk 扩展组件：对话框、表格、右键菜单
from bs4 import BeautifulSoup                 # HTML 解析器，快速定位标签
from PIL import Image, ImageTk                # Pillow：图片加载、缩放、显示到 Tk
import sys, datetime, random, time            # 系统、日期、随机数、睡眠
import ddddocr                                # OCR 离线识别验证码
import os                                     # 【新增】用于删除旧日志
import subprocess
os.remove('log.txt') if os.path.exists('log.txt') else None  # 【新增】直接删除旧日志
ocr_attempt = 0          # 全局计数器：OCR 累计重试次数

# ---------- 2. 全局配置 ----------
USE_OCR = None                                # 程序启动时由用户选择：True=自动识别 / False=手动输入
session = requests.Session()                  # 复用 TCP 连接，带 Cookie 的会话对象
post_address = picture_address = ''           # 动态抓到的 POST 地址 & 验证码地址

# ---------- 3. 日志双通道 ----------
log = open('log.txt', 'a', encoding='utf-8')  # 追加写日志文件
class Dual:                                   # 【类】同时向终端和文件写日志
    def write(self, txt):                     # 重定向 sys.stdout.write
        sys.__stdout__.write(txt)             # 控制台
        log.write(txt)                        # 文件
        log.flush()                           # 立即落盘
    def flush(self):                          # flush 接口，保持兼容
        sys.__stdout__.flush()
sys.stdout = Dual()                           # 替换默认 stdout，实现“双通道”日志

# ---------- 4. 工具函数 ----------
def ask_ocr_mode():
    root = tk.Tk()
    root.withdraw()      # 隐藏白窗口
    root.update()        # 立即生效
    ans = messagebox.askyesno(
        "验证码识别方式",
        "是否使用 OCR 自动识别验证码？\n\n"
        "是 = 自动识别(最多重试10次)\n"
        "否 = 手动输入")
    root.destroy()
    return ans

def random_delay():                           # 【函数】随机 50-70 ms 缓冲，防请求过快
    time.sleep(random.uniform(0.05, 0.07))

def _error_line(zkz, mz, reason):
    # 生成20个字段：准考证、姓名 + 17个错误原因 + 录取学校为空
    return f"{zkz} {mz} " + (f"{reason} " * 17).strip() + " "
# ---------- 5. 验证码统一获取 ----------
def fetch_captcha(parent):
    """
    返回 (yzm:str, flag:str)
    flag = "ok" | "empty" | "abort"
    """
    global ocr_attempt       # 引用全局计数器

    get_url()
    if not picture_address:
        return "", "abort"

    try:
        img_resp = session.get(picture_address, timeout=10)
        img_resp.raise_for_status()
        with open("captcha.jpg", "wb") as f:
            f.write(img_resp.content)
    except Exception as e:
        print("验证码下载失败:", e)
        return "", "abort"

    # ---- OCR 分支 ----
    if USE_OCR:
        while ocr_attempt < 10:
            ocr_attempt += 1
            with open("captcha.jpg", "rb") as f:
                code = ddddocr.DdddOcr(show_ad=False).classification(f.read()).strip()
            print(f"[OCR {ocr_attempt}/10] {code}")
            if code:
                ocr_attempt = 0               # 成功后重置
                return code, "ok"
            # 失败 -> 刷新验证码
            random_delay()
            try:
                img_resp = session.get(picture_address, timeout=10)
                img_resp.raise_for_status()
                with open("captcha.jpg", "wb") as f:
                    f.write(img_resp.content)
            except:
                return "", "abort"
        # 10 次用完 -> 转手动
        ocr_attempt = 0

    # ---- 手动弹窗 ----
    top = tk.Toplevel(parent)
    top.title("验证码")
    top.attributes('-topmost', True)
    top.grab_set(); top.focus_force()

    img = Image.open("captcha.jpg").resize(
        (lambda w, h: (w*3, h*3))(*Image.open("captcha.jpg").size), Image.LANCZOS)
    tk_img = ImageTk.PhotoImage(img)
    top.tk_img = tk_img
    tk.Label(top, image=tk_img).pack()

    entry = tk.Entry(top, font=('Consolas', 19), justify='center')
    entry.pack(pady=5); entry.focus()

    result = tk.StringVar(value="__ABORT__")

    def on_ok(_=None):
        val = entry.get().strip()
        result.set(val if val else "__EMPTY__")
        top.destroy()
    def on_close(_=None):
        result.set("__ABORT__")
        top.destroy()

    entry.bind('<Return>', on_ok)
    top.protocol("WM_DELETE_WINDOW", on_close)

    btn = tk.Frame(top); btn.pack(pady=5)
    tk.Button(btn, text="确认", command=on_ok, width=8).pack(side='left', padx=5)
    tk.Button(btn, text="关闭", command=on_close, width=8).pack(side='left', padx=5)

    top.update_idletasks()
    x = (top.winfo_screenwidth() - top.winfo_width()) // 2
    y = (top.winfo_screenheight() - top.winfo_height()) // 2
    top.geometry(f"+{x}+{y}")
    top.wait_window()

    val = result.get()
    if val == "__ABORT__":
        return "", "abort"
    if val == "__EMPTY__":
        return "", "empty"
    return val, "ok"

# ---------- 6. 地址抓取 ----------
def get_url():                                # 【函数】访问首页，刷新 cookie 与地址
    global post_address, picture_address
    url = 'http://zk.jyj.zhanjiang.gov.cn/'
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/138.0.0.0 Safari/537.36 Edg/138.0.0.0"
    }
    try:
        html = session.get(url, headers=headers, timeout=30).text
        print(html)
    except Exception as e:
        print("首页失败:", e)
        post_address = picture_address = ''
        return

    soup = BeautifulSoup(html, 'html.parser')
    post_pattern = re.compile(r'form\.on\("submit\(_studentScore\)",\s*function\(data\)[^"]+"([^"]+)"')
    m = post_pattern.search(html)
    post_url = m.group(1) if m else None

    captcha_img = soup.find('img', id='vd_score')
    captcha_url = captcha_img['src'] if captcha_img else None

    picture_address = 'http://zk.jyj.zhanjiang.gov.cn/' + captcha_url if captcha_url else ''
    post_address    = 'http://zk.jyj.zhanjiang.gov.cn/' + post_url    if post_url    else ''
    print('picture_address:',picture_address)
    print('post_address:',post_address)
    

# ---------- 7. 查询接口 ----------
def get_score_line(zkz: str, mz: str, yzm: str) -> str:
    """查询一条成绩，返回固定 19 列的字符串：姓名 + 17 科"""
    if not post_address:
        return _error_line(zkz, mz, "地址获取失败")

    headers = {
        "Host": "zk.jyj.zhanjiang.gov.cn",
        "X-Requested-With": "XMLHttpRequest",
        "User-Agent": "Mozilla/5.0",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Referer": "http://zk.jyj.zhanjiang.gov.cn/",
    }
    body = {"iksh": zkz, "ixm": mz, "ivalidate": yzm, "access_token": "zjzk201911"}

    try:
        resp = session.post(post_address, headers=headers, data=body, timeout=5)
        print(resp.text)
        resp.raise_for_status()
    except Exception as e:
        print("查询接口错误:", e)
        return _error_line(zkz, mz, "查询失败")

    txt = resp.text.encode().decode('unicode_escape', errors='ignore')
    print(zkz, mz, yzm, txt)

    # 1. 提取成绩 HTML
    m = re.search(r'"score":"(.*?)"(?=,"remark)', resp.text)
    if not m or str(m.group(1)) == 'None':
        return _error_line(zkz, mz, "验证码错误")
    html = m.group(1).encode().decode('unicode_escape', errors='ignore')

    # 2. 分离准考证和姓名
    name_match = re.search(r'>(\d{6,10}\s[\u4e00-\u9fa5]+)', html)
    if not name_match:
        return _error_line(zkz, mz, "抓取失败")
    name_str = name_match.group(1).strip()
    parts_name = name_str.split(maxsplit=1)  # 分割成 [准考证, 姓名]
    if len(parts_name) != 2:
        return _error_line(zkz, mz, "抓取失败")

    # 3. 提取科目成绩
    kv_str = re.search(r'<hr>([^<]+)', html)
    if not kv_str:
        return _error_line(zkz, mz, "抓取失败")
    kv_str = kv_str.group(1)
    kv_str = kv_str.replace('：', ':')
    scores = {}
    for piece in kv_str.split(','):
        if ':' not in piece:
            continue
        k, v = piece.split(':', 1)
        scores[k.strip()] = v.strip()
    # 4. 提取录取学校
    remark_match = re.search(r'"remark":"([^"]*)"', resp.text)
    school = remark_match.group(1) if remark_match else ''
    school = school.encode().decode('unicode_escape', errors='ignore')

    # 5. 按顺序构建20个字段（含录取学校）
    order = ['总分', '语文', '数学', '英语', '体育', '物史', '化道', '生地',
             '音乐', '美术', '信息科技', '物理(含实验)', '历史', 
             '化学(含实验)', '道法', '生物学(含实验)', '地理']
    
    return ' '.join(parts_name + [scores.get(k, ' ') for k in order] + [school])


# ---------- 8. 主界面 ----------
class ScoreApp:                               # 【类】整个 GUI 主程序
    def __init__(self, master: tk.Tk):        # 【方法】构造器
        self.master = master                  # 保存根窗口引用
        master.title("成绩抓取工具")
        master.geometry("1400x480")
        self.input_csv = 'students.csv'       # 默认学生名单
        self.output_csv = 'scores.csv'        # 默认输出文件
        self.headers = ['准考证', '姓名', '总分', '语文', '数学', '英语',
                '体育', '物史', '化道', '生地',
                '音乐', '美术', '信息科技',
                '物理(含实验)', '历史', '化学(含实验)', '道法', '生物学(含实验)', '地理']
        self.headers.append('录取学校')
        self.students = []                    # 所有学生列表
        self.all_scores = []                  # 所有成绩列表
        self.current = 0                      # 当前抓取索引
        self.aborted = False                  # 用户是否点击“中断”

        # 进度条
        self.label = tk.Label(master, text="成绩抓取进度", font=("微软雅黑", 14))
        self.label.pack(pady=5)
        self.progress = ttk.Progressbar(master, length=500, mode='determinate')
        self.progress.pack(pady=5)

        # 表格
        tree_frame = tk.Frame(master)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.tree = ttk.Treeview(tree_frame, columns=self.headers,
                                 show='headings', height=10)
        for h in self.headers:
            self.tree.heading(h, text=h)
            self.tree.column(h, width=80, anchor='center')
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<Button-3>", self.on_right_click)  # 右键菜单
        self.right_menu = Menu(self.tree, tearoff=0)
        self.right_menu.add_command(label="重试", command=self.retry_selected)

        # 底部按钮
        bottom = tk.Frame(master)
        bottom.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        tk.Label(bottom, text="TOM", font=('TkDefaultFont', 8)).pack(side=tk.LEFT, padx=10)
        btn_box = tk.Frame(bottom)
        btn_box.pack(expand=True)
        btn_style = {"width": 10, "height": 1, "font": ("微软雅黑", 10)}
        tk.Button(btn_box, text="保存", command=self.save_scores, **btn_style).pack(side='left', padx=6)
        self.start_btn = tk.Button(btn_box, text="开始抓取", command=self.start, **btn_style)
        self.start_btn.pack(side='left', padx=6)
        self.retry_btn = tk.Button(btn_box, text="重试", command=self.retry_all_fail,
                                   state=tk.DISABLED, **btn_style)
        self.retry_btn.pack(side='left', padx=6)
        tk.Button(btn_box, text="退出", command=master.destroy, **btn_style).pack(side='left', padx=6)

        #性能监控
        tk.Button(btn_box, text="性能监控", command=self.open_monitor, **btn_style).pack(side='left', padx=6)

    # ---------- 9. 右键重试 ----------
    def on_right_click(self, event):            # 【回调】右键菜单
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        self.tree.selection_set(row_id)
        self.right_menu.post(event.x_root, event.y_root)

    def retry_selected(self):                   # 【方法】单条右键重试
        if not self.tree.selection():
            return
        idx = self.tree.index(self.tree.selection()[0])
        stu = self.students[idx]
        self._retry_one(idx, stu['准考证'], stu['姓名'])

    # ---------- 10. 批量重试 ----------
    def retry_all_fail(self):                   # 【方法】批量重试所有失败行
        fail_rows = [i for i, d in enumerate(self.all_scores)
                     if any(d[h] in {"地址获取失败", "查询失败", "抓取失败", "验证码下载失败",
                                     "未输入验证码", "验证码错误"}
                            for h in self.headers[1:])]
        if not fail_rows:
            messagebox.showinfo("提示", "没有需要重试的失败项。")
            return
        if messagebox.askyesno("批量重试", f"共 {len(fail_rows)} 条失败，是否重试？"):
            for idx in fail_rows:
                stu = self.students[idx]
                self._retry_one(idx, stu['准考证'], stu['姓名'])
                self.master.update()

    def _retry_one(self, idx, zkz, mz):       # 【方法】真正重试一条
        yzm, flag = fetch_captcha(self.master)
        if flag == "abort":
            messagebox.showinfo("中断", "用户取消，重试结束。")
            return
        new_line = get_score_line(zkz, mz, yzm) if flag == "ok" else _error_line(zkz, mz, "未输入验证码")
        parts = new_line.strip().split()
        score_data = dict(zip(self.headers, parts[:20] + [''] * (20 - len(parts))))
        if idx < len(self.all_scores):
            self.all_scores[idx] = score_data
            self.tree.item(self.tree.get_children()[idx],
                           values=[score_data[h] for h in self.headers])
        else:
            self.all_scores.append(score_data)
            self.tree.insert('', tk.END, values=[score_data[h] for h in self.headers])
        self.save_scores()

    # ---------- 11. 开始/继续 ----------
    def start(self):                            # 【方法】开始或继续抓取
        self.aborted = False
        self.start_btn.config(state=tk.DISABLED)
        if self.current == 0:                 # 首次运行
            self.load_students()
            self.progress['maximum'] = len(self.students)
        self.master.after(100, self.process_next)

    def load_students(self):                  # 【方法】加载学生名单
        with open(self.input_csv, newline='', encoding='utf-8') as f:
            self.students = [row for row in csv.DictReader(f)]

    def process_next(self):
        if self.current >= len(self.students):
            self.save_scores()
            self.start_btn.config(text="开始抓取", state=tk.NORMAL)
            self.retry_btn.config(state=tk.NORMAL)
            messagebox.showinfo("完成", "成绩已全部写入 scores.csv")
            return
        if self.aborted:
            self.start_btn.config(text="继续抓取", state=tk.NORMAL)
            self.retry_btn.config(state=tk.NORMAL)
            return

        stu = self.students[self.current]
        zkz, mz = stu['准考证'], stu['姓名']
        max_retry = 10                          # 真正重试次数
        retry_count = 0

        while retry_count < max_retry:
            yzm, flag = fetch_captcha(self.master)
            if flag == "abort":
                self.aborted = True
                self.start_btn.config(text="继续抓取", state=tk.NORMAL)
                self.retry_btn.config(state=tk.NORMAL)
                return
            if flag == "empty":
                new_line = _error_line(zkz, mz, "未输入验证码")
                break

            new_line = get_score_line(zkz, mz, yzm)
            if new_line != _error_line(zkz, mz, "验证码错误"):
                break

            retry_count += 1
            print(f"[验证码错误] 第 {retry_count} 次重试")
            random_delay()
        else:
            new_line = _error_line(zkz, mz, "验证码错误")

        parts = new_line.strip().split()
        score_data = dict(zip(self.headers, parts[:20] + [''] * (20 - len(parts))))
        if self.current < len(self.all_scores):
            self.all_scores[self.current] = score_data
            self.tree.item(self.tree.get_children()[self.current],
                           values=[score_data[h] for h in self.headers])
        else:
            self.all_scores.append(score_data)
            self.tree.insert('', tk.END, values=[score_data[h] for h in self.headers])

        self.current += 1
        self.progress['value'] = self.current
        random_delay()
        self.master.after(100, self.process_next)

    def save_scores(self):                    # 【方法】立即写入 CSV
        with open(self.output_csv, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=self.headers)
            writer.writeheader()
            writer.writerows(self.all_scores)
    def open_monitor(self):
        monitor_path = os.path.join(os.path.dirname(__file__), 'performance_monitoring.py')
        # 把当前进程 PID 作为第一个参数
        subprocess.Popen([sys.executable, monitor_path, str(os.getpid())])

# ---------- 12. 程序入口 ----------
if __name__ == '__main__':                    # 【脚本入口】仅在直接运行时触发
    root = tk.Tk()                            # 【对象】根窗口
    USE_OCR = ask_ocr_mode()                  # 先让用户选择 OCR/手动
    app = ScoreApp(root)                      # 【对象】实例化主界面
    root.mainloop()                           # 【方法】进入 Tk 主事件循环