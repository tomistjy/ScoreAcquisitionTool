# -*- coding: utf-8 -*-
"""
prepare.py  **中考成绩抓取
修正：关闭验证码窗口=中断；空验证码+确认=未输入验证码
"""

# ---------- 1. 标准库 & 第三方库导入 ----------
import re
import requests
import csv
import tkinter as tk
from tkinter import messagebox, ttk, Menu
from bs4 import BeautifulSoup
from PIL import Image, ImageTk
import sys, datetime, random, time
import ddddocr
import os
import subprocess
os.remove('log.txt') if os.path.exists('log.txt') else None
ocr_attempt = 0

# ---------- 2. 全局配置 ----------
USE_OCR = None
session = requests.Session()
post_address = picture_address = ''

# ---------- 3. 日志双通道 ----------
log = open('log.txt', 'a', encoding='utf-8')
class Dual:
    def write(self, txt):
        sys.__stdout__.write(txt)
        log.write(txt)
        log.flush()
    def flush(self):
        sys.__stdout__.flush()
sys.stdout = Dual()

# ---------- 4. 工具函数 ----------
def ask_ocr_mode():
    root = tk.Tk()
    root.withdraw()
    root.update()
    ans = messagebox.askyesno(
        "验证码识别方式",
        "是否使用 OCR 自动识别验证码？\n\n"
        "是 = 自动识别(最多重试10次)\n"
        "否 = 手动输入")
    root.destroy()
    return ans

def random_delay():
    time.sleep(random.uniform(0.05, 0.07))

def _error_line(zkz, mz, reason):
    return f"{zkz} {mz} " + (f"{reason} " * 17).strip() + " "

# ---------- 5. 验证码统一获取 ----------
def fetch_captcha(parent):
    global ocr_attempt
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

    if USE_OCR:
        while ocr_attempt < 10:
            ocr_attempt += 1
            with open("captcha.jpg", "rb") as f:
                code = ddddocr.DdddOcr(show_ad=False).classification(f.read()).strip()
            print(f"[OCR {ocr_attempt}/10] {code}")
            if code:
                ocr_attempt = 0
                return code, "ok"
            random_delay()
            try:
                img_resp = session.get(picture_address, timeout=10)
                img_resp.raise_for_status()
                with open("captcha.jpg", "wb") as f:
                    f.write(img_resp.content)
            except:
                return "", "abort"
        ocr_attempt = 0

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
def get_url():
    global post_address, picture_address
    url = '（不公开）'
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
    picture_address = '（不公开）' + captcha_url if captcha_url else ''
    post_address    = '（不公开）' + post_url    if post_url    else ''
    print('picture_address:', picture_address)
    print('post_address:', post_address)

# ---------- 7. 查询接口 ----------
def get_score_line(zkz: str, mz: str, yzm: str) -> str:
    if not post_address:
        return _error_line(zkz, mz, "地址获取失败")
    headers = {
        "Host": "（不公开）",
        "X-Requested-With": "XMLHttpRequest",
        "User-Agent": "Mozilla/5.0",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Referer": "（不公开）",
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
    m = re.search(r'"score":"(.*?)"(?=,"remark)', resp.text)
    if not m or str(m.group(1)) == 'None':
        return _error_line(zkz, mz, "验证码错误")
    html = m.group(1).encode().decode('unicode_escape', errors='ignore')
    name_match = re.search(r'>(\d{6,10}\s[\u4e00-\u9fa5]+)', html)
    if not name_match:
        return _error_line(zkz, mz, "抓取失败")
    name_str = name_match.group(1).strip()
    parts_name = name_str.split(maxsplit=1)
    if len(parts_name) != 2:
        return _error_line(zkz, mz, "抓取失败")
    kv_str = re.search(r'<hr>([^<]+)', html)
    if not kv_str:
        return _error_line(zkz, mz, "抓取失败")
    kv_str = kv_str.group(1).replace('：', ':')
    scores = {}
    for piece in kv_str.split(','):
        if ':' not in piece:
            continue
        k, v = piece.split(':', 1)
        scores[k.strip()] = v.strip()
    remark_match = re.search(r'"remark":"([^"]*)"', resp.text)
    school = remark_match.group(1) if remark_match else ''
    school = school.encode().decode('unicode_escape', errors='ignore')
    order = ['总分', '语文', '数学', '英语', '体育', '物史', '化道', '生地',
             '音乐', '美术', '信息科技', '物理(含实验)', '历史',
             '化学(含实验)', '道法', '生物学(含实验)', '地理']
    return ' '.join(parts_name + [scores.get(k, ' ') for k in order] + [school])

# ---------- 8. 主界面 ----------
class ScoreApp:
    def __init__(self, master: tk.Tk):
        self.master = master
        master.title("成绩抓取工具")
        master.geometry("1400x480")
        self.input_csv = 'students.csv'
        self.output_csv = 'scores.csv'
        self.headers = ['准考证', '姓名', '总分', '语文', '数学', '英语',
                        '体育', '物史', '化道', '生地',
                        '音乐', '美术', '信息科技',
                        '物理(含实验)', '历史', '化学(含实验)', '道法', '生物学(含实验)', '地理', '录取学校']
        self.students = []
        self.all_scores = []
        self.current = 0
        self.aborted = False

        self.label = tk.Label(master, text="成绩抓取进度", font=("微软雅黑", 14))
        self.label.pack(pady=5)
        self.progress = ttk.Progressbar(master, length=500, mode='determinate')
        self.progress.pack(pady=5)

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
        self.tree.bind("<Button-3>", self.on_right_click)
        self.right_menu = Menu(self.tree, tearoff=0)
        self.right_menu.add_command(label="重试", command=self.retry_selected)

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
        tk.Button(btn_box, text="性能监控", command=self.open_monitor, **btn_style).pack(side='left', padx=6)

    # ---------- 9. 右键重试 ----------
    def on_right_click(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        self.tree.selection_set(row_id)
        self.right_menu.post(event.x_root, event.y_root)

    def retry_selected(self):
        if not self.tree.selection():
            return
        idx = self.tree.index(self.tree.selection()[0])
        stu = self.students[idx]
        self._retry_one(idx, stu['准考证'], stu['姓名'])

    # ---------- 10. 批量重试 ----------
    def retry_all_fail(self):
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

    def _retry_one(self, idx, zkz, mz):
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
    def start(self):
        self.aborted = False
        self.start_btn.config(state=tk.DISABLED)
        if self.current == 0:
            self.load_students()
            self.progress['maximum'] = len(self.students)
        self.master.after(100, self.process_next)

    def load_students(self):
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
        max_retry = 10
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

    def save_scores(self):
        with open(self.output_csv, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=self.headers)
            writer.writeheader()
            writer.writerows(self.all_scores)

    def open_monitor(self):
        monitor_path = os.path.join(os.path.dirname(__file__), 'performance_monitoring.py')
        subprocess.Popen([sys.executable, monitor_path, str(os.getpid())])

# ---------- 12. 程序入口 ----------
if __name__ == '__main__':
    root = tk.Tk()
    USE_OCR = ask_ocr_mode()
    app = ScoreApp(root)
    root.mainloop()
