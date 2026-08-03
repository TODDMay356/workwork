# -*- coding: utf-8 -*-
"""
出单监控 — 工作台版 wrapper
用法: python 出单监控_workbench.py --today T.xlsx --yesterday Y.xlsx --password xxx --date 2026-07-31

不做任何 tkinter 交互，所有参数从命令行传入。
"""
import os
import sys
import argparse
from unittest.mock import patch

# ---- CLI 参数 ----
parser = argparse.ArgumentParser()
parser.add_argument("--today", required=True, help="今天的报表路径")
parser.add_argument("--yesterday", required=True, help="昨天的报表路径")
parser.add_argument("--password", required=True)
parser.add_argument("--date", required=True, help="报表业务日期 YYYY-MM-DD")
args = parser.parse_args()

# ---- Monkey-patch tkinter 为无头模式 ----
# 在 import 原脚本之前注入模拟值

# 1) 密码验证 → 用 CLI 参数
import tkinter.simpledialog as _sd
_askstring = _sd.askstring
def _mock_askstring(title, prompt, **kw):
    if "密码" in title:
        return args.password
    if "日期" in prompt:
        return args.date
    return None
_sd.askstring = _mock_askstring

# 2) 文件选择 → 用 CLI 参数
import tkinter.filedialog as _fd
_askopen = _fd.askopenfilename
def _mock_askopen(**kw):
    if "今天" in (kw.get("title") or ""):
        return args.today
    if "昨天" in (kw.get("title") or ""):
        return args.yesterday
    return None
_fd.askopenfilename = _mock_askopen

# 3) showinfo / showerror → 静默
import tkinter.messagebox as _mb
_mb.showinfo = lambda *a, **kw: None
_mb.showerror = lambda *a, **kw: None

# 4) Tk root → 不创建窗口
import tkinter as _tk
_orig_Tk = _tk.Tk
class _FakeTk:
    def withdraw(self): pass
    def destroy(self): pass
    def mainloop(self): pass
_tk.Tk = lambda *a, **kw: _FakeTk()

# ---- 切到原脚本目录，运行 ----
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)
sys.path.insert(0, SCRIPT_DIR)

# 动态加载原脚本
import importlib.util as _iu
spec = _iu.spec_from_file_location("order_monitor",
    os.path.join(SCRIPT_DIR, "出单监控.py"))
mod = _iu.module_from_spec(spec)
spec.loader.exec_module(mod)
# 显式调 main，然后停住窗口
try:
    mod.main()
except Exception as e:
    print(f"\n程序出错: {e}")
    import traceback
    traceback.print_exc()
print("\n--- 出单监控完成 ---")
input("按回车键关闭窗口...")
