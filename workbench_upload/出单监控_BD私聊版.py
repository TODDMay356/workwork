# -*- coding: utf-8 -*-
"""
出单监控_BD私聊版.py （修复版）

相对上一版的改动（仅 3 处核心修复 + 配套）：
1. 删除本脚本自带的空 FEISHU_APP_ID/SECRET 和自带 get_tenant_access_token，
   私聊与多维表格共用【原脚本同一个 App / 同一个 token】，
   彻底消除"空凭证导致整段失败"的问题，且 token 只取一次。
2. token / 多维表格 / BD私聊 / 本地保存 各自独立 try/except：
   任一环节失败都不影响其余环节；本地 txt 保存放最后且保证执行。
3. 周报是否发送，按【报表业务日】判断星期，而非脚本运行当天。
配套：日报/周报各自只渲染自己的状态分类；split 不再硬解包 8 元组。

使用方式：本文件与原始【出单监控.py】放同一文件夹后运行。
"""

import os
import sys
import json
import importlib.util
from datetime import datetime

import requests


# ======================== 需要你补充：BD 邮箱 ========================
# 只把【要发送】的 BD 放进来；不在此表里的 BD（如 陆明、丁xx、未分配BD）会被自动跳过、不发。
# 名字必须与 Excel【所属BD】完全一致；邮箱用 BD 登录飞书的真实邮箱（@inflyway.com）。
# 留空的会被安全跳过（不会乱发），填好谁就发给谁。
BD_EMAILS = {
    "唐丽": "tangli@inflyway.com",
    "左雅婧": "zuoyajing@camelfin.com",
    "车姚玲": "cheyaoling@inflyway.com",
    "朱佳吉": "zhujiaji@inflyway.com",
    "吴鑫彪": "wuxinbiao@inflyway.com",
}

# ======================== 周报发送日 ========================
# 0=周一 1=周二 ... 6=周日
WEEKLY_SEND_WEEKDAY = 0

# ======================== 发送范围 ========================
# None = 给 BD_EMAILS 里所有 BD 发送（正式放量）。
# 若要临时只发某一个 BD 做验证，把它设成该 BD 名字，如 "唐丽"。
TEST_BD = None

# ======================== 原脚本路径 ========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ORIGINAL_SCRIPT = os.path.join(BASE_DIR, "出单监控.py")

# ======================== 状态分类 ========================
# 顺序 = BD 跟进漏斗：先看要追的(待出单/未出单)，再看进展(试水)，最后看达成(首笔出单)
DAILY_STATUS = [
    "🆕 新审核通过-待出单",
    "👀 3天内未出单",
    "🐢 小额滞留",
    "🧪 测试交易",
    "✅ 新出单",
]
WEEKLY_STATUS = [
    "🔴 3-30天未出单",
    "🟠 30-60天未出单",
]

# 展示标签：对齐 BD 生命周期词汇（原脚本 notify 的 key 不变，只在文案层映射）
STATUS_DISPLAY = {
    "🆕 新审核通过-待出单": "审核通过-待出单（通道已通过审批）",
    "👀 3天内未出单": "未出单·3天内",
    "🐢 小额滞留": "小额滞留（长期累计TPV<$100）",
    "🧪 测试交易": "测试交易（今日开始有成功交易且累计TPV<$100）",
    "✅ 新出单": "新出单（今日累计TPV≥$100）",
    "🔴 3-30天未出单": "未出单·3-30天",
    "🟠 30-60天未出单": "未出单·30-60天",
}
# 概览行简称
STATUS_SHORT = {
    "🆕 新审核通过-待出单": "待出单",
    "👀 3天内未出单": "3天内",
    "🐢 小额滞留": "滞留",
    "🧪 测试交易": "测试",
    "✅ 新出单": "新出单",
    "🔴 3-30天未出单": "3-30天",
    "🟠 30-60天未出单": "30-60天",
}
# 未出单类：显示"已开通X天未出单"，不显示无意义的 TPV:$0
UNPAID_STATUS = {"👀 3天内未出单", "🔴 3-30天未出单", "🟠 30-60天未出单"}
# 刚过审待上量：还没交易，显示"待上量"
APPROVED_WAITING = "🆕 新审核通过-待出单"
# 小额滞留：试水后停滞，单独渲染
STUCK_STATUS = "🐢 小额滞留"


# ======================== 加载原始脚本 ========================
def load_original_module():
    if not os.path.exists(ORIGINAL_SCRIPT):
        raise FileNotFoundError(
            f"找不到原始脚本：{ORIGINAL_SCRIPT}\n"
            f"请确认本文件和【出单监控.py】放在同一个文件夹。"
        )
    spec = importlib.util.spec_from_file_location("original_monitor", ORIGINAL_SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


# ======================== 飞书私聊发送 ========================
def send_private_message(token, email, text):
    """通过 BD 邮箱直接私聊发送文本消息（receive_id_type=email，无需 open_id）。"""
    url = "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=email"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json; charset=utf-8",
    }
    payload = {
        "receive_id": email,
        "msg_type": "text",
        "content": json.dumps({"text": text}, ensure_ascii=False),
    }
    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=30)
        data = resp.json()
    except Exception as e:
        print(f"  ❌ 飞书私聊请求失败: {e}")
        return False
    if data.get("code") != 0:
        print(f"  ❌ 飞书私聊发送失败: {data}")
        return False
    print("  ✅ 飞书私聊发送成功")
    return True


# ======================== BD 分组工具 ========================
def clean_bd_name(bd):
    bd = str(bd).strip()
    if bd.lower() in ["", "nan", "none", "null"]:
        return "未分配BD"
    return bd


def split_notify_by_bd(notify, target_status):
    """按 BD 拆分 notify。只取 item[3]=BD，避免硬解包整条元组。"""
    bd_map = {}
    for status in target_status:
        for item in notify.get(status, []):
            bd_name = clean_bd_name(item[3])
            bd_map.setdefault(bd_name, {}).setdefault(status, []).append(item)
    return bd_map


# ======================== 消息内容生成 ========================
def _fmt_date(d):
    return d.strftime("%m-%d") if d else "-"


def build_overview(bd_notify, status_order):
    """BD 当日概览行：按漏斗顺序列出各桶数量。"""
    parts = [f"{STATUS_SHORT.get(s, s)} {len(bd_notify.get(s, []))}" for s in status_order]
    return "概览：" + " | ".join(parts)


def build_bd_message(bd_name, bd_notify, report_date_str, title, status_order):
    """只渲染 status_order 内的分类，并按状态差异化展示关键字段。"""
    try:
        report_date = datetime.strptime(report_date_str, "%Y-%m-%d")
    except Exception:
        report_date = None

    lines = [
        f"📊 {title}",
        f"BD：{bd_name} · 报表日 {report_date_str}",
        build_overview(bd_notify, status_order),
        "",
    ]
    total = 0
    for status in status_order:
        items = bd_notify.get(status, [])
        if not items:
            continue
        total += len(items)
        lines.append(f"【{STATUS_DISPLAY.get(status, status)}】{len(items)}个")
        lines.append("—" * 30)
        for item in items:
            uid, name, site, bd, ch_time, first_txn, tpv, _ = item
            lines.append(f"• {name} | {site}")
            if status in UNPAID_STATUS:
                # 未出单：核心是"已开通几天"，不显示无意义的 TPV:$0
                if report_date and ch_time:
                    days = (report_date.date() - ch_time.date()).days
                    tail = f"已开通{days}天未出单"
                else:
                    tail = "未出单"
                lines.append(f"  ID:{uid} | 通道通过:{_fmt_date(ch_time)} | {tail}")
            elif status == APPROVED_WAITING:
                # 刚过审：可上生产、尚无交易
                lines.append(f"  ID:{uid} | 通道通过:{_fmt_date(ch_time)} | 待上量")
            elif status == STUCK_STATUS:
                # 小额滞留：首笔 + TPV + 距首笔滞留天数
                if report_date and first_txn:
                    sd = (report_date.date() - first_txn.date()).days
                    tail = f"停滞{sd}天"
                else:
                    tail = "停滞"
                tpv_str = f"${tpv:,.2f}" if tpv else "$0"
                lines.append(f"  ID:{uid} | 首笔:{_fmt_date(first_txn)} | TPV:{tpv_str} | {tail}")
            else:
                # 首笔出单 / 测试交易：看首笔日期 + TPV
                tpv_str = f"${tpv:,.2f}" if tpv else "$0"
                lines.append(f"  ID:{uid} | 首笔:{_fmt_date(first_txn)} | TPV:{tpv_str}")
        lines.append("")
    if total == 0:
        lines.append("暂无需要跟进的商户。")
    return "\n".join(lines)


# ======================== 发送统一逻辑 ========================
def _send_for_map(token, bd_map, report_date_str, title, status_order):
    if not bd_map:
        print("  📌 无需要发送的内容。")
        return
    for bd_name, bd_notify in bd_map.items():
        if TEST_BD and bd_name != TEST_BD:
            print(f"  📌 测试版跳过BD：{bd_name}")
            continue
        email = BD_EMAILS.get(bd_name, "")
        if not email:
            print(f"  ⚠️ 找不到BD邮箱，跳过：{bd_name}")
            continue
        msg = build_bd_message(bd_name, bd_notify, report_date_str, title, status_order)
        print(f"  发送给：{bd_name}")
        send_private_message(token, email, msg)


def send_bd_daily_messages(token, notify, report_date_str):
    print("\n📨 BD每日私聊...")
    daily_map = split_notify_by_bd(notify, DAILY_STATUS)
    _send_for_map(token, daily_map, report_date_str, "商户出单监控日报", DAILY_STATUS)


def send_bd_weekly_messages(token, notify, report_date_str):
    # 修复点3：按报表业务日判断星期，而非运行当天
    try:
        report_weekday = datetime.strptime(report_date_str, "%Y-%m-%d").weekday()
    except Exception:
        report_weekday = datetime.now().weekday()
    if report_weekday != WEEKLY_SEND_WEEKDAY:
        print("\n📌 报表业务日非周报发送日，跳过长期未出单周报。")
        return
    print("\n📨 BD每周未出单私聊...")
    weekly_map = split_notify_by_bd(notify, WEEKLY_STATUS)
    _send_for_map(token, weekly_map, report_date_str, "商户长期未出单周报", WEEKLY_STATUS)


def send_bd_messages(token, notify, report_date_str):
    send_bd_daily_messages(token, notify, report_date_str)
    send_bd_weekly_messages(token, notify, report_date_str)


# ======================== 处理主流程（与 GUI 解耦，便于测试）========================
def run_pipeline(original, today_file, yesterday_file):
    report_date_str = (
        original.extract_date_from_filename(today_file)
        or datetime.now().strftime("%Y-%m-%d")
    )
    print(f"📅 报表日期: {report_date_str}")

    print("📖 读取报表...")
    df_today = original.read_excel_sheet2(today_file)
    df_yesterday = original.read_excel_sheet2(yesterday_file)
    print(f"  今天 {len(df_today)} 条 / 昨天 {len(df_yesterday)} 条")

    print("\n🔄 分析数据...")
    results, notify = original.classify_from_excel(df_today, df_yesterday, report_date_str)

    print("\n📊 分析结果:")
    for cat in DAILY_STATUS + WEEKLY_STATUS:
        items = notify.get(cat, [])
        if items:
            print(f"  {cat}: {len(items)} 个")

    # ---- 原始总群通知（独立保护）----
    print("\n📨 发送原始总群通知...")
    try:
        original.send_webhook_notification(notify, report_date_str)
    except Exception as e:
        print(f"  ❌ 群通知失败: {e}")

    # ---- 取一次 token：私聊 + 多维表格共用原脚本 App ----
    token = None
    try:
        token = original.get_tenant_access_token()
        print("  ✅ Token 获取成功")
    except Exception as e:
        print(f"  ❌ Token 获取失败，跳过多维表格与BD私聊: {e}")

    # ---- 多维表格（独立保护）----
    if token:
        print("\n🔗 更新飞书多维表格...")
        try:
            original.sync_to_bitable(token, results)
        except Exception as e:
            print(f"  ❌ 多维表格更新失败: {e}")
            import traceback
            traceback.print_exc()

    # ---- BD 私聊（独立保护，失败不影响其它）----
    if token:
        try:
            send_bd_messages(token, notify, report_date_str)
        except Exception as e:
            print(f"  ❌ BD私聊失败: {e}")
            import traceback
            traceback.print_exc()

    # ---- 本地 txt（不依赖 token，始终执行）----
    try:
        original.save_local_report(notify, report_date_str, os.path.dirname(today_file))
    except Exception as e:
        print(f"  ❌ 本地保存失败: {e}")

    print("\n✅ 全部完成！")


# ======================== 主入口（GUI）========================
def main():
    import tkinter as tk
    from tkinter import filedialog, simpledialog, messagebox

    original = load_original_module()

    root = tk.Tk()
    root.withdraw()
    pwd = simpledialog.askstring("密码验证", "请输入密码：", show="*", parent=root)
    if pwd != original.PASSWORD:
        messagebox.showerror("错误", "密码不正确。")
        sys.exit(1)

    messagebox.showinfo("选择文件", "请选择【今天】的商户审核耗时报表")
    today_file = filedialog.askopenfilename(
        title="选择今天的报表", filetypes=[("Excel文件", "*.xls *.xlsx")]
    )
    if not today_file:
        sys.exit(1)

    messagebox.showinfo("选择文件", "请选择【昨天】的商户审核耗时报表")
    yesterday_file = filedialog.askopenfilename(
        title="选择昨天的报表", filetypes=[("Excel文件", "*.xls *.xlsx")]
    )
    if not yesterday_file:
        sys.exit(1)

    root.destroy()

    run_pipeline(original, today_file, yesterday_file)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ 程序出错: {e}")
        import traceback
        traceback.print_exc()
    input("\n按回车键退出...")
