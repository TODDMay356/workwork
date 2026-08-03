# -*- coding: utf-8 -*-
"""
商户出单监控脚本 v12
架构：
  第一步：读两个Excel
  第二步：纯Excel对比，生成分类结果（不涉及多维表格）
  第三步：发飞书群通知
  第四步：更新多维表格（有则更新无则新增，自动过滤不存在的字段）
  第五步：保存本地txt
修复：
  - 用户ID不经过float转换，避免19位大整数精度丢失
  - get_remark_by_days 去掉≤1天档，新审核通过只由"通道反馈=报表当天"判定
  - 飞书通知增加用户ID
"""

import os
import re
import sys
import time
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from datetime import datetime
from decimal import Decimal

import pandas as pd
import requests

# ======================== 配置 ========================
FEISHU_APP_ID = "cli_aa83101e3a38dcd9"
FEISHU_APP_SECRET = ""
BITABLE_APP_TOKEN = "Pk9pbLmmLaS6DOswZnZcS0wYnzh"
BITABLE_TABLE_ID = "tblT1BVy9xmzRL74"
WEBHOOK_URL = "https://open.feishu.cn/open-apis/bot/v2/hook/84dcd2f5-67e8-4427-8448-cb5a0d2ebf91"
PASSWORD = "Landy371407tt"

DATE_FIELDS = [
    "站点_商户提交时间", "站点_风控审核时间", "支付方式_商户提交时间", "支付方式_风控审核时间",
    "提交通道时间", "通道结果反馈时间", "第一笔成功交易时间", "出单日期"
]


# ======================== 工具函数 ========================

def normalize_site(site):
    if isinstance(site, list):
        site = site[0].get("text", "") if site else ""
    if pd.isna(site) or site is None:
        return ""
    site = str(site).strip()
    site = site.replace("https://", "").replace("http://", "")
    if site.startswith("www."):
        site = site[4:]
    site = site.rstrip("/")
    site = site.lower()
    return site


def to_timestamp_ms(val):
    if pd.isna(val) or val is None or str(val).strip() == "":
        return None
    if isinstance(val, (pd.Timestamp, datetime)):
        dt = val.to_pydatetime() if isinstance(val, pd.Timestamp) else val
        return int(dt.timestamp() * 1000)
    for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"]:
        try:
            return int(datetime.strptime(str(val).strip(), fmt).timestamp() * 1000)
        except ValueError:
            continue
    return None


def to_float(val):
    if pd.isna(val) or val is None or str(val).strip() == "":
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def to_date(val):
    if pd.isna(val) or val is None or str(val).strip() == "":
        return None
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.to_pydatetime() if isinstance(val, pd.Timestamp) else val
    for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"]:
        try:
            return datetime.strptime(str(val).strip(), fmt)
        except ValueError:
            continue
    return None


def extract_date_from_filename(filepath):
    match = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(filepath))
    return match.group(1) if match else None


def clean_uid(x):
    """清理用户ID，不经过float，避免大整数精度丢失"""
    if pd.isna(x) or str(x).strip() == "":
        return ""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    if "e" in s.lower():
        # 科学计数法用Decimal精确还原
        try:
            s = str(int(Decimal(s)))
        except Exception:
            pass
    return s


def read_excel_sheet2(filepath):
    converters = {"用户ID": str}
    try:
        df = pd.read_excel(filepath, sheet_name=1, header=1, converters=converters)
    except Exception:
        try:
            df = pd.read_excel(filepath, sheet_name=1, header=0, converters=converters)
        except Exception:
            df = pd.read_excel(filepath, header=0, converters=converters)
    df.columns = [str(c).strip().replace('\n', '').replace('\r', '') for c in df.columns]
    if df.shape[0] > 0 and str(df.iloc[0, 0]).strip() in ["用户ID", "字段说明", ""]:
        df = df.iloc[1:].reset_index(drop=True)
    if "用户ID" in df.columns:
        df["用户ID"] = df["用户ID"].apply(clean_uid)
    return df


def get_remark_by_days(days_open):
    """开通天数对应的备注状态。注意：不含≤1天档（新审核通过单独由通道反馈日期判定）"""
    if days_open <= 3:
        return "👀 3天内未出单"
    elif days_open <= 30:
        return "🔴 3-30天未出单"
    elif days_open <= 60:
        return "🟠 30-60天未出单"
    return None


# ================================================================
# 第二步：纯Excel对比（不涉及多维表格）
# ================================================================

# 小额滞留判定阈值
STUCK_DAYS_THRESHOLD = 3      # 距首笔交易满 N 天，累计仍未上量(<$100)
STUCK_GROWTH_EPS = 0.01       # 较昨日 TPV 增长小于此值，视为"零增长/真停滞"

# 状态显示标签（群通知/本地存档/控制台用，去图标；notify 的 key 保持带图标不变）
STATUS_LABEL = {
    "✅ 新出单": "新出单（今日累计TPV≥$100）",
    "🧪 测试交易": "测试交易（今日开始有成功交易且累计TPV<$100）",
    "🐢 小额滞留": "小额滞留（长期累计TPV<$100）",
    "🆕 新审核通过-待出单": "审核通过-待出单（通道已通过审批）",
    "👀 3天内未出单": "未出单·3天内",
    "🔴 3-30天未出单": "未出单·3-30天",
    "🟠 30-60天未出单": "未出单·30-60天",
}


def classify_from_excel(df_today, df_yesterday, report_date_str):
    """纯Excel对比，返回分类结果。唯一键=用户ID+站点（避免同站点多商户冲突）"""
    report_date = datetime.strptime(report_date_str, "%Y-%m-%d")

    # 昨天数据（key = 用户ID|站点）
    yesterday_map = {}
    for _, row in df_yesterday.iterrows():
        sk = normalize_site(row.get("站点", ""))
        uid_y = clean_uid(row.get("用户ID", ""))
        if sk:
            yesterday_map[f"{uid_y}|{sk}"] = {
                "tpv": to_float(row.get("累计TPV_USD", 0)),
                "first_txn": to_date(row.get("第一笔成功交易时间")),
            }

    results = []
    notify = {}

    for _, row in df_today.iterrows():
        site_raw = str(row.get("站点", "")).strip()
        site_norm = normalize_site(site_raw)
        if not site_norm:
            continue

        uid = clean_uid(row.get("用户ID", ""))
        match_key = f"{uid}|{site_norm}"   # 唯一键

        tpv_today = to_float(row.get("累计TPV_USD", 0))
        first_txn = to_date(row.get("第一笔成功交易时间"))
        site_submit = to_date(row.get("站点_商户提交时间"))
        channel_fb = to_date(row.get("通道结果反馈时间"))
        submit_ch = to_date(row.get("提交通道时间"))

        name = str(row.get("商户名称", "")).strip()
        bd = str(row.get("所属BD", "")).strip()

        yd = yesterday_map.get(match_key, {})
        tpv_yd = yd.get("tpv", 0.0)
        first_txn_yd = yd.get("first_txn")

        # 老商户排除
        if first_txn and site_submit and first_txn < site_submit:
            continue

        # 新出单：今天TPV>=100 且 昨天<100
        if tpv_today >= 100:
            if tpv_yd < 100:
                results.append({"site_key": match_key, "row": row,
                                "action": "new_order", "status": "已出单", "remark": "",
                                "order_date": report_date})
                notify.setdefault("✅ 新出单", []).append(
                    (uid, name, site_raw, bd, channel_fb, first_txn, tpv_today, "✅ 新出单"))
            continue

        # 测试交易 / 小额滞留：TPV<100 但有交易活动
        if (first_txn is not None) or (tpv_today > 0):
            if first_txn_yd is None:
                # 今天才首次产生交易 → 测试交易
                results.append({"site_key": match_key, "row": row,
                                "action": "test_txn", "status": "测试交易", "remark": ""})
                notify.setdefault("🧪 测试交易", []).append(
                    (uid, name, site_raw, bd, channel_fb, first_txn, tpv_today, "🧪 测试交易"))
            else:
                # 已试水但累计仍<100：距首笔满N天 且 较昨日零增长 → 小额滞留
                if first_txn is not None:
                    stuck_days = (report_date.date() - first_txn.date()).days
                    growth = tpv_today - tpv_yd
                    if stuck_days >= STUCK_DAYS_THRESHOLD and abs(growth) < STUCK_GROWTH_EPS:
                        results.append({"site_key": match_key, "row": row,
                                        "action": "stuck", "status": "小额滞留", "remark": ""})
                        notify.setdefault("🐢 小额滞留", []).append(
                            (uid, name, site_raw, bd, channel_fb, first_txn, tpv_today, "🐢 小额滞留"))
            continue

        # 无交易活动 + 通道反馈为空 → 通道审核中/待提交通道
        if channel_fb is None:
            status = "通道审核中" if submit_ch else "待提交通道"
            results.append({"site_key": match_key, "row": row,
                            "action": "set_status", "status": status, "remark": ""})
            continue

        # 通道已通过，开通天数
        days_open = (report_date.date() - channel_fb.date()).days

        # 新审核通过 = 通道反馈日期 == 报表日期
        if channel_fb.date() == report_date.date():
            status = "🆕 新审核通过-待出单"
            results.append({"site_key": match_key, "row": row,
                            "action": "new_approved", "status": status, "remark": ""})
            notify.setdefault(status, []).append(
                (uid, name, site_raw, bd, channel_fb, first_txn, tpv_today, status))
            continue

        # 存量未出单
        if days_open > 60:
            continue
        remark = get_remark_by_days(days_open)
        if remark is None:
            continue
        results.append({"site_key": match_key, "row": row,
                        "action": "update_remark", "status": "未出单", "remark": remark})
        notify.setdefault(remark, []).append(
            (uid, name, site_raw, bd, channel_fb, first_txn, tpv_today, remark))

    return results, notify


# ================================================================
# 第三步：飞书群通知
# ================================================================

def send_webhook_notification(notify, report_date_str):
    order = ["✅ 新出单", "🧪 测试交易", "🐢 小额滞留", "🆕 新审核通过-待出单",
             "👀 3天内未出单", "🔴 3-30天未出单", "🟠 30-60天未出单"]
    lines = ["📊 **商户出单监控日报**", f"📅 报表日期: {report_date_str}",
             f"⏰ 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}", ""]
    total = 0
    for cat in order:
        items = notify.get(cat, [])
        if not items:
            continue
        total += len(items)
        lines.append(f"**{STATUS_LABEL.get(cat, cat)}** ({len(items)}个)")
        lines.append("—" * 30)
        for uid, name, site, bd, ch_time, first_txn, tpv, _ in items:
            ch_str = ch_time.strftime('%Y-%m-%d') if ch_time else "-"
            f_str = first_txn.strftime('%Y-%m-%d') if first_txn else "-"
            tpv_str = f"${tpv:,.2f}" if tpv else "$0"
            lines.append(f"• {name} | {site}")
            lines.append(f"  ID:{uid} | BD:{bd} | 通道通过:{ch_str} | 首笔:{f_str} | TPV:{tpv_str}")
        lines.append("")
    if total == 0:
        lines.append("今日无需关注的商户变更。")
    payload = {"msg_type": "text", "content": {"text": "\n".join(lines)}}
    try:
        resp = requests.post(WEBHOOK_URL, json=payload, timeout=30)
        data = resp.json()
        if data.get("code") == 0 or data.get("StatusCode") == 0:
            print("✅ 飞书群通知发送成功")
        else:
            print(f"⚠️ 飞书群通知发送异常: {data}")
    except Exception as e:
        print(f"❌ 飞书群通知发送失败: {e}")


# ================================================================
# 第四步：更新多维表格（分析完成后才执行）
# ================================================================

def get_tenant_access_token():
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    resp = requests.post(url, json={"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET}, timeout=30)
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"获取token失败: {data}")
    return data["tenant_access_token"]


def get_bitable_field_names(token):
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/fields"
    headers = {"Authorization": f"Bearer {token}"}
    names = set()
    page_token = None
    while True:
        params = {"page_size": 100}
        if page_token:
            params["page_token"] = page_token

        data = None
        for attempt in range(3):
            try:
                resp = requests.get(url, headers=headers, params=params, timeout=15)
                data = resp.json()
                break
            except Exception:
                print(f"  ⚠️ 读取字段超时，重试 {attempt+1}/3...")
                time.sleep(2)
        if data is None:
            print("  ❌ 读取字段重试失败")
            break
        if data.get("code") != 0:
            print(f"  ⚠️ 读取字段失败: {data.get('msg', data)}")
            break

        for f in data.get("data", {}).get("items", []) or []:
            names.add(f.get("field_name"))

        has_more = data.get("data", {}).get("has_more", False)
        page_token = data.get("data", {}).get("page_token")
        if not has_more or not page_token:
            break
    return names

def list_all_bitable_records(token):
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/records"
    headers = {"Authorization": f"Bearer {token}"}
    site_map = {}
    page_token = None
    while True:
        params = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token

        data = None
        for attempt in range(3):
            try:
                resp = requests.get(url, headers=headers, params=params, timeout=15)
                data = resp.json()
                break
            except Exception:
                print(f"  ⚠️ 读取记录超时，重试 {attempt+1}/3...")
                time.sleep(2)
        if data is None:
            print("  ❌ 读取记录重试失败")
            break
        if data.get("code") != 0:
            print(f"  ⚠️ 读取记录失败: {data.get('msg', data)}")
            break

        for item in data.get("data", {}).get("items", []) or []:
            fields = item.get("fields", {})
            sk = normalize_site(fields.get("站点", ""))
            uid_b = clean_uid(fields.get("用户ID", ""))
            if sk:
                site_map[f"{uid_b}|{sk}"] = {"record_id": item.get("record_id"), "fields": fields}

        has_more = data.get("data", {}).get("has_more", False)
        page_token = data.get("data", {}).get("page_token")
        if not has_more or not page_token:
            break
        time.sleep(0.2)
    return site_map


def filter_fields(fields, valid):
    return {k: v for k, v in fields.items() if k in valid}


NUMBER_FIELDS = [
    "站点_上线时长", "支付方式_上线时长", "站点_网站审核耗时", "支付方式_网站审核耗时",
    "通道审核耗时", "集成耗时", "累计TPV_USD"
]


def build_full_fields(row, status, remark, order_date=None):
    fields = {}
    for col in row.index:
        if col == "用户ID":
            fields[col] = clean_uid(row[col])
        elif col in NUMBER_FIELDS:
            fields[col] = to_float(row[col])
        elif col in DATE_FIELDS:
            ts = to_timestamp_ms(row[col])
            if ts is not None:
                fields[col] = ts
        else:
            val = row[col]
            fields[col] = "" if (pd.isna(val) or val is None) else str(val)
    fields["是否出单"] = status
    fields["备注"] = remark
    if order_date:
        fields["出单日期"] = to_timestamp_ms(order_date)
    return fields


def batch_create(token, records, valid):
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/records/batch_create"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    created = 0
    for i in range(0, len(records), 500):
        batch = records[i:i+500]
        payload = {"records": [{"fields": filter_fields(r, valid)} for r in batch]}
        resp = requests.post(url, headers=headers, json=payload, timeout=60)
        data = resp.json()
        if data.get("code") != 0:
            print(f"  ⚠️ 新增失败: {data.get('msg', data)}")
        else:
            created += len(batch)
        time.sleep(0.3)
    return created


def batch_update(token, updates, valid):
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/records/batch_update"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    updated = 0
    for i in range(0, len(updates), 500):
        batch = updates[i:i+500]
        payload = {"records": [{"record_id": rid, "fields": filter_fields(f, valid)} for rid, f in batch]}
        resp = requests.post(url, headers=headers, json=payload, timeout=60)
        data = resp.json()
        if data.get("code") != 0:
            print(f"  ⚠️ 更新失败: {data.get('msg', data)}")
        else:
            updated += len(batch)
        time.sleep(0.3)
    return updated


def sync_to_bitable(token, results):
    """把分类结果同步到多维表格。有则更新，无则新增。"""
    print("\n📋 读取多维表格字段名...")
    valid = get_bitable_field_names(token)
    print(f"  字段数: {len(valid)}")

    # 字段对齐校验：把报表里存在、但多维表格没有的关键字段打印出来，
    # 避免 filter_fields 静默丢弃导致“看起来成功但没写进去”。
    key_cols = set(DATE_FIELDS + NUMBER_FIELDS + ["用户ID", "站点", "商户名称", "所属BD", "是否出单", "备注"])
    missing = [c for c in key_cols if c not in valid]
    if missing:
        print(f"  ⚠️ 以下关键字段在多维表格中不存在，将被跳过不写入：{missing}")
        print("     （请核对多维表格字段名是否与报表列名完全一致）")

    print("📖 读取多维表格已有记录...")
    bitable_map = list_all_bitable_records(token)
    print(f"  已有: {len(bitable_map)} 条")

    creates = []
    updates = []

    for r in results:
        site_key = r["site_key"]
        row = r["row"]
        action = r["action"]
        status = r["status"]
        remark = r["remark"]
        existing = bitable_map.get(site_key)

        # 统一策略：整行覆盖。
        # 每个 action 只决定三个计算字段（是否出单/备注/出单日期），
        # 其余字段一律用报表整行的最新值，existing 与 create 走同一套 build_full_fields。
        # 这样多维表格永远 = 报表现状，不会出现字段僵死。
        if action == "new_order":
            db_status, db_remark, order_date = "已出单", "", r.get("order_date")
        elif action == "test_txn":
            db_status, db_remark, order_date = "测试交易", "", None
        elif action == "new_approved":
            db_status, db_remark, order_date = status, remark, None
        elif action == "set_status":
            db_status, db_remark, order_date = status, "", None
        elif action == "update_remark":
            db_status, db_remark, order_date = "未出单", remark, None
        else:
            continue

        full = build_full_fields(row, db_status, db_remark, order_date)
        if existing:
            updates.append((existing["record_id"], full))
        else:
            creates.append(full)

    if creates:
        print(f"\n📝 新增 {len(creates)} 条...")
        n = batch_create(token, creates, valid)
        print(f"  ✅ 成功 {n} 条")
    else:
        print("\n📝 无新增")

    if updates:
        print(f"📝 更新 {len(updates)} 条...")
        n = batch_update(token, updates, valid)
        print(f"  ✅ 成功 {n} 条")
    else:
        print("📝 无更新")


# ================================================================
# 第五步：保存本地txt
# ================================================================

def save_local_report(notify, report_date_str, output_dir):
    output_file = os.path.join(output_dir, f"出单监控结果_{report_date_str}.txt")
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(f"商户出单监控结果 - {report_date_str}\n")
        f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("=" * 60 + "\n\n")
        order = ["✅ 新出单", "🧪 测试交易", "🐢 小额滞留", "🆕 新审核通过-待出单",
                 "👀 3天内未出单", "🔴 3-30天未出单", "🟠 30-60天未出单"]
        for cat in order:
            items = notify.get(cat, [])
            if not items:
                continue
            f.write(f"\n【{STATUS_LABEL.get(cat, cat)}】({len(items)}个)\n")
            f.write("-" * 40 + "\n")
            for uid, name, site, bd, ch_time, first_txn, tpv, _ in items:
                ch_str = ch_time.strftime('%Y-%m-%d') if ch_time else "-"
                f_str = first_txn.strftime('%Y-%m-%d') if first_txn else "-"
                f.write(f"  ID:{uid} | {name} | {site} | BD:{bd}\n")
                f.write(f"    通道通过:{ch_str} | 首笔:{f_str} | TPV:${tpv:,.2f}\n")
    print(f"💾 结果已保存: {output_file}")


# ================================================================
# 主入口
# ================================================================

def main():
    # ---- 密码 ----
    root = tk.Tk()
    root.withdraw()
    pwd = simpledialog.askstring("密码验证", "请输入密码：", show="*", parent=root)
    if pwd != PASSWORD:
        messagebox.showerror("错误", "密码不正确。")
        sys.exit(1)

    # ---- 选文件 ----
    messagebox.showinfo("选择文件", "请选择【今天】的商户审核耗时报表")
    today_file = filedialog.askopenfilename(title="选择今天的报表", filetypes=[("Excel文件", "*.xls *.xlsx")])
    if not today_file:
        sys.exit(1)
    messagebox.showinfo("选择文件", "请选择【昨天】的商户审核耗时报表")
    yesterday_file = filedialog.askopenfilename(title="选择昨天的报表", filetypes=[("Excel文件", "*.xls *.xlsx")])
    if not yesterday_file:
        sys.exit(1)

    # ---- 报表日期：文件名先猜一个默认值，弹窗让用户确认/修改 ----
    default_date = extract_date_from_filename(today_file) or datetime.now().strftime("%Y-%m-%d")
    report_date_str = None
    while report_date_str is None:
        entered = simpledialog.askstring(
            "确认报表日期",
            "请确认本次报表的业务日期（格式 YYYY-MM-DD）：\n\n"
            "该日期用于判定“当天过审=新审核通过”及开通天数，请务必与实际业务日一致。",
            initialvalue=default_date, parent=root,
        )
        if entered is None:  # 用户取消
            sys.exit(1)
        entered = entered.strip()
        try:
            datetime.strptime(entered, "%Y-%m-%d")
            report_date_str = entered
        except ValueError:
            messagebox.showerror("格式错误", f"“{entered}”不是有效日期，请按 YYYY-MM-DD 重新输入。")

    root.destroy()

    # ---- 第一步：读Excel ----
    print(f"📅 报表日期: {report_date_str}")

    print("📖 读取今天报表...")
    df_today = read_excel_sheet2(today_file)
    print(f"  {len(df_today)} 条")

    print("📖 读取昨天报表...")
    df_yesterday = read_excel_sheet2(yesterday_file)
    print(f"  {len(df_yesterday)} 条")

    # ---- 第二步：纯Excel对比分类 ----
    print("\n🔄 分析数据（纯Excel对比）...")
    results, notify = classify_from_excel(df_today, df_yesterday, report_date_str)

    print(f"\n📊 分析结果:")
    for cat in ["✅ 新出单", "🧪 测试交易", "🐢 小额滞留", "🆕 新审核通过-待出单",
                "👀 3天内未出单", "🔴 3-30天未出单", "🟠 30-60天未出单"]:
        items = notify.get(cat, [])
        if items:
            print(f"  {STATUS_LABEL.get(cat, cat)}: {len(items)} 个")

    new_approved = notify.get("🆕 新审核通过-待出单", [])
    if new_approved:
        print(f"\n🆕 新审核通过明细（通道反馈={report_date_str}）:")
        for _, name, site, _, ch_time, _, _, _ in new_approved:
            print(f"  {name} | {site} | 通道反馈: {ch_time.strftime('%Y-%m-%d %H:%M') if ch_time else '-'}")

    # ---- 第三步：发飞书群通知 ----
    print("\n📨 发送飞书群通知...")
    send_webhook_notification(notify, report_date_str)

    # ---- 第四步：更新多维表格 ----
    print("\n🔗 连接飞书更新多维表格...")
    try:
        token = get_tenant_access_token()
        print("  ✅ Token获取成功")
        sync_to_bitable(token, results)
    except Exception as e:
        print(f"  ❌ 多维表格更新失败: {e}")
        import traceback
        traceback.print_exc()

    # ---- 第五步：保存本地txt ----
    save_local_report(notify, report_date_str, os.path.dirname(today_file))

    print("\n✅ 全部完成！")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ 程序出错: {e}")
        import traceback
        traceback.print_exc()
    input("\n按回车键退出...")
