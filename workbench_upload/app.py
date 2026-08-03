"""
app.py - 工作台 Flask 主程序

启动：python app.py
默认监听：http://127.0.0.1:5050

路由：
  GET  /                         工作台首页
  GET  /conversion               每日转化率工具（static/conversion.html）
  GET  /merchant                 单商户成功率工具（static/merchant.html）
  GET  /api/health               探活
  GET  /api/status               配置总览（首页用）
  POST /api/feishu/webhook       转发 webhook 消息到所有已配置机器人
  POST /api/feishu/bitable       写入多维表格（自动加 Bearer token）
  POST /api/feishu/notify-me     (可选) App 身份发 DM
"""
from __future__ import annotations
import os
import sys
import json
import logging
import subprocess
from pathlib import Path

from flask import (
    Flask, render_template, send_from_directory, jsonify, request, abort
)

import feishu


HERE = Path(__file__).resolve().parent
CONFIG_PATH = HERE / "config.json"
EXAMPLE_CONFIG_PATH = HERE / "config.example.json"
STATIC_DIR = HERE / "static"


# ---------- 配置加载 ----------

def _load_config() -> dict:
    if not CONFIG_PATH.exists():
        # 没填配置也不致命，首页会提示"未配置"
        if EXAMPLE_CONFIG_PATH.exists():
            try:
                return json.loads(EXAMPLE_CONFIG_PATH.read_text(encoding="utf-8"))
            except Exception:
                pass
        return {}
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        print(f"[config] 解析 config.json 失败：{e}", file=sys.stderr)
        return {}


config = _load_config()
server_cfg = (config.get("server") or {})
HOST = server_cfg.get("host", "127.0.0.1")
PORT = int(server_cfg.get("port", 5050))
DEBUG = bool(server_cfg.get("debug", False))


# ---------- 日志 ----------

logging.basicConfig(
    level=logging.DEBUG if DEBUG else logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("workbench")


# ---------- Flask app ----------

app = Flask(__name__, template_folder="templates", static_folder="static")


@app.after_request
def _cors(resp):
    """本机同源调用为主，这里加宽松 CORS 是给将来部署留口子（不影响本地）。"""
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    return resp


@app.route("/options", methods=["OPTIONS"])
def _preflight():
    return ("", 204)


# ---------- 工作台首页 ----------

@app.route("/")
def index():
    summary = {}
    if config:
        try:
            summary = feishu.get().config_summary()
        except Exception:
            # feishu 客户端未初始化（配置缺失）也不让首页崩
            summary = {"feishu_app_configured": False}
    return render_template("index.html", summary=summary, has_config=bool(config))


# ---------- 工具页 ----------

@app.route("/conversion")
def conversion():
    if not (STATIC_DIR / "conversion.html").exists():
        abort(404, description="static/conversion.html 不存在")
    return send_from_directory(STATIC_DIR, "conversion.html")


@app.route("/merchant")
def merchant():
    if not (STATIC_DIR / "merchant.html").exists():
        abort(404, description="static/merchant.html 不存在")
    return send_from_directory(STATIC_DIR, "merchant.html")


# ---------- API ----------

@app.route("/api/health")
def health():
    return jsonify({
        "ok": True,
        "service": "workbench",
        "feishu_ready": bool(config.get("feishu", {}).get("app_id")),
    })


@app.route("/api/status")
def status():
    if not config:
        return jsonify({
            "config_loaded": False,
            "hint": "未找到 config.json。请复制 config.example.json 为 config.json 并填写。",
        })
    try:
        s = feishu.get().config_summary()
    except Exception as e:
        return jsonify({"config_loaded": True, "error": str(e)})
    return jsonify({"config_loaded": True, "summary": s})


@app.route("/api/feishu/webhook", methods=["POST"])
def relay_webhook():
    """
    接收 HTML 发的消息体（msg_type / content / 可选 card）。
    自动签名后转发到配置里所有已配置的 webhook（团队 + 个人）。
    """
    body = request.get_json(silent=True) or {}
    if not body:
        abort(400, description="body 为空或不是 JSON")
    if "msg_type" not in body:
        abort(400, description="缺少 msg_type")
    try:
        client = feishu.get()
    except Exception:
        abort(503, description="飞书客户端未初始化（请先填 config.json）")
    results = client.broadcast_webhook(body)
    if not results:
        return jsonify({
            "ok": False,
            "msg": "config.json 中没有配置任何 webhook（feishu.webhook.team / personal）",
        }), 400
    # 聚合：只要有一个成功就算 ok
    any_ok = any(r.get("status", 0) == 200 and (r.get("data") or {}).get("code") in (0, "0")
                 for r in results.values())
    return jsonify({"ok": any_ok, "results": results})


@app.route("/api/feishu/bitable", methods=["POST"])
def relay_bitable():
    """
    接收 HTML 构造的 {records: [...]} 写入多维表格。
    自动用工作台持有的 tenant_access_token 加 Bearer 头。
    """
    body = request.get_json(silent=True) or {}
    records = body.get("records")
    if not isinstance(records, list) or not records:
        abort(400, description="缺少 records 数组")
    # HTML 发的格式是 [{fields: {...}}, ...]；如果只发了 fields 列表也兼容
    if isinstance(records[0], dict) and "fields" not in records[0] and records[0]:
        records = [{"fields": r} for r in records]
    try:
        client = feishu.get()
    except Exception:
        abort(503, description="飞书客户端未初始化（请先填 config.json）")
    res = client.bitable_batch_create(records)
    return jsonify(res), (200 if res.get("ok") else 502)


@app.route("/api/feishu/open-apis/bitable/v1/apps/<app_token>/tables/<table_id>/records/batch_create", methods=["POST"])
def relay_bitable_feishu_style(app_token, table_id):
    """
    兼容 HTML 端 feishuUrl() 构造的飞书原生 URL 格式。
    工作台忽略 URL 里的 app_token/table_id，统一用 config.json 里的配置。
    """
    body = request.get_json(silent=True) or {}
    records = body.get("records")
    if not isinstance(records, list) or not records:
        abort(400, description="缺少 records 数组")
    if isinstance(records[0], dict) and "fields" not in records[0] and records[0]:
        records = [{"fields": r} for r in records]
    try:
        client = feishu.get()
    except Exception:
        abort(503, description="飞书客户端未初始化")
    res = client.bitable_batch_create(records)
    return jsonify(res), (200 if res.get("ok") else 502)


@app.route("/api/feishu/notify-me", methods=["POST"])
def notify_me():
    """
    可选：用 App 身份给指定 open_id 发私信。
    Body: {"receive_id": "ou_xxx", "content": {"text": "..."} | md, "msg_type": "text"|"interactive"}
    需要：App 有 im:message:send_as_bot 权限。
    """
    body = request.get_json(silent=True) or {}
    receive_id = body.get("receive_id") or (config.get("feishu") or {}).get("my_open_id")
    if not receive_id:
        abort(400, description="缺少 receive_id（可填在 config.json 的 feishu.my_open_id）")
    try:
        client = feishu.get()
    except Exception:
        abort(503, description="飞书客户端未初始化")
    res = client.send_dm(
        receive_id=receive_id,
        content=body.get("content") or {"text": body.get("text", "")},
        msg_type=body.get("msg_type", "text"),
        id_type=body.get("id_type", "open_id"),
    )
    return jsonify(res), (200 if res.get("ok") else 502)


# ---------- 外部脚本触发 ----------

ORDER_MONITOR_DIR = Path(r"D:\商户数据\出单耗时监控\出单监控")
ORDER_MONITOR_WRAPPER = ORDER_MONITOR_DIR / "出单监控_workbench.py"


@app.route("/api/run/order-monitor", methods=["POST"])
def run_order_monitor():
    today = request.files.get("today")
    yesterday = request.files.get("yesterday")
    password = request.form.get("password", "")
    report_date = request.form.get("date", "")

    if not today or not yesterday:
        abort(400, description="请上传今天和昨天的 Excel 报表")
    if not password:
        abort(400, description="请输入密码")
    if not report_date:
        abort(400, description="请确认报表业务日期")

    today_path = ORDER_MONITOR_DIR / f"_wb_today_{report_date}.xlsx"
    yesterday_path = ORDER_MONITOR_DIR / f"_wb_yesterday_{report_date}.xlsx"
    today.save(str(today_path))
    yesterday.save(str(yesterday_path))

    if not ORDER_MONITOR_WRAPPER.exists():
        abort(500, description=f"wrapper 不存在：{ORDER_MONITOR_WRAPPER}")

    # 前台启动，CREATE_NEW_CONSOLE 打开独立 CMD 窗口
    subprocess.Popen(
        [sys.executable, str(ORDER_MONITOR_WRAPPER),
         "--today", str(today_path),
         "--yesterday", str(yesterday_path),
         "--password", password,
         "--date", report_date],
        cwd=str(ORDER_MONITOR_DIR),
        creationflags=subprocess.CREATE_NEW_CONSOLE,
    )

    return jsonify({"ok": True, "msg": "出单监控窗口已打开"})


# ---------- 启动 ----------

def main():
    if not CONFIG_PATH.exists():
        log.warning("未找到 config.json —— 工作台会以空配置启动，所有飞书接口会返回 400/503。")
        log.warning("请复制 config.example.json 为 config.json 并填写。")
    feishu.init(config, logger=log.info)
    log.info(f"工作台启动: http://{HOST}:{PORT}")
    log.info(f"首页:        http://{HOST}:{PORT}/")
    log.info(f"转化率工具:  http://{HOST}:{PORT}/conversion")
    log.info(f"成功率工具:  http://{HOST}:{PORT}/merchant")
    log.info("Ctrl+C 停止")
    try:
        app.run(host=HOST, port=PORT, debug=DEBUG, use_reloader=False, threaded=True)
    finally:
        try:
            feishu.get().stop()
        except Exception:
            pass


if __name__ == "__main__":
    main()
