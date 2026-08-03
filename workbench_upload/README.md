# 跨境支付运营工作台

本机运行的 Flask 小服务，解决"HTML 网页 + 飞书"两件事：

1. **CORS 拦截** —— 浏览器直接 `fetch('https://open.feishu.cn/...')` 会被拦，
   所有飞书调用都经本服务中转。
2. **多工具入口** —— 当前已接入"每日转化率数据监控 v3"，
   留好"单商户成功率"工具的位置。
3. **Token 自动续期** —— 你只填一次 AppID/AppSecret，
   工作台后台每 1.5 小时自动换新 token，不用管 2 小时过期。

---

## 一次性安装

双击 `install.bat`：

- 创建 `venv\` 虚拟环境
- 安装 Flask + requests
- 复制 `config.example.json` 为 `config.json`

---

## 填写飞书凭据

用记事本打开 `config.json`，按下面顺序填：

### 1. AppID / AppSecret

在 https://open.feishu.cn/app 创建**企业自建应用**，
记录 AppID 和 AppSecret。

App 需要勾选以下权限：
- `im:message` （发消息）
- `im:message:send_as_bot` （以 Bot 身份发）
- `bitable:app` 或具体的 `bitable:app:readonly` / `bitable:app:write`

> 如果你**不要** App 身份发 DM，可以省略 im 权限。

### 2. 团队群 Webhook

1. 打开团队的飞书群 → 设置 → 群机器人
2. 添加"自定义机器人" → 勾选"签名校验" → 记下 webhook URL 和 secret
3. 填到 `feishu.webhook.team.url` 和 `feishu.webhook.team.secret`

### 3. 个人告警 Webhook

建一个**只有你 + 机器人**的飞书群（或自建单聊 + 机器人），
按上面同样方式拿到 webhook 填到 `feishu.webhook.personal`。

> 这就是"推给你个人"的实现 —— 机器人发的消息会出现在你的"个人告警"群里，
> 不需要 App 权限，最简单。

### 4. 多维表格

建一个飞书多维表格，URL 长这样：
```
https://xxx.feishu.cn/base/bascnXXXXXXXXXX?table=tblYYYYYYYYYY
```
- `bascnXXXXXXXX` 填到 `bitable.app_token`
- `tblYYYYYYYY` 填到 `bitable.table_id`
- App 需被加为该表格的协作者（应用 → 添加协作者）

字段建好这些（列名一字不差）：
`日期 / 周期 / 来源 / 指标 / 所处环节 / 层级 / 本期值 / 上期值 / 环比 / 是否异常 / 角色`

---

## 启动

双击 `start.bat`，浏览器会自动打开 http://127.0.0.1:5050

工作台首页会显示：
- 飞书配置状态（App / Token / Webhook / 多维表格）
- 两个工具入口卡片

---

## 使用

1. 打开工作台首页
2. 确认"飞书配置状态"全绿
3. 点"一键配置并打开" → 浏览器新标签打开"每日转化率"工具
4. 上传 Excel 文件 → 工具自动解析
5. **解析完成后自动同步**（已勾选）→ 写多维表格 + 推送告警

如果需要重新配置：
- 重新编辑 `config.json` → 双击 `start.bat` 重启

---

## 目录结构

```
workbench/
├── app.py                  Flask 主程序
├── feishu.py               飞书客户端（token + 中转）
├── config.example.json     配置模板（你复制为 config.json）
├── config.json             你的真实配置（gitignore）
├── templates/
│   └── index.html          工作台首页
├── static/
│   ├── conversion.html     每日转化率工具（直接放这）
│   └── merchant.html       单商户成功率工具（占位）
├── requirements.txt
├── start.bat               Windows 启动
├── install.bat             Windows 一次性安装
└── README.md               本文件
```

---

## 常见问题

**Q: 工作台关了浏览器，工具还能用吗？**
A: 飞书同步相关功能不可用（因为依赖工作台中转），但工具本身的分析 / 图表 / 导出 Excel
都不受影响。HTML 端有"复制 curl"按钮可以临时手动推。

**Q: 想加第三个工具？**
A: 把 HTML 复制到 `static/yourtool.html`，修改 `app.py` 加一行：
```python
@app.route('/yourtool')
def yourtool():
    return send_from_directory(STATIC_DIR, 'yourtool.html')
```
然后在 `templates/index.html` 复制一个 tool-card 改成你的入口。

**Q: token 显示"未生效"？**
A: 检查 `app_id` / `app_secret` 是否正确（注意不要带空格），
或 App 是否在飞书开放平台被禁用。

**Q: 怎么确认 webhook 发成功了？**
A: 在工作台首页按"刷新状态"，再点 `/api/health` 看返回 JSON。
手动测试：在终端执行：
```bash
curl -X POST http://127.0.0.1:5050/api/feishu/webhook ^
  -H "Content-Type: application/json" ^
  -d "{\"msg_type\":\"text\",\"content\":{\"text\":\"hello from workbench\"}}"
```
去飞书群看是否收到消息。

**Q: 能在多台电脑用吗？**
A: 当前版本是本机服务。如果要在多台用，把整个 workbench 目录拷过去，
  或者部署到云服务器（修改 `server.host` 为 `0.0.0.0` 即可）。
