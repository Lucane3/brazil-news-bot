# 🇧🇷 巴西资讯日报自动推送系统

> **无需开机 · GitHub Actions云端定时执行 · 每天自动发送到QQ邮箱**

---

## 🚀 快速部署（3步完成）

### 第1步：获取QQ邮箱授权码

打开 QQ邮箱网页版 → **设置** → **账户** → 找到 **POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV服务**

↓ 开启 **SMTP服务**，按提示发送短信验证后会得到一个 **16位授权码**，复制保存。

> ⚠️ **这是授权码，不是QQ密码！** 形如：`abcdefghijklmnop`

---

### 第2步：推送到GitHub并设置密钥

```bash
# 在项目目录打开终端，执行以下命令：

# 1. 初始化Git仓库
git init
git add .
git commit -m "初始化巴西资讯日报系统"

# 2. 在GitHub上新建仓库（如 brazil-news-bot），然后关联推送
git remote add origin https://github.com/你的用户名/brazil-news-bot.git
git branch -M main
git push -u origin main
```

推送成功后，打开GitHub仓库页面：

→ **Settings** → **Secrets and variables** → **Actions** → **New repository secret**

| Name | Secret |
|------|--------|
| `QQ_MAIL_AUTH_CODE` | 粘贴第1步获取的16位授权码 |

---

### 第3步：启用工作流

推送代码后，GitHub Actions 会自动检测 `.github/workflows/daily.yml`。

进入仓库 **Actions** 标签页，点击 **"I understand my workflows, go ahead and enable them"** 启用。

可以点击 **"Run workflow"** 手动触发一次测试。

---

## ⏰ 执行时间

- **北京时间每天上午 9:00** 自动执行（UTC 1:00）
- GitHub Actions 免费额度：每月2000分钟，本任务每次约1分钟，**完全够用**
- 如需修改时间，编辑 `.github/workflows/daily.yml` 中的 `cron` 行

---

## 📡 信息源覆盖（7个源）

| 来源 | 类型 | 覆盖领域 |
|------|------|----------|
| Google News Brasil | RSS | 巴西财经、贸易综合 |
| BCB Focus Report | 网页 | 巴西央行货币政策、Selic利率 |
| IBGE Notícias | RSS | 巴西宏观经济指标（GDP、通胀等） |
| Reuters Brasil | RSS | 国际视角巴西财经市场 |
| Xinhua Português | RSS | 中国视角巴西政经 |
| 南美侨报网 | RSS | 华人视角跨境贸易 |
| Receita Federal | 网页 | 巴西税务政策变动 |

---

## 📧 邮件效果预览

邮件包含：
- 📰 当日巴西相关资讯汇总（自动去重、分类）
- 📊 信息源统计（覆盖几个源、多少条）
- 🔗 每条资讯可点击跳转原文
- 🎨 巴西国旗配色HTML格式

---

## 🛠 本地测试（可选）

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 设置授权码并测试邮件
set QQ_MAIL_AUTH_CODE=你的授权码
python test_email.py

# 3. 完整运行（采集+发送）
python main.py
```

---

## 📋 常见问题

| 问题 | 解决 |
|------|------|
| 收不到邮件 | 检查QQ邮箱垃圾箱；确认授权码正确、SMTP服务已开启 |
| GitHub Actions执行失败 | 查看Actions日志；确认Secret已正确设置、拼写为`QQ_MAIL_AUTH_CODE` |
| 资讯太少 | 个别源可能因网络波动抓取失败属正常现象，次日会自动重试 |
| 想换邮箱 | 修改`main.py`顶部`RECEIVER_EMAIL`变量即可 |

---

## 🔒 安全说明

- QQ邮箱授权码存储在GitHub Secrets中，加密保护，不会泄露
- 本程序仅读取公开新闻资源，不涉及任何私人数据
- 所有代码开源透明，可自行审计
