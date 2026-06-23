#!/usr/bin/env python3
"""
本地测试脚本：验证邮件发送功能是否正常
运行前设置环境变量：set QQ_MAIL_AUTH_CODE=你的QQ邮箱授权码
"""
import os
import smtplib
import ssl
from email.header import Header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from datetime import datetime, timezone, timedelta

SMTP_HOST = "smtp.qq.com"
SMTP_PORT = 465
SENDER = "847988716@qq.com"
RECEIVER = "847988716@qq.com"
AUTH_CODE = os.environ.get("QQ_MAIL_AUTH_CODE", "")

def test_email():
    if not AUTH_CODE:
        print("❌ 未设置 QQ_MAIL_AUTH_CODE 环境变量")
        print("   请先设置：set QQ_MAIL_AUTH_CODE=你的授权码")
        print("   获取方式：QQ邮箱 → 设置 → 账户 → POP3/SMTP → 生成授权码")
        return False
    
    bj_now = datetime.now(timezone(timedelta(hours=8)))
    date_str = bj_now.strftime("%Y年%m月%d日 %H:%M")
    
    html = f"""<html><body style="font-family:'Microsoft YaHei',sans-serif;">
<h2>✅ 巴西资讯机器人 · 测试邮件</h2>
<p>发送时间：{date_str}</p>
<p>如果你收到这封邮件，说明 <b>QQ邮箱SMTP配置成功</b>！</p>
<hr>
<p style="color:#999;">接下来只需推送到GitHub即可实现每日自动推送 🇧🇷</p>
</body></html>"""
    
    msg = MIMEMultipart("alternative")
    msg["From"] = SENDER
    msg["To"] = RECEIVER
    msg["Subject"] = Header(f"巴西资讯日报测试 | {date_str}", "utf-8")
    msg.attach(MIMEText("巴西资讯日报测试邮件", "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))
    
    try:
        ctx = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ctx) as server:
            server.login(SENDER, AUTH_CODE)
            server.sendmail(SENDER, RECEIVER, msg.as_string())
        print(f"✅ 测试邮件发送成功！请检查 {RECEIVER} 收件箱")
        return True
    except smtplib.SMTPAuthenticationError:
        print("❌ 认证失败，请确认授权码正确（不是QQ密码！）")
        return False
    except Exception as e:
        print(f"❌ 发送失败：{e}")
        return False

if __name__ == "__main__":
    test_email()
