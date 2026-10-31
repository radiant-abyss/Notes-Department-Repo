"""Independent SMTP alerts; no QQ messages or credentials included in email bodies."""
from __future__ import annotations
import asyncio
import json
import re
import smtplib
import ssl
import time
from email.message import EmailMessage
from pathlib import Path


def valid_email(value):
    return bool(isinstance(value, str) and len(value) <= 254 and re.fullmatch(r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9.-]*[A-Za-z0-9])?\.[A-Za-z]{2,}", value))


def read_json(path):
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except FileNotFoundError:
        return {}


def write_json(path, value):
    temp = path.with_suffix(path.suffix + '.tmp')
    temp.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')
    temp.replace(path)


class EmailAlerts:
    def __init__(self, root):
        self.root = Path(root)
        self.settings_path = self.root / 'email_alert.json'
        self.state_path = self.root / 'email_alert_state.json'
        try:
            self.state = read_json(self.state_path)
        except (OSError, ValueError):
            self.state = {}
        self.lock = asyncio.Lock()
        self.last_test = None

    def settings(self):
        return read_json(self.settings_path)

    def configure_recipient(self, recipient):
        if not valid_email(recipient):
            raise ValueError('邮箱格式不正确。例：配置邮箱 name@example.com')
        config = self.settings()
        config['recipient'] = recipient
        write_json(self.settings_path, config)

    def smtp_settings(self):
        config = self.settings()
        path = Path(config.get('smtp_source', ''))
        if not path.is_file():
            raise ValueError('SMTP source missing')
        # Read only the known mail section; never import or run the watcher.
        mail = read_json(path).get('notify', {}).get('mail', {})
        if not mail.get('user') or not mail.get('code') or not mail.get('smtp'):
            raise ValueError('SMTP credentials missing')
        if not valid_email(config.get('recipient')) or not valid_email(mail['user']):
            raise ValueError('Invalid mail address')
        return config, mail

    def status(self):
        try:
            config, _ = self.smtp_settings()
            return ('邮件告警已配置' if config.get('enabled') else '邮件告警未启用') + '\n收件邮箱：' + config['recipient'] + '\n发送“测试邮件”检查送达。'
        except Exception:
            return '邮件告警尚未就绪，请在本机检查 email_alert.json 和发件服务。授权码不要通过 QQ 发送。'

    def send(self, subject, body):
        config, mail = self.smtp_settings()
        message = EmailMessage()
        message['From'] = mail['user']
        message['To'] = config['recipient']
        message['Subject'] = subject
        message.set_content(body)
        port = int(mail.get('port', 465))
        context = ssl.create_default_context()
        if port == 465:
            server = smtplib.SMTP_SSL(mail['smtp'], port, timeout=20, context=context)
        else:
            server = smtplib.SMTP(mail['smtp'], port, timeout=20)
        with server:
            if port != 465:
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()
            server.login(mail['user'], mail['code'])
            refused = server.send_message(message)
            if refused:
                raise RuntimeError('Recipient refused')

    async def test(self):
        async with self.lock:
            now = time.monotonic()
            if self.last_test is not None and now - self.last_test < 60:
                return '测试邮件每分钟最多一次，请稍后再试。'
            self.last_test = now
            try:
                await asyncio.to_thread(self.send, '笔记部 Bot 邮件告警测试', '这是一封邮件通道测试，收到说明本次测试邮件已送达。不是实际掉线告警。')
                return '测试邮件已由发件服务器接受，请检查收件箱或垃圾邮件。'
            except Exception as exc:
                print('[邮件告警] 测试发送失败：' + type(exc).__name__)
                return '测试邮件发送失败，请检查本机 SMTP 配置、授权码和网络。'

    async def observe(self, healthy, reason, now=None):
        now = time.time() if now is None else now
        async with self.lock:
            config = self.settings()
            if not config.get('enabled'):
                return
            state = self.state
            if healthy:
                if not state:
                    return
                state.setdefault('healthy_since', now)
                if now - state['healthy_since'] >= 120:
                    self.state = {}
                write_json(self.state_path, self.state)
                return
            state.pop('healthy_since', None)
            state.setdefault('bad_since', now)
            if now - state['bad_since'] < 60 or state.get('notified'):
                write_json(self.state_path, state)
                return
            if state.get('attempts', 0) >= 3 or now - state.get('last_attempt', -1e20) < 300:
                return
            state['attempts'] = state.get('attempts', 0) + 1
            state['last_attempt'] = now
            write_json(self.state_path, state)
            try:
                await asyncio.to_thread(self.send, '笔记部 Bot：NapCat 连接异常',
                    'NapCat/QQ 连接异常持续至少一分钟。\n原因：' + reason +
                    '\n请检查本机 NapCat 登录页面及 Bot 日志；如需安全验证，请本人扫码。'
                    '\n自动重连仍在继续。这封邮件不表示程序已恢复，也不包含密码、令牌或聊天记录。')
                state['notified'] = True
                print('[邮件告警] 发件服务器已接受掉线告警。')
            except Exception as exc:
                print('[邮件告警] 发送失败：' + type(exc).__name__ + '；同一故障最多尝试三次，间隔五分钟。')
            write_json(self.state_path, state)

    async def monitor(self, client):
        while True:
            try:
                if not self.settings().get('enabled'):
                    await asyncio.sleep(15)
                    continue
                healthy = False
                reason = 'OneBot WebSocket 未连接'
                if client.ws is not None:
                    reason = 'QQ 在线状态检查失败或返回离线'
                    response = await client.call_api('get_status', timeout=10)
                    data = response.get('data') or {}
                    healthy = response.get('status') == 'ok' and response.get('retcode') == 0 and data.get('online') is True and data.get('good') is not False
                await self.observe(healthy, reason)
            except asyncio.CancelledError:
                raise
            except Exception as exc:
                print('[邮件告警] 状态检查异常：' + type(exc).__name__)
                try:
                    await self.observe(False, 'QQ 在线状态接口无法响应')
                except Exception as inner:
                    print('[邮件告警] 本地告警配置或状态不可用：' + type(inner).__name__)
            await asyncio.sleep(15)
