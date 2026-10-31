"""Offline mail alert state machine and owner authorization tests."""
import asyncio
import tempfile
from pathlib import Path
from unittest.mock import patch
import email_alert as m
import bot

async def main():
    with tempfile.TemporaryDirectory() as folder:
        root = Path(folder)
        m.write_json(root/'email_alert.json', {'enabled':True,'recipient':'owner@example.com'})
        alert = m.EmailAlerts(root)
        sent=[]
        alert.send=lambda *args: sent.append(args)
        await alert.observe(False, 'offline', 0)
        await alert.observe(False, 'offline', 59)
        assert not sent
        await alert.observe(False, 'offline', 60)
        assert len(sent)==1
        await alert.observe(False, 'offline', 600)
        assert len(sent)==1
        alert = m.EmailAlerts(root)
        alert.send=lambda *args: sent.append(args)
        await alert.observe(False,'offline',900)
        assert len(sent)==1, 'restart must not resend same outage'
        await alert.observe(True,'',910)
        await alert.observe(False,'offline',920)
        assert len(sent)==1, 'brief reconnect must not rearm'
        await alert.observe(True,'',930)
        await alert.observe(True,'',1050)
        await alert.observe(False,'offline',1060)
        await alert.observe(False,'offline',1120)
        assert len(sent)==2
        alert.state={}
        failures=[]
        def fail(*args):
            failures.append(1)
            raise OSError('secret must not be logged')
        alert.send=fail
        for t in [0,60,61,359,360,660,960]:
            await alert.observe(False,'offline',t)
        assert len(failures)==3 and not alert.state.get('notified')
        before=alert.settings()
        for value in ['','bad','a@example.com\r\nBcc:x@y.com']:
            try: alert.configure_recipient(value)
            except ValueError: pass
            else: raise AssertionError(value)
        assert alert.settings()==before
        alert.configure_recipient('new@example.com')
        assert alert.settings()['recipient']=='new@example.com'
        alert.send=lambda *args:sent.append(args)
        await alert.test(); await alert.test()
        assert len(sent)==3, 'test mail throttle'
        app=bot.NotesDepartmentBot({'department_group_id':'w','owner_qq':'owner','owner_name':'Owner'})
        app.email_alerts=alert
        class Client:
            def __init__(self):self.messages=[]
            async def send_private(self,qq,text):self.messages.append(text)
        app.client=Client()
        user={'active':1,'qq':'other','role':'admin'}
        with patch.object(bot.db,'get_user',return_value=user),patch.object(bot.db,'ensure_owner'):
            await app.handle_private({'user_id':'other','raw_message':'配置邮箱 attacker@example.com'})
            assert alert.settings()['recipient']=='new@example.com'
            assert '仅总负责人' in app.client.messages[-1]
            await app.handle_private({'user_id':'owner','raw_message':'配置邮箱 owner@example.com'})
            assert alert.settings()['recipient']=='owner@example.com'
        # Monitor must catch QQ offline even with an open WebSocket.
        class Offline:
            ws=object()
            async def call_api(self,*a,**k):return {'status':'ok','retcode':0,'data':{'online':False}}
        seen=[]
        async def observe(healthy,reason):seen.append(healthy)
        async def stop(*args):raise asyncio.CancelledError
        alert.observe=observe
        with patch.object(m.asyncio,'sleep',stop):
            try:await alert.monitor(Offline())
            except asyncio.CancelledError:pass
        assert seen==[False]
    print('PASS: outage delay/dedup/restart/rearm/retry cap, validation/test throttle, owner authorization, QQ offline detection; no SMTP calls.')

asyncio.run(main())
