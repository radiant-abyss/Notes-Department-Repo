"""Offline regression for menu state, progress dialogs and failed delivery."""
import asyncio
import contextlib
import io
import tempfile
import time
from pathlib import Path
import bot
import db

class Client:
    def __init__(self): self.messages = []
    async def send_private(self, qq, text): self.messages.append((qq, text))

async def checks():
    app = bot.NotesDepartmentBot({'department_group_id':'w','owner_qq':'owner','owner_name':'Owner'})
    app.client = Client()
    async def send(text, qq='a'):
        await app.handle_private({'user_id':qq,'raw_message':text})
    task = db.create_task('w','notes','math','sets',2,'owner')
    for qq in ('a','b'): assert db.claim_task('w',task,qq) == 'ok'
    await send('4'); assert ('w','a') not in bot.progress_sessions
    await send('帮助'); assert ('w','a') in bot.menu_sessions
    assert '全部帮助' not in app.client.messages[-1][1]
    await send('6','b'); assert ('w','a') in bot.menu_sessions
    await send('4'); await send(str(task)); assert ('w','a') in bot.progress_sessions
    await send('1'); assert db.get_assignment('w',task,'a')['progress_text'] == '1'
    assert not db.get_assignment('w',task,'b')['progress_text']
    await send('帮助'); await send('3'); await send(str(task)); assert 'a' in bot.submit_sessions
    await send('取消'); assert 'a' not in bot.submit_sessions
    await send('帮助'); await send('完善信息'); assert ('w','a') not in bot.menu_sessions
    await send('帮助'); assert ('w','a') not in bot.menu_sessions
    await send('0012345678'); assert bot.profile_sessions['a'] == ['0012345678']
    await send('取消'); await send('帮助'); await send('7')
    for value in ('notes','math','sets','2'): await send(value)
    assert bot.request_sessions[('w','a')][3] == '2'
    await send('取消'); await send('帮助'); await send('99')
    assert ('w','a') in bot.menu_sessions
    await send('取消'); assert ('w','a') not in bot.menu_sessions
    await send('帮助'); bot.menu_sessions[('w','a')] = ('choose',time.monotonic()-1)
    await send('5'); assert '过期' in app.client.messages[-1][1]
    for option in ('1','2','5','6'):
        await send('帮助'); await send(option); assert ('w','a') not in bot.menu_sessions
    await send(f'更新进度{task}'); await send(''); await send('字'*1001)
    assert ('w','a') in bot.progress_sessions
    assert db.submit_task('w',task,'a','result') == 'ok'
    await send('stale'); assert ('w','a') not in bot.progress_sessions
    assert db.get_assignment('w',task,'a')['progress_text'] == '1'
    await send(f'更新进度 {task}','owner'); assert ('w','owner') not in bot.progress_sessions
    await send('帮助','b'); db.deactivate_user('w','b'); await send('4','b')
    assert ('w','b') not in bot.menu_sessions
    await send('帮助','unknown'); assert ('w','unknown') not in bot.menu_sessions
    await send('全部帮助'); assert '发布任务' not in app.client.messages[-1][1]

    # User-facing send methods must surface failures, without dumping private response data.
    class FakeSender:
        async def call_api(self, action, params): return self.response
        send_message_checked = bot.OneBotClient.send_message_checked
    sender = FakeSender()
    sender.response = {'status':'ok','retcode':0,'data':{'message_id':1}}
    assert (await bot.OneBotClient.send_private(sender,'123','hello'))['retcode'] == 0
    for detail in ('1006514 secret-value','Timeout secret-value','secret-value'):
        sender.response = {'status':'failed','retcode':1200,'message':detail}
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            try: await bot.OneBotClient.send_private(sender,'123','hello')
            except RuntimeError as exc: assert 'secret-value' not in str(exc)
            else: raise AssertionError('send failure ignored')
        assert 'secret-value' not in out.getvalue()
        assert '消息发送失败' in out.getvalue()

with tempfile.TemporaryDirectory(prefix='notes-menu-test-') as folder:
    db.DB_PATH = Path(folder)/'test.db'
    db.init_db(); db.ensure_owner('w','owner')
    for qq in ('a','b'): db.register_user('w',qq,qq)
    asyncio.run(checks())
print('PASS: menu/session isolation, numeric fields, progress state checks, expiry/cancel, send error propagation and redaction. No real messages.')
