import asyncio,hashlib,json,sqlite3,tempfile
from contextlib import closing
from pathlib import Path
import db,bot
ROOT=Path(__file__).resolve().parent

def snapshot(path):
    with closing(sqlite3.connect(path)) as conn:
        result={}
        for (name,) in conn.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"):
            cols=[r[1] for r in conn.execute(f'PRAGMA table_info("{name}")') if r[1] not in ('progress_text','progress_updated_at')]
            data=conn.execute('SELECT '+','.join('"'+x+'"' for x in cols)+' FROM "'+name+'" ORDER BY rowid').fetchall()
            result[name]=hashlib.sha256(repr(data).encode()).hexdigest()
        return result

with tempfile.TemporaryDirectory(prefix='notes-progress-') as tmp:
    tmp=Path(tmp);clone=tmp/'clone.db'
    clone=tmp/'clone.db';db.DB_PATH=clone;db.init_db()
    db.register_user('migration','sample','测试成员')
    with db.connect() as conn:
        conn.execute("ALTER TABLE task_assignments DROP COLUMN progress_text")
        conn.execute("ALTER TABLE task_assignments DROP COLUMN progress_updated_at")
    before=snapshot(clone);db.DB_PATH=clone;backup=db.migrate_task_progress()
    assert backup and backup.exists() and snapshot(clone)==before
    assert db.migrate_task_progress() is None
    db.DB_PATH=tmp/'test.db';db.init_db()
    for qq in ('a','b','c'):db.register_user('w',qq,qq)
    db.ensure_owner('w','owner')
    tid=db.create_task('w','笔记','数学','集合',2,'owner')
    assert db.claim_task('w',tid,'a')=='ok' and db.claim_task('w',tid,'b')=='ok'
    original=snapshot(db.DB_PATH)
    assert db.update_task_progress('w',tid,'a','完成初稿')=='ok'
    assert snapshot(db.DB_PATH)==original
    assert db.update_task_progress('w',tid,'b','整理例题')=='ok'
    assert db.get_assignment('w',tid,'a')['progress_text']=='完成初稿'
    assert db.get_assignment('w',tid,'b')['progress_text']=='整理例题'
    assert db.update_task_progress('w',tid,'c','越权')=='not_claimed'
    assert db.update_task_progress('other',tid,'a','越界')=='inactive'
    assert db.update_task_progress('w',tid,'a','')=='empty'
    assert db.update_task_progress('w',tid,'a','长'*1001)=='too_long'
    assert db.submit_task('w',tid,'a','正式成果')=='ok'
    assert db.update_task_progress('w',tid,'a','不能改')=='submitted'
    summary=bot.todo_summary(db.get_user('w','a'),db.list_my_active_tasks('w','a'),'a')
    assert '等待审核：1 项' in summary and '我的待办：1 项待处理' in summary # missing profile only
    db.update_user_profile('w','a','001','计算学部','无')
    summary=bot.todo_summary(db.get_user('w','a'),db.list_my_active_tasks('w','a'),'a')
    assert '我的待办：0 项待处理' in summary
    assert db.reject_submission('w',tid,'a','owner','补充证明')=='ok'
    summary=bot.todo_summary(db.get_user('w','a'),db.list_my_active_tasks('w','a'),'a')
    assert '需修改后重新提交' in summary and '补充证明' in summary
    assert db.update_task_progress('w',tid,'a','已补证明')=='ok'
    assert db.submit_task('w',tid,'a','修改版')=='ok'
    assert db.approve_submission('w',tid,'a','owner')=='ok'
    assert db.update_task_progress('w',tid,'a','完成不能改')=='completed'
    assert not db.list_my_active_tasks('w','a')
    assert db.export_data('w')['assignments'][0]['progress_text']=='已补证明'
    class Client:
        def __init__(self):self.sent=[]
        async def send_private(self,qq,text):self.sent.append((qq,text))
        async def send_group(self,*args):raise AssertionError('Unexpected group message')
    async def run():
        app=bot.NotesDepartmentBot({'department_group_id':'w','owner_qq':'owner','owner_name':'负责人'});app.client=Client()
        async def send(text,qq='b'):await app.handle_private({'user_id':qq,'raw_message':text})
        await send(f'更新进度{tid} 已整理完例题')
        assert db.get_assignment('w',tid,'b')['progress_text']=='已整理完例题'
        await send('我的待办');assert '已整理完例题' in app.client.sent[-1][1]
        await send('我的任务');assert '更新时间' in app.client.sent[-1][1]
        await send(f'任务 {tid}','owner');assert '已补证明' in app.client.sent[-1][1] and '已整理完例题' in app.client.sent[-1][1]
        await send(f'任务 {tid}','c');assert '已补证明' not in app.client.sent[-1][1]
        await send('更新进度 无效');assert '用法' in app.client.sent[-1][1]
    asyncio.run(run())
    with db.connect() as conn:conn.execute("UPDATE tasks SET deleted_at='deleted' WHERE id=?",(tid,))
    assert db.update_task_progress('w',tid,'b','已删除')=='deleted'
    assert not db.list_my_active_tasks('w','b')
    with db.connect() as conn:conn.execute("UPDATE users SET active=0 WHERE qq='b'")
    assert db.update_task_progress('w',tid,'b','停用')=='inactive'
for name in ('bot.py','db.py'):compile(Path(name).read_text(encoding='utf-8'),name,'exec')
print('PASS: synthetic legacy DB migration; original data preserved; fresh init; idempotence; member isolation; task/review status unchanged by progress; submitted/completed/deleted/inactive guards; todo categories; private command integration; admin visibility; export; syntax. No real QQ messages.')
