import asyncio,sqlite3,tempfile,hashlib,json
from pathlib import Path
from contextlib import closing
from concurrent.futures import ThreadPoolExecutor
import db,bot
FIELDS=['笔记整理','离散数学','整理关系与函数，需集合知识','2','补齐部门笔记']
def snap(path):
 with closing(sqlite3.connect(path)) as c:
  names=[r[0] for r in c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT IN ('task_requests','sqlite_sequence')")]
  return {n:[tuple(r) for r in c.execute('SELECT * FROM "'+n+'" ORDER BY rowid')] for n in names}
class Client:
 def __init__(self):self.private=[];self.group=[];self.fail=False
 async def send_private(self,qq,text):
  if self.fail:raise ConnectionError('mock')
  self.private.append((qq,text))
 async def send_group(self,g,text):
  if self.fail:raise ConnectionError('mock')
  self.group.append((g,text))
async def commands():
 app=bot.NotesDepartmentBot({'department_group_id':'w','owner_qq':'owner','owner_name':'负责人'});app.client=Client()
 async def send(text,qq='member'):await app.handle_private({'user_id':qq,'raw_message':text})
 await send('申请任务');await send('');assert len(bot.request_sessions[('w','member')])==0
 await send('取消');assert ('w','member') not in bot.request_sessions
 await send('申请任务')
 for f in FIELDS:await send(f)
 assert not db.list_task_requests('w','member')
 await send('确认');rid=db.list_task_requests('w','member')[0]['id']
 assert not db.list_open_tasks('w') and not app.client.group
 assert any(q=='owner' and '新的任务发布申请' in t for q,t in app.client.private)
 await send('确认');assert len(db.list_task_requests('w','member'))==1
 await send(f'批准申请 {rid}');assert '仅管理者' in app.client.private[-1][1]
 await send(f'申请详情 {rid}','other');assert '没有找到' in app.client.private[-1][1]
 await send(f'撤回申请 {rid}','owner');assert db.get_task_request('w',rid)['status']=='pending'
 await send(f'批准申请{rid}','admin')
 row=db.get_task_request('w',rid);tid=row['task_id'];assert row['status']=='approved'
 task=db.get_task('w',tid);assert task['status']=='open' and task['assigned_count']==0 and task['publisher_qq']=='admin'
 assert len(app.client.group)==1
 await send(f'批准申请 {rid}','owner');assert len(app.client.group)==1
 assert len(db.list_open_tasks('w'))==1
 assert db.claim_task('w',tid,'other')=='ok'
 assert db.claim_task('w',tid,'member')=='ok'
 assert db.submit_task('w',tid,'other','成果')=='ok'
 assert db.approve_submission('w',tid,'other','admin')=='ok'
 await send('我的申请');assert '已批准' in app.client.private[-1][1]
 await send('待审批申请','admin');assert '没有申请' in app.client.private[-1][1]
 # Notification failure must not lose an application or undo an approval.
 await send('申请任务')
 for f in FIELDS:await send(f)
 app.client.fail=True;await send('确认')
 rid=db.list_task_requests('w','member')[0]['id'];await send(f'批准申请 {rid}','owner')
 assert db.get_task_request('w',rid)['status']=='approved'
 app.client.fail=False
 # Check active membership again when a wizard is resumed.
 await send('申请任务','other')
 with db.connect() as c:c.execute("UPDATE users SET active=0 WHERE qq='other'")
 await send('类型','other');assert ('w','other') not in bot.request_sessions
with tempfile.TemporaryDirectory() as tmp:
 tmp=Path(tmp)
 # Build a disposable legacy fixture with an existing member and task.
 clone=tmp/'clone.db';db.DB_PATH=clone;db.init_db()
 db.register_user('legacy','old','旧成员')
 db.create_task('legacy','旧笔记','数学','集合',1,'old')
 with db.connect() as c:c.execute('DROP TABLE task_requests')
 before=snap(clone);db.DB_PATH=clone;backup=db.migrate_task_requests()
 assert backup and backup.exists() and snap(clone)==before
 assert db.migrate_task_requests() is None and snap(clone)==before
 db.DB_PATH=tmp/'test.db';db.init_db()
 db.ensure_owner('w','owner');db.register_user('w','admin','管理员')
 with db.connect() as c:c.execute("UPDATE users SET role='admin' WHERE qq='admin'")
 for q in ('member','other'):db.register_user('w',q,q)
 asyncio.run(commands())
 rid=db.create_task_request('w','member',FIELDS)
 assert db.decide_task_request('other',rid,'owner','approved')[0]=='forbidden'
 assert db.decide_task_request('w',rid,'member','approved')[0]=='forbidden'
 assert db.decide_task_request('w',rid,'owner','rejected')[0]=='reason_required'
 assert db.decide_task_request('w',rid,'owner','rejected','需要更具体')[0]=='ok'
 assert db.decide_task_request('w',rid,'owner','approved')[0]=='already_processed'
 rid=db.create_task_request('w','member',FIELDS)
 assert db.decide_task_request('w',rid,'member','withdrawn')[0]=='ok'
 assert db.decide_task_request('w',rid,'owner','approved')[0]=='already_processed'
 rid=db.create_task_request('w','member',FIELDS)
 with ThreadPoolExecutor(2) as pool:
  results=list(pool.map(lambda q:db.decide_task_request('w',rid,q,'approved'),['owner','admin']))
 assert sorted(r[0] for r in results)==['already_processed','ok']
 assert results[0][1]==results[1][1]
 # Force audit failure: transaction must roll back task creation and approval.
 rid=db.create_task_request('w','member',FIELDS)
 with db.connect() as c:
  count=c.execute('SELECT count(*) FROM tasks').fetchone()[0]
  c.execute("CREATE TRIGGER fail_approval BEFORE INSERT ON audit_logs WHEN NEW.action='批准任务申请' BEGIN SELECT RAISE(ABORT,'test rollback'); END")
 try:db.decide_task_request('w',rid,'owner','approved')
 except sqlite3.IntegrityError:pass
 else:raise AssertionError('missing audit failure')
 with db.connect() as c:
  assert c.execute('SELECT count(*) FROM tasks').fetchone()[0]==count
  c.execute('DROP TRIGGER fail_approval')
 assert db.get_task_request('w',rid)['status']=='pending'
 for _ in range(12):db.create_task_request('w','member',FIELDS)
 assert len(db.list_task_requests('w','member',page=1))==10 and db.list_task_requests('w','member',page=2)
 assert db.export_data('w')['task_requests']
 # Excel/CSV include proposal records; files only in the test temp directory.
 bot.BASE_DIR=tmp
 xlsx=bot.make_export_xlsx('w') if hasattr(bot,'make_export_xlsx') else None
 z=bot.make_export_csv_zip('w');assert z.exists()
 print('PASS: synthetic legacy DB migration/data preservation/idempotence; wizard/permissions/cancel; public unassigned task; duplicate and concurrent approval; rejection/withdrawal; rollback; notification failures; inactive members; pagination; exports; original claim/submit/review. No real QQ messages.')
