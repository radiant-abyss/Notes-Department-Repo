import asyncio,hashlib,importlib.util,json,sqlite3,tempfile
from contextlib import closing
from pathlib import Path
import db,bot

ROOT=Path(__file__).resolve().parent

def snapshot(path):
    with closing(sqlite3.connect(path)) as c:
        result={}
        for (name,) in c.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"):
            columns=[r[1] for r in c.execute(f'PRAGMA table_info("{name}")') if r[1] not in ('student_number','major_college','minor_college')]
            sql='SELECT '+','.join('"'+x+'"' for x in columns)+' FROM "'+name+'" ORDER BY rowid'
            result[name]=hashlib.sha256(repr(c.execute(sql).fetchall()).encode()).hexdigest()
        return result

with tempfile.TemporaryDirectory(prefix='notes-profile-') as tmp:
    tmp=Path(tmp)
    clone=tmp/'clone.db';db.DB_PATH=clone;db.init_db()
    db.register_user('migration','sample','测试成员')
    with db.connect() as conn:
        conn.execute("ALTER TABLE users DROP COLUMN student_number")
        conn.execute("ALTER TABLE users DROP COLUMN major_college")
        conn.execute("ALTER TABLE users DROP COLUMN minor_college")
    before=snapshot(clone);db.DB_PATH=clone;backup=db.migrate_user_profile()
    assert snapshot(clone)==before
    assert db.migrate_user_profile() is None
    if backup:assert backup.is_file()
    # Synthetic old member with a completed, reviewed assignment.
    spec=importlib.util.spec_from_file_location('old_db',ROOT/'db.py');old=importlib.util.module_from_spec(spec);spec.loader.exec_module(old)
    old.DB_PATH=tmp/'synthetic.db';old.init_db();old.register_user('w','100','旧成员')
    with old.connect() as c:
        c.execute("INSERT INTO tasks(id,workspace,task_type,subject,required_knowledge,publisher_qq,status,created_at,completed_at,required_people) VALUES(1,'w','笔记','数学','集合','100','completed','t','done',1)")
        c.execute("INSERT INTO task_assignments(task_id,workspace,qq,status,claimed_at,submitted_at,completed_at,submission_text,review_comment,reviewed_at,reviewer_qq) VALUES(1,'w','100','completed','t','s','done','成果','通过','done','owner')")
    before=snapshot(old.DB_PATH);db.DB_PATH=old.DB_PATH;db.migrate_user_profile()
    member=dict(db.get_user('w','100'))
    updated=db.update_user_profile('w','100','001234','未来技术学院-计算机学院','无')
    assert updated['id']==member['id'] and updated['student_number']=='001234'
    assert snapshot(db.DB_PATH)==before
    assert db.export_data('w')['users'][0]['minor_college']=='无'
    db.register_user('w','100','旧成员');assert db.get_user('w','100')['student_number']=='001234'
    for values in [('', '学院','无'),('123','未来技术学院','无'),('123','学院','')]:
        try:db.update_user_profile('w','100',*values)
        except ValueError:pass
        else:raise AssertionError('Invalid profile accepted')
    assert db.get_user('w','100')['student_number']=='001234'
    class Client:
        def __init__(self):self.messages=[]
        async def send_private(self,qq,text):self.messages.append((qq,text))
    async def exercise():
        b=bot.NotesDepartmentBot({'department_group_id':'w','owner_qq':'owner','owner_name':'负责人'});b.client=Client()
        async def member(qq):return True
        b.is_group_member=member
        async def send(text,qq='100'):await b.handle_private({'user_id':qq,'raw_message':text})
        await send('完善信息');await send('000567');await send('未来技术学院')
        assert len(bot.profile_sessions['100'])==1
        await send('未来技术学院-物理学院');await send('无')
        assert '100' not in bot.profile_sessions and db.get_user('w','100')['student_number']=='000567'
        assert snapshot(db.DB_PATH)==before
        await send('完善信息');await send('取消');assert '100' not in bot.profile_sessions
        await send('我的信息');assert '辅修学院：无' in b.client.messages[-1][1]
        await send('成员列表');assert '仅管理者' in b.client.messages[-1][1]
        await send('注册 新成员 | 000777 | 数学学院 | 无','101')
        assert db.get_user('w','101')['student_number']=='000777'
        await send('完善信息 000888 | 数学学院 | 物理学院','101')
        assert db.get_user('w','101')['minor_college']=='物理学院'
        await send('信息未齐成员','owner');assert '负责人' in b.client.messages[-1][1]
    asyncio.run(exercise())
print('PASS: synthetic legacy DB migration/idempotence; all pre-existing table values preserved; completed assignment preserved; leading zero; registration; self-service edit; validation; cancellation; authorization; export fields. No network/QQ messages.')
