from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Optional

DB_PATH = Path(__file__).with_name("notes_department.db")


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


@contextmanager
def connect():
    conn = sqlite3.connect(DB_PATH, timeout=10)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA foreign_keys=ON;")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _table_exists(conn: sqlite3.Connection, table: str) -> bool:
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)
    ).fetchone() is not None


def _column_names(conn: sqlite3.Connection, table: str) -> set[str]:
    if not _table_exists(conn, table):
        return set()
    return {row["name"] for row in conn.execute(f"PRAGMA table_info({table})")}


def _backup_before_migration(conn: sqlite3.Connection, version: str = "v5") -> Optional[Path]:
    if not DB_PATH.exists():
        return None
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = DB_PATH.with_name(f"notes_department.backup_before_{version}_{stamp}.db")
    dst = sqlite3.connect(backup_path)
    try:
        conn.backup(dst)
    finally:
        dst.close()
    return backup_path


def _assignment_schema_is_v5(conn: sqlite3.Connection) -> bool:
    if not _table_exists(conn, "task_assignments"):
        return False
    sql_row = conn.execute(
        "SELECT sql FROM sqlite_master WHERE type='table' AND name='task_assignments'"
    ).fetchone()
    sql = (sql_row["sql"] or "") if sql_row else ""
    cols = _column_names(conn, "task_assignments")
    required = {
        "submission_text", "submitted_at", "review_comment", "reviewed_at", "reviewer_qq"
    }
    return "submitted" in sql and required.issubset(cols)


def init_db() -> None:
    """初始化并自动迁移旧数据库到 v5。

    v5 在 v4 多人任务基础上新增：
    - task_assignments 支持 claimed / submitted / completed；
    - 轻量成果说明、提交时间、审核意见、审核人；
    - audit_logs 管理员操作日志；
    - 保留全部既有成员、任务、认领与完成记录。
    """
    with connect() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                workspace TEXT NOT NULL,
                qq TEXT NOT NULL,
                name TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'member'
                    CHECK(role IN ('owner', 'admin', 'member')),
                registered_at TEXT NOT NULL,
                active INTEGER NOT NULL DEFAULT 1,
                UNIQUE(workspace, qq)
            );

            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                workspace TEXT NOT NULL,
                task_type TEXT NOT NULL,
                subject TEXT NOT NULL,
                required_knowledge TEXT NOT NULL,
                publisher_qq TEXT NOT NULL,
                assignee_qq TEXT,
                status TEXT NOT NULL DEFAULT 'open'
                    CHECK(status IN ('open', 'claimed', 'completed')),
                created_at TEXT NOT NULL,
                claimed_at TEXT,
                completed_at TEXT
            );
            """
        )

        task_cols = _column_names(conn, "tasks")
        needs_v4 = (
            "required_people" not in task_cols
            or "deleted_at" not in task_cols
            or not _table_exists(conn, "task_assignments")
        )
        needs_v5 = _table_exists(conn, "task_assignments") and not _assignment_schema_is_v5(conn)
        backup_path = _backup_before_migration(conn, "v5") if (needs_v4 or needs_v5) else None

        if "required_people" not in task_cols:
            conn.execute(
                "ALTER TABLE tasks ADD COLUMN required_people INTEGER NOT NULL DEFAULT 1"
            )
        if "deleted_at" not in task_cols:
            conn.execute("ALTER TABLE tasks ADD COLUMN deleted_at TEXT")

        # 新安装直接创建 v5 assignment 表。
        if not _table_exists(conn, "task_assignments"):
            conn.executescript(
                """
                CREATE TABLE task_assignments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_id INTEGER NOT NULL,
                    workspace TEXT NOT NULL,
                    qq TEXT NOT NULL,
                    status TEXT NOT NULL DEFAULT 'claimed'
                        CHECK(status IN ('claimed', 'submitted', 'completed')),
                    claimed_at TEXT NOT NULL,
                    submitted_at TEXT,
                    completed_at TEXT,
                    submission_text TEXT,
                    review_comment TEXT,
                    reviewed_at TEXT,
                    reviewer_qq TEXT,
                    UNIQUE(task_id, qq),
                    FOREIGN KEY(task_id) REFERENCES tasks(id) ON DELETE CASCADE
                );
                """
            )
        elif not _assignment_schema_is_v5(conn):
            # SQLite 无法直接修改 CHECK(status...)，因此重建表并复制旧记录。
            conn.execute("DROP TABLE IF EXISTS task_assignments_v4_old")
            conn.execute("ALTER TABLE task_assignments RENAME TO task_assignments_v4_old")
            conn.executescript(
                """
                CREATE TABLE task_assignments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_id INTEGER NOT NULL,
                    workspace TEXT NOT NULL,
                    qq TEXT NOT NULL,
                    status TEXT NOT NULL DEFAULT 'claimed'
                        CHECK(status IN ('claimed', 'submitted', 'completed')),
                    claimed_at TEXT NOT NULL,
                    submitted_at TEXT,
                    completed_at TEXT,
                    submission_text TEXT,
                    review_comment TEXT,
                    reviewed_at TEXT,
                    reviewer_qq TEXT,
                    UNIQUE(task_id, qq),
                    FOREIGN KEY(task_id) REFERENCES tasks(id) ON DELETE CASCADE
                );
                """
            )
            old_cols = _column_names(conn, "task_assignments_v4_old")
            select_cols = ["id", "task_id", "workspace", "qq", "status", "claimed_at"]
            for c in ["submitted_at", "completed_at", "submission_text", "review_comment", "reviewed_at", "reviewer_qq"]:
                select_cols.append(c if c in old_cols else f"NULL AS {c}")
            rows = conn.execute(
                f"SELECT {', '.join(select_cols)} FROM task_assignments_v4_old"
            ).fetchall()
            for r in rows:
                status = r["status"] if r["status"] in {"claimed", "submitted", "completed"} else "claimed"
                conn.execute(
                    """
                    INSERT INTO task_assignments
                    (id, task_id, workspace, qq, status, claimed_at, submitted_at, completed_at,
                     submission_text, review_comment, reviewed_at, reviewer_qq)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        r["id"], r["task_id"], r["workspace"], r["qq"], status,
                        r["claimed_at"], r["submitted_at"], r["completed_at"],
                        r["submission_text"], r["review_comment"], r["reviewed_at"], r["reviewer_qq"],
                    ),
                )
            conn.execute("DROP TABLE task_assignments_v4_old")

        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS audit_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                workspace TEXT NOT NULL,
                actor_qq TEXT NOT NULL,
                action TEXT NOT NULL,
                detail TEXT,
                created_at TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_users_workspace_role
            ON users(workspace, role);

            CREATE INDEX IF NOT EXISTS idx_tasks_workspace_status
            ON tasks(workspace, status);

            CREATE INDEX IF NOT EXISTS idx_assignments_task
            ON task_assignments(task_id, status);

            CREATE INDEX IF NOT EXISTS idx_assignments_user
            ON task_assignments(workspace, qq, status);

            CREATE INDEX IF NOT EXISTS idx_audit_workspace_time
            ON audit_logs(workspace, id DESC);
            """
        )

        # 从最早期单人任务字段迁移到 assignment 表，幂等执行。
        task_cols = _column_names(conn, "tasks")
        if "assignee_qq" in task_cols:
            rows = conn.execute(
                """
                SELECT id, workspace, assignee_qq, status, claimed_at, completed_at, created_at
                FROM tasks
                WHERE assignee_qq IS NOT NULL AND TRIM(assignee_qq) <> ''
                """
            ).fetchall()
            for row in rows:
                a_status = "completed" if row["status"] == "completed" else "claimed"
                claimed_at = row["claimed_at"] or row["created_at"] or now_text()
                completed_at = row["completed_at"] if a_status == "completed" else None
                conn.execute(
                    """
                    INSERT OR IGNORE INTO task_assignments
                    (task_id, workspace, qq, status, claimed_at, completed_at)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (row["id"], row["workspace"], str(row["assignee_qq"]), a_status, claimed_at, completed_at),
                )

        conn.execute(
            "UPDATE tasks SET required_people=1 WHERE required_people IS NULL OR required_people < 1"
        )
        ids = [r["id"] for r in conn.execute("SELECT id FROM tasks").fetchall()]
        for task_id in ids:
            _sync_task_status(conn, task_id)

        if backup_path:
            print(f"[数据库] 已自动迁移到 v5；迁移前备份：{backup_path.name}")
    migrate_user_profile()
    migrate_task_progress()
    migrate_task_requests()
    import volunteer
    volunteer.migrate(connect, _backup_before_migration)



def migrate_task_progress():
    """Add per-member latest progress without changing assignment status."""
    with connect() as conn:
        cols = _column_names(conn, "task_assignments")
        if not cols:
            raise RuntimeError("任务认领表不存在，请先初始化数据库")
        missing = [name for name in ("progress_text", "progress_updated_at") if name not in cols]
        if not missing:
            return None
        backup = _backup_before_migration(conn, "progress_v7")
        conn.execute("BEGIN IMMEDIATE")
        cols = _column_names(conn, "task_assignments")
        for name in missing:
            if name not in cols:
                conn.execute(f"ALTER TABLE task_assignments ADD COLUMN {name} TEXT NOT NULL DEFAULT ''")
        return backup


def update_task_progress(workspace: str, task_id: int, qq: str, text: str) -> str:
    text = text.strip()
    if not text:
        return "empty"
    if len(text) > 1000:
        return "too_long"
    with connect() as conn:
        conn.execute("BEGIN IMMEDIATE")
        if not conn.execute("SELECT 1 FROM users WHERE workspace=? AND qq=? AND active=1", (workspace, str(qq))).fetchone():
            return "inactive"
        task = conn.execute("SELECT status, deleted_at FROM tasks WHERE workspace=? AND id=?", (workspace, task_id)).fetchone()
        if not task:
            return "not_found"
        if task["deleted_at"] is not None:
            return "deleted"
        assignment = conn.execute("SELECT status FROM task_assignments WHERE workspace=? AND task_id=? AND qq=?", (workspace, task_id, str(qq))).fetchone()
        if not assignment:
            return "not_claimed"
        if assignment["status"] == "completed" or task["status"] == "completed":
            return "completed"
        if assignment["status"] == "submitted":
            return "submitted"
        conn.execute("UPDATE task_assignments SET progress_text=?, progress_updated_at=? WHERE workspace=? AND task_id=? AND qq=? AND status='claimed'", (text, now_text(), workspace, task_id, str(qq)))
        return "ok"


def migrate_user_profile():
    """Add columns without replacing member rows or task records."""
    with connect() as conn:
        cols = _column_names(conn, "users")
        if not cols:
            raise RuntimeError("成员表不存在，请先初始化数据库")
        missing = [c for c in ("student_number", "major_college", "minor_college") if c not in cols]
        if not missing:
            return None
        backup = _backup_before_migration(conn, "profile_v6")
        conn.execute("BEGIN IMMEDIATE")
        cols = _column_names(conn, "users")
        for column in missing:
            if column not in cols:
                conn.execute(f"ALTER TABLE users ADD COLUMN {column} TEXT NOT NULL DEFAULT ''")
        return backup


def validate_user_profile(student_number, major_college, minor_college):
    values = [str(v).strip() if v is not None else "" for v in (student_number, major_college, minor_college)]
    for label, value in zip(("学号", "学院（主修）", "辅修学院"), values):
        if not value:
            raise ValueError(f"{label}不能为空；没有辅修学院请填“无”")
        if len(value) > 100 or any(ch in value for ch in "\r\n|｜"):
            raise ValueError(f"{label}格式不正确，请填写不超过 100 字的单项内容")
    if any(ch.isspace() for ch in values[0]):
        raise ValueError("学号中不能包含空格")
    if values[1].startswith("未来技术学院") and (not values[1].startswith("未来技术学院-") or not values[1][len("未来技术学院-"):].strip()):
        raise ValueError("未来技术学院请填写：未来技术学院-某某学院")
    return tuple(values)


def update_user_profile(workspace, qq, student_number, major_college, minor_college):
    values = validate_user_profile(student_number, major_college, minor_college)
    with connect() as conn:
        result = conn.execute(
            "UPDATE users SET student_number=?, major_college=?, minor_college=? WHERE workspace=? AND qq=? AND active=1",
            (*values, workspace, str(qq)),
        )
        if result.rowcount != 1:
            raise ValueError("未找到有效成员，请先注册或联系负责人")
    return get_user(workspace, str(qq))


def get_user(workspace: str, qq: str, include_inactive: bool = False):
    sql = "SELECT * FROM users WHERE workspace=? AND qq=?"
    args = [workspace, str(qq)]
    if not include_inactive:
        sql += " AND active=1"
    with connect() as conn:
        return conn.execute(sql, args).fetchone()


def ensure_owner(workspace: str, owner_qq: str, owner_name: str = "Owner"):
    owner_qq = str(owner_qq)
    with connect() as conn:
        row = conn.execute(
            "SELECT * FROM users WHERE workspace=? AND qq=?",
            (workspace, owner_qq),
        ).fetchone()
        if row:
            conn.execute(
                "UPDATE users SET role='owner', active=1, name=? WHERE workspace=? AND qq=?",
                (owner_name, workspace, owner_qq),
            )
        else:
            conn.execute(
                """
                INSERT INTO users(workspace, qq, name, role, registered_at, active)
                VALUES (?, ?, ?, 'owner', ?, 1)
                """,
                (workspace, owner_qq, owner_name, now_text()),
            )
    return get_user(workspace, owner_qq)


def register_user(workspace: str, qq: str, name: str, owner_qq: Optional[str] = None, *, student_number=None, major_college=None, minor_college=None):
    qq = str(qq)
    name = name.strip()
    if not name:
        raise ValueError("姓名不能为空")
    profile = None
    if any(v is not None for v in (student_number, major_college, minor_college)):
        profile = validate_user_profile(student_number, major_college, minor_college)
    with connect() as conn:
        row = conn.execute(
            "SELECT * FROM users WHERE workspace=? AND qq=?",
            (workspace, qq),
        ).fetchone()
        if row:
            role = row["role"]
            if owner_qq and qq == str(owner_qq):
                role = "owner"
            conn.execute(
                "UPDATE users SET name=?, role=?, active=1 WHERE workspace=? AND qq=?",
                (name, role, workspace, qq),
            )
        else:
            role = "owner" if owner_qq and qq == str(owner_qq) else "member"
            conn.execute(
                """
                INSERT INTO users(workspace, qq, name, role, registered_at, active)
                VALUES (?, ?, ?, ?, ?, 1)
                """,
                (workspace, qq, name, role, now_text()),
            )
        if profile is not None:
            conn.execute("UPDATE users SET student_number=?, major_college=?, minor_college=? WHERE workspace=? AND qq=?", (*profile, workspace, qq))
    return get_user(workspace, qq)


def list_users(workspace: str):
    with connect() as conn:
        return conn.execute(
            """
            SELECT * FROM users
            WHERE workspace=? AND active=1
            ORDER BY CASE role WHEN 'owner' THEN 0 WHEN 'admin' THEN 1 ELSE 2 END,
                     id ASC
            """,
            (workspace,),
        ).fetchall()


def list_admins(workspace: str):
    with connect() as conn:
        return conn.execute(
            """
            SELECT * FROM users u
            WHERE workspace=? AND active=1 AND (role IN ('owner', 'admin') OR EXISTS (SELECT 1 FROM volunteer_managers v WHERE v.workspace=u.workspace AND v.qq=u.qq))
            ORDER BY CASE role WHEN 'owner' THEN 0 ELSE 1 END, id ASC
            """,
            (workspace,),
        ).fetchall()


def list_inactive_users(workspace: str):
    with connect() as conn:
        return conn.execute(
            "SELECT * FROM users WHERE workspace=? AND active=0 ORDER BY id ASC",
            (workspace,),
        ).fetchall()


def resolve_user(workspace: str, key: str, active: Optional[bool] = True):
    """按 QQ 或精确姓名查找。active=True/False/None 分别代表活跃/停用/全部。"""
    key = key.strip()
    active_sql = ""
    if active is True:
        active_sql = " AND active=1"
    elif active is False:
        active_sql = " AND active=0"

    with connect() as conn:
        exact = conn.execute(
            f"SELECT * FROM users WHERE workspace=? AND qq=?{active_sql}",
            (workspace, key),
        ).fetchone()
        if exact:
            return ("ok", exact)

        rows = conn.execute(
            f"SELECT * FROM users WHERE workspace=? AND name=?{active_sql} ORDER BY id ASC",
            (workspace, key),
        ).fetchall()
        if len(rows) == 1:
            return ("ok", rows[0])
        if len(rows) > 1:
            return ("ambiguous", rows)
        return ("not_found", None)


def set_role(workspace: str, qq: str, role: str) -> bool:
    if role not in {"owner", "admin", "member"}:
        raise ValueError("非法角色")
    with connect() as conn:
        cur = conn.execute(
            "UPDATE users SET role=? WHERE workspace=? AND qq=? AND active=1",
            (role, workspace, str(qq)),
        )
        return cur.rowcount == 1


def _sync_task_status(conn: sqlite3.Connection, task_id: int) -> None:
    task = conn.execute(
        "SELECT id, required_people, status FROM tasks WHERE id=?", (task_id,)
    ).fetchone()
    if not task:
        return
    counts = conn.execute(
        """
        SELECT COUNT(*) AS assigned_count,
               SUM(CASE WHEN status='submitted' THEN 1 ELSE 0 END) AS submitted_count,
               SUM(CASE WHEN status='completed' THEN 1 ELSE 0 END) AS completed_count
        FROM task_assignments
        WHERE task_id=?
        """,
        (task_id,),
    ).fetchone()
    assigned = int(counts["assigned_count"] or 0)
    completed = int(counts["completed_count"] or 0)
    required = max(1, int(task["required_people"] or 1))

    if assigned >= required and completed >= required:
        conn.execute(
            """
            UPDATE tasks
            SET status='completed', completed_at=COALESCE(completed_at, ?)
            WHERE id=?
            """,
            (now_text(), task_id),
        )
    elif assigned > 0:
        conn.execute(
            "UPDATE tasks SET status='claimed', completed_at=NULL WHERE id=?",
            (task_id,),
        )
    else:
        conn.execute(
            "UPDATE tasks SET status='open', completed_at=NULL WHERE id=?",
            (task_id,),
        )


def deactivate_user(workspace: str, qq: str) -> bool:
    """停用成员，并释放其所有尚未完成的认领；已完成历史保留。"""
    qq = str(qq)
    with connect() as conn:
        row = conn.execute(
            "SELECT role FROM users WHERE workspace=? AND qq=? AND active=1",
            (workspace, qq),
        ).fetchone()
        if not row or row["role"] == "owner":
            return False
        conn.execute(
            "UPDATE users SET active=0, role='member' WHERE workspace=? AND qq=?",
            (workspace, qq),
        )
        affected = [
            r["task_id"]
            for r in conn.execute(
                """
                SELECT task_id FROM task_assignments
                WHERE workspace=? AND qq=? AND status='claimed'
                """,
                (workspace, qq),
            ).fetchall()
        ]
        conn.execute(
            """
            DELETE FROM task_assignments
            WHERE workspace=? AND qq=? AND status='claimed'
            """,
            (workspace, qq),
        )
        for task_id in affected:
            _sync_task_status(conn, task_id)
        return True


def reactivate_user(workspace: str, qq: str) -> bool:
    with connect() as conn:
        cur = conn.execute(
            "UPDATE users SET active=1 WHERE workspace=? AND qq=? AND active=0",
            (workspace, str(qq)),
        )
        return cur.rowcount == 1


def create_task(
    workspace: str,
    task_type: str,
    subject: str,
    required_knowledge: str,
    required_people: int,
    publisher_qq: str,
) -> int:
    required_people = int(required_people)
    if required_people < 1:
        raise ValueError("人数需求必须至少为 1")
    with connect() as conn:
        cur = conn.execute(
            """
            INSERT INTO tasks
            (workspace, task_type, subject, required_knowledge, publisher_qq,
             status, created_at, required_people, deleted_at)
            VALUES (?, ?, ?, ?, ?, 'open', ?, ?, NULL)
            """,
            (
                workspace,
                task_type.strip(),
                subject.strip(),
                required_knowledge.strip(),
                str(publisher_qq),
                now_text(),
                required_people,
            ),
        )
        return int(cur.lastrowid)


def _get_task_with_conn(
    conn: sqlite3.Connection,
    workspace: str,
    task_id: int,
    include_deleted: bool = False,
):
    sql = """
        SELECT t.*, pu.name AS publisher_name
        FROM tasks t
        LEFT JOIN users pu ON pu.workspace=t.workspace AND pu.qq=t.publisher_qq
        WHERE t.workspace=? AND t.id=?
    """
    if not include_deleted:
        sql += " AND t.deleted_at IS NULL"
    row = conn.execute(sql, (workspace, task_id)).fetchone()
    if not row:
        return None

    task = dict(row)
    assignees = conn.execute(
        """
        SELECT a.qq, a.status, a.claimed_at, a.submitted_at, a.completed_at,
               a.progress_text, a.progress_updated_at,
               a.submission_text, a.review_comment, a.reviewed_at, a.reviewer_qq,
               COALESCE(u.name, a.qq) AS name,
               COALESCE(u.active, 0) AS user_active,
               COALESCE(rv.name, a.reviewer_qq) AS reviewer_name
        FROM task_assignments a
        LEFT JOIN users u ON u.workspace=a.workspace AND u.qq=a.qq
        LEFT JOIN users rv ON rv.workspace=a.workspace AND rv.qq=a.reviewer_qq
        WHERE a.task_id=? AND a.workspace=?
        ORDER BY a.id ASC
        """,
        (task_id, workspace),
    ).fetchall()
    task["assignees"] = [dict(r) for r in assignees]
    task["assigned_count"] = len(assignees)
    task["submitted_count"] = sum(1 for r in assignees if r["status"] == "submitted")
    task["completed_count"] = sum(1 for r in assignees if r["status"] == "completed")
    return task


def get_task(workspace: str, task_id: int, include_deleted: bool = False):
    with connect() as conn:
        return _get_task_with_conn(conn, workspace, task_id, include_deleted)


def _list_task_ids(conn: sqlite3.Connection, workspace: str, where: str, args: tuple, limit: int):
    rows = conn.execute(
        f"""
        SELECT id FROM tasks
        WHERE workspace=? {where}
        ORDER BY id DESC LIMIT ?
        """,
        (workspace, *args, limit),
    ).fetchall()
    return [r["id"] for r in rows]


def list_open_tasks(workspace: str, limit: int = 20):
    """列出仍有空位、可继续认领的任务。"""
    with connect() as conn:
        ids = _list_task_ids(
            conn,
            workspace,
            "AND deleted_at IS NULL AND status<>'completed'",
            (),
            max(limit * 4, 50),
        )
        result = []
        for task_id in ids:
            task = _get_task_with_conn(conn, workspace, task_id)
            if task and task["assigned_count"] < task["required_people"]:
                result.append(task)
                if len(result) >= limit:
                    break
        return result


def list_all_tasks(workspace: str, limit: int = 30):
    with connect() as conn:
        ids = _list_task_ids(conn, workspace, "AND deleted_at IS NULL", (), limit)
        return [_get_task_with_conn(conn, workspace, tid) for tid in ids]


def list_deleted_tasks(workspace: str, limit: int = 30):
    with connect() as conn:
        ids = _list_task_ids(conn, workspace, "AND deleted_at IS NOT NULL", (), limit)
        return [_get_task_with_conn(conn, workspace, tid, include_deleted=True) for tid in ids]


def list_my_active_tasks(workspace: str, qq: str):
    with connect() as conn:
        ids = [
            r["task_id"]
            for r in conn.execute(
                """
                SELECT a.task_id
                FROM task_assignments a
                JOIN tasks t ON t.id=a.task_id
                WHERE a.workspace=? AND a.qq=? AND a.status IN ('claimed', 'submitted')
                  AND t.deleted_at IS NULL
                ORDER BY a.task_id DESC
                """,
                (workspace, str(qq)),
            ).fetchall()
        ]
        return [_get_task_with_conn(conn, workspace, tid) for tid in ids]


def list_my_completed_tasks(workspace: str, qq: str, limit: int = 20):
    with connect() as conn:
        ids = [
            r["task_id"]
            for r in conn.execute(
                """
                SELECT a.task_id
                FROM task_assignments a
                JOIN tasks t ON t.id=a.task_id
                WHERE a.workspace=? AND a.qq=? AND a.status='completed'
                  AND t.deleted_at IS NULL
                ORDER BY a.completed_at DESC, a.task_id DESC LIMIT ?
                """,
                (workspace, str(qq), limit),
            ).fetchall()
        ]
        return [_get_task_with_conn(conn, workspace, tid) for tid in ids]


def claim_task(workspace: str, task_id: int, qq: str) -> str:
    qq = str(qq)
    with connect() as conn:
        conn.execute("BEGIN IMMEDIATE")
        task = conn.execute(
            """
            SELECT * FROM tasks
            WHERE workspace=? AND id=?
            """,
            (workspace, task_id),
        ).fetchone()
        if not task:
            return "not_found"
        if task["deleted_at"] is not None:
            return "deleted"
        if task["status"] == "completed":
            return "completed"
        existing = conn.execute(
            "SELECT status FROM task_assignments WHERE task_id=? AND qq=?",
            (task_id, qq),
        ).fetchone()
        if existing:
            return "already_claimed"
        count = conn.execute(
            "SELECT COUNT(*) AS n FROM task_assignments WHERE task_id=?",
            (task_id,),
        ).fetchone()["n"]
        if int(count) >= int(task["required_people"]):
            return "full"

        conn.execute(
            """
            INSERT INTO task_assignments(task_id, workspace, qq, status, claimed_at, completed_at)
            VALUES (?, ?, ?, 'claimed', ?, NULL)
            """,
            (task_id, workspace, qq, now_text()),
        )
        _sync_task_status(conn, task_id)
        return "ok"


def release_task(workspace: str, task_id: int, requester_qq: str) -> str:
    """成员取消自己的未完成认领。"""
    with connect() as conn:
        task = conn.execute(
            "SELECT id, deleted_at FROM tasks WHERE workspace=? AND id=?",
            (workspace, task_id),
        ).fetchone()
        if not task:
            return "not_found"
        if task["deleted_at"] is not None:
            return "deleted"
        cur = conn.execute(
            """
            DELETE FROM task_assignments
            WHERE task_id=? AND workspace=? AND qq=? AND status='claimed'
            """,
            (task_id, workspace, str(requester_qq)),
        )
        if cur.rowcount != 1:
            return "not_claimed"
        _sync_task_status(conn, task_id)
        return "ok"


def remove_assignment(workspace: str, task_id: int, target_qq: str) -> str:
    """管理员移除指定成员的未完成认领。"""
    with connect() as conn:
        task = conn.execute(
            "SELECT id, deleted_at FROM tasks WHERE workspace=? AND id=?",
            (workspace, task_id),
        ).fetchone()
        if not task:
            return "not_found"
        if task["deleted_at"] is not None:
            return "deleted"
        cur = conn.execute(
            """
            DELETE FROM task_assignments
            WHERE task_id=? AND workspace=? AND qq=? AND status='claimed'
            """,
            (task_id, workspace, str(target_qq)),
        )
        if cur.rowcount != 1:
            return "not_claimed"
        _sync_task_status(conn, task_id)
        return "ok"


def get_assignment(workspace: str, task_id: int, qq: str):
    with connect() as conn:
        row = conn.execute(
            """
            SELECT * FROM task_assignments
            WHERE workspace=? AND task_id=? AND qq=?
            """,
            (workspace, task_id, str(qq)),
        ).fetchone()
        return dict(row) if row else None


def submit_task(workspace: str, task_id: int, requester_qq: str, submission_text: str) -> str:
    submission_text = str(submission_text).strip()
    if not submission_text:
        return "empty"
    with connect() as conn:
        task = conn.execute(
            "SELECT id, deleted_at FROM tasks WHERE workspace=? AND id=?",
            (workspace, task_id),
        ).fetchone()
        if not task:
            return "not_found"
        if task["deleted_at"] is not None:
            return "deleted"
        row = conn.execute(
            "SELECT status FROM task_assignments WHERE task_id=? AND workspace=? AND qq=?",
            (task_id, workspace, str(requester_qq)),
        ).fetchone()
        if not row:
            return "not_claimed"
        if row["status"] == "submitted":
            return "already_submitted"
        if row["status"] == "completed":
            return "already_completed"
        conn.execute(
            """
            UPDATE task_assignments
            SET status='submitted', submission_text=?, submitted_at=?,
                review_comment=NULL, reviewed_at=NULL, reviewer_qq=NULL
            WHERE task_id=? AND workspace=? AND qq=? AND status='claimed'
            """,
            (submission_text, now_text(), task_id, workspace, str(requester_qq)),
        )
        _sync_task_status(conn, task_id)
        return "ok"


def withdraw_submission(workspace: str, task_id: int, requester_qq: str) -> str:
    with connect() as conn:
        cur = conn.execute(
            """
            UPDATE task_assignments
            SET status='claimed', submitted_at=NULL, reviewer_qq=NULL, reviewed_at=NULL,
                review_comment='成员主动撤回提交'
            WHERE task_id=? AND workspace=? AND qq=? AND status='submitted'
            """,
            (task_id, workspace, str(requester_qq)),
        )
        if cur.rowcount != 1:
            row = conn.execute(
                "SELECT status FROM task_assignments WHERE task_id=? AND workspace=? AND qq=?",
                (task_id, workspace, str(requester_qq)),
            ).fetchone()
            return "not_submitted" if row else "not_claimed"
        _sync_task_status(conn, task_id)
        return "ok"


def approve_submission(workspace: str, task_id: int, target_qq: str, reviewer_qq: str, volunteer_hours=None) -> str:
    import volunteer
    units = None if volunteer_hours is None else volunteer.parse_hours(volunteer_hours)
    with connect() as conn:
        conn.execute("BEGIN IMMEDIATE")
        if not conn.execute("SELECT 1 FROM users WHERE workspace=? AND qq=? AND active=1 AND (role IN ('owner','admin') OR EXISTS (SELECT 1 FROM volunteer_managers m WHERE m.workspace=users.workspace AND m.qq=users.qq))", (workspace, str(reviewer_qq))).fetchone():
            raise ValueError("仅有效管理者可审核任务。")
        if units is not None and not volunteer._admin(conn, workspace, reviewer_qq):
            raise ValueError("只有总负责人或高级管理员可以登记志愿时长；普通管理员请不带小时数审核。")
        task = conn.execute(
            "SELECT id, deleted_at FROM tasks WHERE workspace=? AND id=?",
            (workspace, task_id),
        ).fetchone()
        if not task:
            return "not_found"
        if task["deleted_at"] is not None:
            return "deleted"
        cur = conn.execute(
            """
            UPDATE task_assignments
            SET status='completed', completed_at=?, reviewed_at=?, reviewer_qq=?, review_comment='审核通过'
            WHERE task_id=? AND workspace=? AND qq=? AND status='submitted'
            """,
            (now_text(), now_text(), str(reviewer_qq), task_id, workspace, str(target_qq)),
        )
        if cur.rowcount != 1:
            row = conn.execute(
                "SELECT status FROM task_assignments WHERE task_id=? AND workspace=? AND qq=?",
                (task_id, workspace, str(target_qq)),
            ).fetchone()
            if not row:
                return "not_claimed"
            return "not_submitted"
        volunteer.on_approval(conn, workspace, task_id, target_qq, reviewer_qq, units)
        _sync_task_status(conn, task_id)
        return "ok"


def reject_submission(
    workspace: str,
    task_id: int,
    target_qq: str,
    reviewer_qq: str,
    comment: str,
) -> str:
    comment = str(comment).strip() or "需要继续修改"
    with connect() as conn:
        task = conn.execute(
            "SELECT id, deleted_at FROM tasks WHERE workspace=? AND id=?",
            (workspace, task_id),
        ).fetchone()
        if not task:
            return "not_found"
        if task["deleted_at"] is not None:
            return "deleted"
        cur = conn.execute(
            """
            UPDATE task_assignments
            SET status='claimed', completed_at=NULL, reviewed_at=?, reviewer_qq=?, review_comment=?
            WHERE task_id=? AND workspace=? AND qq=? AND status='submitted'
            """,
            (now_text(), str(reviewer_qq), comment, task_id, workspace, str(target_qq)),
        )
        if cur.rowcount != 1:
            row = conn.execute(
                "SELECT status FROM task_assignments WHERE task_id=? AND workspace=? AND qq=?",
                (task_id, workspace, str(target_qq)),
            ).fetchone()
            if not row:
                return "not_claimed"
            return "not_submitted"
        _sync_task_status(conn, task_id)
        return "ok"


def list_pending_reviews(workspace: str, limit: int = 30):
    with connect() as conn:
        rows = conn.execute(
            """
            SELECT a.task_id, a.qq, a.submission_text, a.submitted_at,
                   COALESCE(u.name, a.qq) AS name,
                   t.task_type, t.subject
            FROM task_assignments a
            JOIN tasks t ON t.id=a.task_id AND t.workspace=a.workspace
            LEFT JOIN users u ON u.workspace=a.workspace AND u.qq=a.qq
            WHERE a.workspace=? AND a.status='submitted' AND t.deleted_at IS NULL
            ORDER BY a.submitted_at ASC, a.task_id ASC
            LIMIT ?
            """,
            (workspace, limit),
        ).fetchall()
        return [dict(r) for r in rows]


def complete_task(workspace: str, task_id: int, requester_qq: str) -> str:
    """兼容旧调用：v5 不再允许成员绕过审核直接完成。"""
    row = get_assignment(workspace, task_id, requester_qq)
    if not row:
        return "not_claimed"
    if row["status"] == "completed":
        return "already_completed"
    return "review_required"



def reopen_task(workspace: str, task_id: int) -> bool:
    """已完成任务退回进行中，保留原参与成员，并把他们的份额全部改回未完成。"""
    with connect() as conn:
        task = conn.execute(
            """
            SELECT * FROM tasks
            WHERE workspace=? AND id=? AND deleted_at IS NULL
            """,
            (workspace, task_id),
        ).fetchone()
        if not task or task["status"] != "completed":
            return False
        conn.execute(
            """
            UPDATE task_assignments
            SET status='claimed', completed_at=NULL, submitted_at=NULL,
                reviewed_at=NULL, reviewer_qq=NULL, review_comment='任务被管理员重开'
            WHERE task_id=? AND workspace=?
            """,
            (task_id, workspace),
        )
        _sync_task_status(conn, task_id)
        return True


def reset_task_assignments(workspace: str, task_id: int) -> bool:
    """清空所有参与者，将任务重新开放为 0/N。"""
    with connect() as conn:
        task = conn.execute(
            "SELECT id FROM tasks WHERE workspace=? AND id=? AND deleted_at IS NULL",
            (workspace, task_id),
        ).fetchone()
        if not task:
            return False
        conn.execute(
            "DELETE FROM task_assignments WHERE task_id=? AND workspace=?",
            (task_id, workspace),
        )
        _sync_task_status(conn, task_id)
        return True


def update_task_field(workspace: str, task_id: int, field: str, value) -> str:
    allowed = {
        "task_type": "task_type",
        "subject": "subject",
        "required_knowledge": "required_knowledge",
        "required_people": "required_people",
    }
    if field not in allowed:
        raise ValueError("不可修改的字段")

    with connect() as conn:
        task = conn.execute(
            "SELECT * FROM tasks WHERE workspace=? AND id=? AND deleted_at IS NULL",
            (workspace, task_id),
        ).fetchone()
        if not task:
            return "not_found"

        if field == "required_people":
            try:
                n = int(value)
            except (TypeError, ValueError):
                return "invalid_people"
            if n < 1:
                return "invalid_people"
            assigned = conn.execute(
                "SELECT COUNT(*) AS n FROM task_assignments WHERE task_id=?",
                (task_id,),
            ).fetchone()["n"]
            if n < int(assigned):
                return "below_assigned"
            conn.execute(
                "UPDATE tasks SET required_people=? WHERE id=?",
                (n, task_id),
            )
            _sync_task_status(conn, task_id)
            return "ok"

        text = str(value).strip()
        if not text:
            return "empty"
        conn.execute(
            f"UPDATE tasks SET {allowed[field]}=? WHERE id=?",
            (text, task_id),
        )
        return "ok"


def soft_delete_task(workspace: str, task_id: int) -> str:
    with connect() as conn:
        row = conn.execute(
            "SELECT deleted_at FROM tasks WHERE workspace=? AND id=?",
            (workspace, task_id),
        ).fetchone()
        if not row:
            return "not_found"
        if row["deleted_at"] is not None:
            return "already_deleted"
        conn.execute(
            "UPDATE tasks SET deleted_at=? WHERE workspace=? AND id=?",
            (now_text(), workspace, task_id),
        )
        return "ok"


def restore_task(workspace: str, task_id: int) -> str:
    with connect() as conn:
        row = conn.execute(
            "SELECT deleted_at FROM tasks WHERE workspace=? AND id=?",
            (workspace, task_id),
        ).fetchone()
        if not row:
            return "not_found"
        if row["deleted_at"] is None:
            return "not_deleted"
        conn.execute(
            "UPDATE tasks SET deleted_at=NULL WHERE workspace=? AND id=?",
            (workspace, task_id),
        )
        _sync_task_status(conn, task_id)
        return "ok"



def log_action(workspace: str, actor_qq: str, action: str, detail: str = "") -> None:
    with connect() as conn:
        conn.execute(
            "INSERT INTO audit_logs(workspace, actor_qq, action, detail, created_at) VALUES (?, ?, ?, ?, ?)",
            (workspace, str(actor_qq), action.strip(), detail.strip(), now_text()),
        )


def list_audit_logs(workspace: str, limit: int = 20, include_volunteer: bool = True):
    limit = max(1, min(int(limit), 100))
    with connect() as conn:
        rows = conn.execute(
            """
            SELECT l.*, COALESCE(u.name, l.actor_qq) AS actor_name
            FROM audit_logs l
            LEFT JOIN users u ON u.workspace=l.workspace AND u.qq=l.actor_qq
            WHERE l.workspace=? AND (? OR (l.action NOT LIKE '%志愿时长%' AND l.action NOT LIKE '时长负责人%'))
            ORDER BY l.id DESC LIMIT ?
            """,
            (workspace, int(include_volunteer), limit),
        ).fetchall()
        return [dict(r) for r in rows]


def period_stats(workspace: str, start_text: str):
    with connect() as conn:
        tasks_created = conn.execute(
            "SELECT COUNT(*) AS n FROM tasks WHERE workspace=? AND created_at>=? AND deleted_at IS NULL",
            (workspace, start_text),
        ).fetchone()["n"]
        tasks_completed = conn.execute(
            "SELECT COUNT(*) AS n FROM tasks WHERE workspace=? AND completed_at>=? AND deleted_at IS NULL",
            (workspace, start_text),
        ).fetchone()["n"]
        assignments_claimed = conn.execute(
            "SELECT COUNT(*) AS n FROM task_assignments WHERE workspace=? AND claimed_at>=?",
            (workspace, start_text),
        ).fetchone()["n"]
        assignments_completed = conn.execute(
            "SELECT COUNT(*) AS n FROM task_assignments WHERE workspace=? AND completed_at>=?",
            (workspace, start_text),
        ).fetchone()["n"]
        participants = conn.execute(
            "SELECT COUNT(DISTINCT qq) AS n FROM task_assignments WHERE workspace=? AND claimed_at>=?",
            (workspace, start_text),
        ).fetchone()["n"]
        pending_reviews = conn.execute(
            "SELECT COUNT(*) AS n FROM task_assignments WHERE workspace=? AND status='submitted'",
            (workspace,),
        ).fetchone()["n"]
        active_tasks = conn.execute(
            "SELECT COUNT(*) AS n FROM tasks WHERE workspace=? AND deleted_at IS NULL AND status<>'completed'",
            (workspace,),
        ).fetchone()["n"]
        return {
            "tasks_created": int(tasks_created or 0),
            "tasks_completed": int(tasks_completed or 0),
            "assignments_claimed": int(assignments_claimed or 0),
            "assignments_completed": int(assignments_completed or 0),
            "participants": int(participants or 0),
            "pending_reviews": int(pending_reviews or 0),
            "active_tasks": int(active_tasks or 0),
        }


def member_stats(workspace: str, qq: str):
    qq = str(qq)
    with connect() as conn:
        counts = conn.execute(
            """
            SELECT COUNT(*) AS total,
                   SUM(CASE WHEN status='claimed' THEN 1 ELSE 0 END) AS claimed,
                   SUM(CASE WHEN status='submitted' THEN 1 ELSE 0 END) AS submitted,
                   SUM(CASE WHEN status='completed' THEN 1 ELSE 0 END) AS completed
            FROM task_assignments
            WHERE workspace=? AND qq=?
            """,
            (workspace, qq),
        ).fetchone()
        return {
            "total": int(counts["total"] or 0),
            "claimed": int(counts["claimed"] or 0),
            "submitted": int(counts["submitted"] or 0),
            "completed": int(counts["completed"] or 0),
        }


def export_data(workspace: str):
    with connect() as conn:
        users = [dict(r) for r in conn.execute(
            "SELECT id, qq, name, role, student_number, major_college, minor_college, registered_at, active FROM users WHERE workspace=? ORDER BY id",
            (workspace,),
        ).fetchall()]
        tasks = [dict(r) for r in conn.execute(
            """
            SELECT id, task_type, subject, required_knowledge, publisher_qq, status,
                   created_at, completed_at, required_people, deleted_at
            FROM tasks WHERE workspace=? ORDER BY id
            """,
            (workspace,),
        ).fetchall()]
        assignments = [dict(r) for r in conn.execute(
            """
            SELECT a.id, a.task_id, a.qq, COALESCE(u.name, a.qq) AS name, a.status,
                   a.progress_text, a.progress_updated_at,
                   a.claimed_at, a.submitted_at, a.completed_at, a.submission_text,
                   a.review_comment, a.reviewed_at, a.reviewer_qq
            FROM task_assignments a
            LEFT JOIN users u ON u.workspace=a.workspace AND u.qq=a.qq
            WHERE a.workspace=? ORDER BY a.id
            """,
            (workspace,),
        ).fetchall()]
        logs = [dict(r) for r in conn.execute(
            """
            SELECT l.id, l.actor_qq, COALESCE(u.name, l.actor_qq) AS actor_name,
                   l.action, l.detail, l.created_at
            FROM audit_logs l
            LEFT JOIN users u ON u.workspace=l.workspace AND u.qq=l.actor_qq
            WHERE l.workspace=? ORDER BY l.id
            """,
            (workspace,),
        ).fetchall()]
        requests = [dict(r) for r in conn.execute("SELECT * FROM task_requests WHERE workspace=? ORDER BY id", (workspace,))]
        return {"users": users, "tasks": tasks, "assignments": assignments, "audit_logs": logs, "task_requests": requests,
                "volunteer_entries": [dict(r) for r in conn.execute("SELECT * FROM volunteer_entries WHERE workspace=? ORDER BY id", (workspace,))],
                "volunteer_managers": [dict(r) for r in conn.execute("SELECT * FROM volunteer_managers WHERE workspace=? ORDER BY qq", (workspace,))],
                "volunteer_changes": [dict(r) for r in conn.execute("SELECT * FROM volunteer_changes WHERE workspace=? ORDER BY id", (workspace,))]}


def create_daily_backup(keep: int = 30):
    if not DB_PATH.exists():
        return None
    backup_dir = DB_PATH.with_name("backups")
    backup_dir.mkdir(exist_ok=True)
    today = datetime.now().strftime("%Y-%m-%d")
    target = backup_dir / f"notes_department_{today}.db"
    if not target.exists():
        src = sqlite3.connect(DB_PATH)
        dst = sqlite3.connect(target)
        try:
            src.backup(dst)
        finally:
            dst.close()
            src.close()
    files = sorted(backup_dir.glob("notes_department_*.db"), key=lambda x: x.name, reverse=True)
    for old in files[max(1, int(keep)):]:
        try:
            old.unlink()
        except OSError:
            pass
    return target


# Proposals become public tasks only after approval; no automatic assignment.
def migrate_task_requests():
    with connect() as conn:
        if _table_exists(conn, "task_requests"):
            return None
        backup = _backup_before_migration(conn, "requests_v8")
        conn.execute("BEGIN IMMEDIATE")
        conn.execute("""CREATE TABLE IF NOT EXISTS task_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT, workspace TEXT NOT NULL,
            applicant_qq TEXT NOT NULL, task_type TEXT NOT NULL, subject TEXT NOT NULL,
            required_knowledge TEXT NOT NULL,
            required_people INTEGER NOT NULL CHECK(required_people >= 1), reason TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'pending' CHECK(status IN ('pending','approved','rejected','withdrawn')),
            created_at TEXT NOT NULL, reviewed_at TEXT, reviewer_qq TEXT,
            review_comment TEXT NOT NULL DEFAULT '', task_id INTEGER UNIQUE REFERENCES tasks(id)
        )""")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_requests_workspace_status ON task_requests(workspace,status,id)")
        return backup


def validate_task_request_field(index, text):
    text = str(text).strip()
    if not text:
        raise ValueError("内容不能为空，请重新输入；发送“取消”可退出。")
    if index == 3:
        try:
            count = int(text)
        except ValueError:
            count = 0
        if not 1 <= count <= 100:
            raise ValueError("人数需求请输入 1 到 100 的整数。")
        return str(count)
    limit = (100, 100, 2000, 3, 1000)[index]
    if len(text) > limit:
        raise ValueError(f"这一项请控制在 {limit} 字以内。")
    return text


def create_task_request(workspace, qq, fields):
    if len(fields) != 5:
        raise ValueError("申请信息不完整。")
    values = [validate_task_request_field(i, v) for i, v in enumerate(fields)]
    with connect() as conn:
        conn.execute("BEGIN IMMEDIATE")
        if not conn.execute("SELECT 1 FROM users WHERE workspace=? AND qq=? AND active=1", (workspace,str(qq))).fetchone():
            raise ValueError("当前成员资格无效，不能申请任务。")
        cur = conn.execute("""INSERT INTO task_requests
            (workspace,applicant_qq,task_type,subject,required_knowledge,required_people,reason,created_at)
            VALUES(?,?,?,?,?,?,?,?)""", (workspace,str(qq),*values[:3],int(values[3]),values[4],now_text()))
        rid = int(cur.lastrowid)
        conn.execute("INSERT INTO audit_logs(workspace,actor_qq,action,detail,created_at) VALUES(?,?,?,?,?)",
                     (workspace,str(qq),"申请发布任务",f"申请 #{rid}",now_text()))
        return rid


def get_task_request(workspace, request_id):
    with connect() as conn:
        row = conn.execute("""SELECT r.*, COALESCE(u.name,r.applicant_qq) AS applicant_name
            FROM task_requests r LEFT JOIN users u ON u.workspace=r.workspace AND u.qq=r.applicant_qq
            WHERE r.workspace=? AND r.id=?""", (workspace,request_id)).fetchone()
        return dict(row) if row else None


def list_task_requests(workspace, applicant_qq=None, pending_only=False, page=1):
    where = "r.workspace=?"
    args = [workspace]
    if applicant_qq is not None:
        where += " AND r.applicant_qq=?"
        args.append(str(applicant_qq))
    if pending_only:
        where += " AND r.status='pending'"
    with connect() as conn:
        return [dict(row) for row in conn.execute(f"""SELECT r.*, COALESCE(u.name,r.applicant_qq) AS applicant_name
            FROM task_requests r LEFT JOIN users u ON u.workspace=r.workspace AND u.qq=r.applicant_qq
            WHERE {where} ORDER BY r.id DESC LIMIT 10 OFFSET ?""", (*args,(max(1,page)-1)*10))]


def decide_task_request(workspace, request_id, actor_qq, decision, comment=""):
    if decision not in ("approved","rejected","withdrawn"):
        raise ValueError("无效审批操作")
    comment = comment.strip()
    if decision == "rejected" and not comment:
        return "reason_required", None
    if len(comment) > 1000:
        return "reason_too_long", None
    # A single transaction serializes approvals/withdrawals and task creation.
    with connect() as conn:
        conn.execute("BEGIN IMMEDIATE")
        actor = conn.execute("SELECT role FROM users WHERE workspace=? AND qq=? AND active=1", (workspace,str(actor_qq))).fetchone()
        if not actor:
            return "forbidden", None
        row = conn.execute("SELECT * FROM task_requests WHERE workspace=? AND id=?", (workspace,request_id)).fetchone()
        if not row:
            return "not_found", None
        allowed = row['applicant_qq'] == str(actor_qq) if decision == 'withdrawn' else (actor['role'] in ('owner','admin') or bool(conn.execute('SELECT 1 FROM volunteer_managers WHERE workspace=? AND qq=?',(workspace,str(actor_qq))).fetchone()))
        if not allowed:
            return "forbidden", None
        if row['status'] != 'pending':
            return "already_processed", row['task_id']
        stamp = now_text()
        task_id = None
        if decision == 'approved':
            cur = conn.execute("""INSERT INTO tasks
                (workspace,task_type,subject,required_knowledge,publisher_qq,status,created_at,required_people,deleted_at)
                VALUES(?,?,?,?,?,'open',?,?,NULL)""",
                (workspace,row['task_type'],row['subject'],row['required_knowledge'],str(actor_qq),stamp,row['required_people']))
            task_id = int(cur.lastrowid)
        conn.execute("UPDATE task_requests SET status=?,reviewed_at=?,reviewer_qq=?,review_comment=?,task_id=? WHERE workspace=? AND id=?",
                     (decision,stamp,str(actor_qq),comment,task_id,workspace,request_id))
        action = {'approved':'批准任务申请','rejected':'驳回任务申请','withdrawn':'撤回任务申请'}[decision]
        conn.execute("INSERT INTO audit_logs(workspace,actor_qq,action,detail,created_at) VALUES(?,?,?,?,?)",
                     (workspace,str(actor_qq),action,f"申请 #{request_id}" + (f" → 任务 #{task_id}" if task_id else "") + (f"｜{comment}" if comment else ""),stamp))
        return "ok", task_id
