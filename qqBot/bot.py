from __future__ import annotations

import asyncio
import csv
import inspect
import json
import re
import uuid
import zipfile
from datetime import datetime, timedelta
from dataclasses import dataclass
from pathlib import Path
from typing import Dict

from openpyxl import Workbook

import websockets

import db

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"

ROLE_LABEL = {
    "owner": "Owner（负责人）",
    "admin": "Admin（管理员）",
    "member": "Member（工作成员）",
}

PRIVATE_HELP = """笔记部 Bot（私聊）

【成员】
注册 <姓名>
我的身份
任务
任务 <编号>
认领 <编号>
我的任务
我的历史
提交 <编号>        # 完成 <编号> 也可
撤回提交 <编号>
取消认领 <编号>

【负责人 / 管理员】
发布任务
所有任务
待审核
通过 <编号> <姓名或QQ>
打回 <编号> <姓名或QQ> <原因>
修改任务 <编号>
重开 <编号>
重新开放 <编号>
删除任务 <编号>
恢复任务 <编号>
已删除任务
移除认领 <编号> <姓名或QQ>
成员列表
管理员列表
操作日志 [条数]
本周统计
本月统计
成员统计 <姓名或QQ>
导出数据
导出CSV

【仅负责人】
管理员添加 <姓名或QQ>
管理员移除 <姓名或QQ>
成员移除 <姓名或QQ>
成员恢复 <姓名或QQ>
停用成员

提交任务后进入“待审核”，只有负责人/管理员审核通过后才计为完成。
发布任务采用：任务类型 → 涉及学科 → 所需知识 → 人数需求 → 确认。
“删除任务”为软删除，可以用“恢复任务 <编号>”找回。"""

GROUP_HELP = "为了避免刷屏，主要操作请私聊我。\n群内只需：@我 任务（查看仍有空位的任务）。"


@dataclass
class PublishSession:
    stage: str = "task_type"
    task_type: str = ""
    subject: str = ""
    required_knowledge: str = ""
    required_people: int = 1


@dataclass
class EditSession:
    task_id: int
    stage: str = "choose_field"
    field: str = ""


publish_sessions: Dict[str, PublishSession] = {}
edit_sessions: Dict[str, EditSession] = {}
delete_sessions: Dict[str, int] = {}
submit_sessions: Dict[str, int] = {}


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        raise RuntimeError("缺少 config.json。请复制 config.example.json 为 config.json 并填写。")
    data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    required = ["ws_url", "department_group_id", "owner_qq", "owner_name"]
    missing = [k for k in required if not str(data.get(k, "")).strip()]
    if missing:
        raise RuntimeError(f"config.json 缺少配置：{', '.join(missing)}")
    data["department_group_id"] = str(data["department_group_id"])
    data["owner_qq"] = str(data["owner_qq"])
    # OneBot Token 为可选项：留空可兼容旧配置；正式使用建议启用。
    data["onebot_token"] = str(data.get("onebot_token", "")).strip()
    return data


def clean_text(text: str) -> str:
    text = text or ""
    text = re.sub(r"\[CQ:at,qq=\d+[^\]]*\]", "", text)
    text = text.strip()
    if text.startswith("/"):
        text = text[1:].strip()
    return re.sub(r"\s+", " ", text)


def split_command(text: str):
    text = clean_text(text)
    if not text:
        return "", ""
    parts = text.split(" ", 1)
    return parts[0], parts[1].strip() if len(parts) > 1 else ""


def is_admin(user) -> bool:
    return bool(user and user["role"] in {"owner", "admin"})


def is_owner(user) -> bool:
    return bool(user and user["role"] == "owner")


def parse_task_id(arg: str):
    try:
        n = int(arg.strip())
        return n if n > 0 else None
    except (TypeError, ValueError):
        return None


def task_state_label(task: dict) -> str:
    if task.get("deleted_at"):
        return "已删除"
    if task["status"] == "completed":
        return "已完成"
    assigned = int(task.get("assigned_count", 0))
    required = int(task.get("required_people", 1))
    if assigned == 0:
        return "待认领"
    if assigned < required:
        return "招募中"
    return "已满员 / 进行中"


def assignee_text(task: dict) -> str:
    people = task.get("assignees", [])
    if not people:
        return "—"
    label = {"claimed": "进行中", "submitted": "待审核", "completed": "✓"}
    return "、".join(f"{p['name']}（{label.get(p['status'], p['status'])}）" for p in people)


def task_summary(task: dict) -> str:
    lines = [
        f"任务 #{task['id']}｜{task_state_label(task)}",
        f"任务类型：{task['task_type']}",
        f"涉及学科：{task['subject']}",
        f"所需知识：{task['required_knowledge']}",
        f"人数需求：{task['required_people']} 人",
        f"认领进度：{task['assigned_count']}/{task['required_people']}",
        f"待审核：{task.get('submitted_count', 0)} 人",
        f"完成进度：{task['completed_count']}/{task['required_people']}",
    ]
    if task.get("publisher_name"):
        lines.append(f"发布人：{task['publisher_name']}")
    lines.append(f"参与成员：{assignee_text(task)}")
    if task.get("deleted_at"):
        lines.append(f"删除时间：{task['deleted_at']}")
    return "\n".join(lines)


def _write_sheet(ws, rows: list[dict]):
    if not rows:
        ws.append(["暂无数据"])
        return
    headers = list(rows[0].keys())
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h) for h in headers])


def make_export_xlsx(workspace: str) -> Path:
    data = db.export_data(workspace)
    export_dir = BASE_DIR / "exports"
    export_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = export_dir / f"notes_department_{stamp}.xlsx"
    wb = Workbook()
    default = wb.active
    wb.remove(default)
    mapping = [
        ("成员", data["users"]),
        ("任务", data["tasks"]),
        ("认领与提交", data["assignments"]),
        ("操作日志", data["audit_logs"]),
    ]
    for title, rows in mapping:
        ws = wb.create_sheet(title)
        _write_sheet(ws, rows)
    wb.save(path)
    return path


def make_export_csv_zip(workspace: str) -> Path:
    data = db.export_data(workspace)
    export_dir = BASE_DIR / "exports"
    export_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_path = export_dir / f"notes_department_csv_{stamp}.zip"
    csv_dir = export_dir / f"csv_{stamp}"
    csv_dir.mkdir(exist_ok=True)
    try:
        for name, rows in data.items():
            path = csv_dir / f"{name}.csv"
            headers = list(rows[0].keys()) if rows else ["暂无数据"]
            with path.open("w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=headers)
                writer.writeheader()
                for row in rows:
                    writer.writerow(row)
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in csv_dir.glob("*.csv"):
                zf.write(path, arcname=path.name)
    finally:
        for path in csv_dir.glob("*.csv"):
            path.unlink(missing_ok=True)
        try:
            csv_dir.rmdir()
        except OSError:
            pass
    return zip_path


def _period_start(kind: str) -> datetime:
    now = datetime.now()
    if kind == "week":
        return (now - timedelta(days=now.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)
    return now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)


def event_at_bot(event: dict) -> bool:
    self_id = str(event.get("self_id", ""))
    message = event.get("message")
    if isinstance(message, list):
        for seg in message:
            if seg.get("type") == "at" and str(seg.get("data", {}).get("qq", "")) == self_id:
                return True
    raw = event.get("raw_message", "") or ""
    return f"[CQ:at,qq={self_id}" in raw


class OneBotClient:
    def __init__(self, ws_url: str, handler, access_token: str = ""):
        self.ws_url = ws_url
        self.handler = handler
        self.access_token = str(access_token or "").strip()
        self.ws = None
        self.pending: Dict[str, asyncio.Future] = {}
        self.send_lock = asyncio.Lock()

    def _connect_kwargs(self) -> dict:
        """构造 WebSocket 连接参数，并兼容 websockets 12~15 的请求头参数名。"""
        kwargs = {"max_size": 8 * 1024 * 1024}
        if not self.access_token:
            return kwargs
        headers = {"Authorization": f"Bearer {self.access_token}"}
        params = inspect.signature(websockets.connect).parameters
        if "additional_headers" in params:
            kwargs["additional_headers"] = headers
        else:
            # websockets 12/13 使用 extra_headers。
            kwargs["extra_headers"] = headers
        return kwargs

    async def run_forever(self):
        while True:
            try:
                print(f"[NapCat] 正在连接 {self.ws_url} ...")
                async with websockets.connect(self.ws_url, **self._connect_kwargs()) as ws:
                    self.ws = ws
                    print("[NapCat] OneBot WebSocket 已连接。")
                    await self.recv_loop(ws)
            except asyncio.CancelledError:
                raise
            except Exception as e:
                print(f"[NapCat] 连接中断：{type(e).__name__}: {e}")
            finally:
                self.ws = None
                for fut in self.pending.values():
                    if not fut.done():
                        fut.set_exception(ConnectionError("NapCat WebSocket 已断开"))
                self.pending.clear()
            print("[NapCat] 3 秒后重连...")
            await asyncio.sleep(3)

    async def recv_loop(self, ws):
        async for raw in ws:
            try:
                data = json.loads(raw)
            except json.JSONDecodeError:
                continue
            echo = data.get("echo")
            if echo and echo in self.pending:
                fut = self.pending.pop(echo)
                if not fut.done():
                    fut.set_result(data)
                continue
            asyncio.create_task(self.handler(data))

    async def call_api(self, action: str, params: dict | None = None, timeout: float = 10):
        if not self.ws:
            raise ConnectionError("NapCat 尚未连接")
        echo = uuid.uuid4().hex
        loop = asyncio.get_running_loop()
        fut = loop.create_future()
        self.pending[echo] = fut
        payload = {"action": action, "params": params or {}, "echo": echo}
        try:
            async with self.send_lock:
                await self.ws.send(json.dumps(payload, ensure_ascii=False))
            return await asyncio.wait_for(fut, timeout=timeout)
        finally:
            self.pending.pop(echo, None)

    async def send_private(self, qq: str, text: str):
        return await self.call_api("send_private_msg", {"user_id": int(qq), "message": text})

    async def send_group(self, group_id: str, text: str):
        return await self.call_api("send_group_msg", {"group_id": int(group_id), "message": text})

    async def send_private_file(self, qq: str, path: Path):
        return await self.call_api(
            "upload_private_file",
            {"user_id": int(qq), "file": str(path.resolve()), "name": path.name},
            timeout=30,
        )


class NotesDepartmentBot:
    def __init__(self, config: dict):
        self.config = config
        self.workspace = config["department_group_id"]
        self.owner_qq = config["owner_qq"]
        self.owner_name = config["owner_name"]
        self.client: OneBotClient | None = None

    async def start(self):
        db.init_db()
        db.ensure_owner(self.workspace, self.owner_qq, self.owner_name)
        backup = db.create_daily_backup()
        if backup:
            print(f"[备份] 今日数据库备份：{backup}")
        self.client = OneBotClient(
            self.config["ws_url"],
            self.handle_event,
            access_token=self.config.get("onebot_token", ""),
        )
        print(f"[Bot] 工作群：{self.workspace}")
        print(f"[Bot] Owner：{self.owner_name} ({self.owner_qq})")
        print("[安全] OneBot Token：" + ("已启用" if self.config.get("onebot_token") else "未启用"))
        backup_task = asyncio.create_task(self.backup_loop())
        try:
            await self.client.run_forever()
        finally:
            backup_task.cancel()

    async def backup_loop(self):
        while True:
            await asyncio.sleep(3600)
            try:
                db.create_daily_backup()
            except Exception as e:
                print(f"[备份] 自动备份失败：{e}")

    async def announce(self, text: str):
        if not self.client:
            return
        try:
            await self.client.send_group(self.workspace, text)
        except Exception as e:
            print(f"[群公告] 发送失败：{e}")

    async def handle_event(self, event: dict):
        post_type = event.get("post_type")
        if post_type == "message":
            if event.get("message_type") == "private":
                await self.handle_private(event)
            elif event.get("message_type") == "group":
                await self.handle_group(event)
            return

        if post_type == "request" and event.get("request_type") == "friend":
            await self.handle_friend_request(event)
            return

        if post_type == "notice" and event.get("notice_type") == "group_decrease":
            if str(event.get("group_id")) == self.workspace:
                qq = str(event.get("user_id"))
                if qq != self.owner_qq:
                    db.deactivate_user(self.workspace, qq)
                    print(f"[成员] {qq} 已离开工作群，账户已停用。")

    async def api_ok(self, response: dict) -> bool:
        return response.get("status") == "ok" or response.get("retcode") == 0

    async def is_group_member(self, qq: str) -> bool:
        assert self.client
        try:
            resp = await self.client.call_api(
                "get_group_member_info",
                {"group_id": int(self.workspace), "user_id": int(qq), "no_cache": True},
            )
            return await self.api_ok(resp) and bool(resp.get("data"))
        except Exception:
            return False

    async def handle_friend_request(self, event: dict):
        if not self.config.get("auto_approve_group_members", True):
            return
        qq = str(event.get("user_id", ""))
        flag = event.get("flag")
        if not qq or not flag:
            return
        assert self.client
        member = await self.is_group_member(qq)
        resp = await self.client.call_api(
            "set_friend_add_request",
            {"flag": flag, "approve": bool(member), "remark": "笔记部成员" if member else ""},
        )
        print(f"[好友申请] QQ={qq} group_member={member} result={resp.get('status', resp.get('retcode'))}")

    async def handle_group(self, event: dict):
        if str(event.get("group_id")) != self.workspace:
            return
        if not event_at_bot(event):
            return
        raw = event.get("raw_message", "") or ""
        cmd, _ = split_command(raw)
        assert self.client

        if cmd in {"帮助", "help"}:
            await self.client.send_group(self.workspace, GROUP_HELP)
            return

        if cmd == "任务":
            rows = db.list_open_tasks(self.workspace, limit=10)
            if not rows:
                text = "当前没有可继续认领的任务。"
            else:
                chunks = ["当前可认领任务："]
                for task in rows:
                    chunks.append(
                        f"#{task['id']}｜{task['task_type']}｜{task['subject']}｜"
                        f"人数 {task['assigned_count']}/{task['required_people']}\n"
                        f"所需知识：{task['required_knowledge']}"
                    )
                chunks.append("\n需要认领时请私聊我：认领 <编号>")
                text = "\n\n".join(chunks)
            await self.client.send_group(self.workspace, text)
            return

        await self.client.send_group(self.workspace, "为避免刷屏，请私聊我进行注册、认领、发布和任务管理。")

    async def handle_private(self, event: dict):
        qq = str(event.get("user_id"))
        raw = event.get("raw_message", "") or ""
        text = clean_text(raw)
        cmd, arg = split_command(raw)
        assert self.client

        if qq == self.owner_qq:
            db.ensure_owner(self.workspace, self.owner_qq, self.owner_name)

        # 会话状态优先于普通命令。
        if qq in publish_sessions:
            await self.handle_publish_session(qq, text)
            return
        if qq in edit_sessions:
            await self.handle_edit_session(qq, text)
            return
        if qq in delete_sessions:
            await self.handle_delete_session(qq, text)
            return
        if qq in submit_sessions:
            await self.handle_submit_session(qq, text)
            return

        if cmd in {"帮助", "help", "菜单"}:
            await self.client.send_private(qq, PRIVATE_HELP)
            return

        if cmd == "注册":
            if not arg:
                await self.client.send_private(qq, "用法：注册 <姓名>\n例如：注册 张三")
                return
            if not await self.is_group_member(qq):
                await self.client.send_private(qq, "注册失败：你目前不在配置的笔记部工作群中。")
                return
            user = db.register_user(self.workspace, qq, arg, self.owner_qq)
            await self.client.send_private(qq, f"注册成功。\n姓名：{user['name']}\n身份：{ROLE_LABEL[user['role']]}")
            return

        user = db.get_user(self.workspace, qq)
        if not user:
            inactive = db.get_user(self.workspace, qq, include_inactive=True)
            if inactive and not inactive["active"]:
                await self.client.send_private(qq, "你的成员资格目前处于停用状态，请联系负责人。")
            else:
                await self.client.send_private(qq, "你还没有注册。请发送：注册 <姓名>")
            return

        if cmd == "我的身份":
            await self.client.send_private(qq, f"姓名：{user['name']}\n身份：{ROLE_LABEL[user['role']]}\nQQ：{qq}")
            return

        if cmd == "任务":
            if arg:
                tid = parse_task_id(arg)
                if tid is None:
                    await self.client.send_private(qq, "用法：任务 <编号>")
                    return
                task = db.get_task(self.workspace, tid)
                await self.client.send_private(qq, task_summary(task) if task else "没有找到这个任务。")
            else:
                rows = db.list_open_tasks(self.workspace)
                if not rows:
                    await self.client.send_private(qq, "当前没有可继续认领的任务。")
                else:
                    chunks = ["当前可认领任务："]
                    for task in rows:
                        chunks.append(
                            f"#{task['id']}｜{task['task_type']}｜{task['subject']}｜"
                            f"人数 {task['assigned_count']}/{task['required_people']}\n"
                            f"所需知识：{task['required_knowledge']}"
                        )
                    chunks.append("\n认领：认领 <编号>")
                    await self.client.send_private(qq, "\n\n".join(chunks))
            return

        if cmd == "认领":
            tid = parse_task_id(arg)
            if tid is None:
                await self.client.send_private(qq, "用法：认领 <编号>")
                return
            result = db.claim_task(self.workspace, tid, qq)
            if result == "ok":
                task = db.get_task(self.workspace, tid)
                await self.client.send_private(qq, f"认领成功。\n\n{task_summary(task)}")
                if task and task["assigned_count"] >= task["required_people"]:
                    await self.announce(
                        f"【任务 #{tid} 已满员】\n{task['task_type']}｜{task['subject']}\n"
                        f"参与成员：{'、'.join(p['name'] for p in task['assignees'])}"
                    )
            else:
                messages = {
                    "not_found": "任务不存在。",
                    "deleted": "这个任务已经被删除。",
                    "completed": f"任务 #{tid} 已完成。",
                    "already_claimed": f"你已经认领过任务 #{tid}。",
                    "full": f"任务 #{tid} 的认领人数已经满了。",
                }
                await self.client.send_private(qq, messages.get(result, "认领失败，请稍后再试。"))
            return

        if cmd == "我的任务":
            rows = db.list_my_active_tasks(self.workspace, qq)
            if not rows:
                await self.client.send_private(qq, "你目前没有进行中的任务。")
            else:
                chunks = ["你当前进行中的任务："]
                for task in rows:
                    chunks.append(
                        f"#{task['id']}｜{task['task_type']}｜{task['subject']}｜"
                        f"人数 {task['assigned_count']}/{task['required_people']}\n"
                        f"所需知识：{task['required_knowledge']}"
                    )
                await self.client.send_private(qq, "\n\n".join(chunks))
            return

        if cmd == "我的历史":
            rows = db.list_my_completed_tasks(self.workspace, qq)
            if not rows:
                await self.client.send_private(qq, "你目前还没有已完成任务。")
            else:
                lines = ["你最近完成的任务："]
                for task in rows:
                    mine = next((p for p in task["assignees"] if p["qq"] == qq), None)
                    when = mine["completed_at"] if mine else ""
                    lines.append(f"#{task['id']}｜{task['task_type']}｜{task['subject']}｜{when}")
                await self.client.send_private(qq, "\n".join(lines))
            return

        if cmd in {"提交", "完成"}:
            tid = parse_task_id(arg)
            if tid is None:
                await self.client.send_private(qq, "用法：提交 <编号>\n例如：提交 12")
                return
            assignment = db.get_assignment(self.workspace, tid, qq)
            if not assignment:
                await self.client.send_private(qq, "你没有认领这个任务。")
                return
            if assignment["status"] == "submitted":
                await self.client.send_private(qq, "这个任务份额已经提交，正在等待审核。")
                return
            if assignment["status"] == "completed":
                await self.client.send_private(qq, "这个任务份额已经审核完成。")
                return
            submit_sessions[qq] = tid
            await self.client.send_private(
                qq,
                f"准备提交任务 #{tid}。\n请发送【成果说明或链接】。\n"
                "可以填写 GitHub/网盘链接、群文件名，或简单说明成果存放位置。\n"
                "发送“取消”放弃本次提交。",
            )
            return

        if cmd == "撤回提交":
            tid = parse_task_id(arg)
            if tid is None:
                await self.client.send_private(qq, "用法：撤回提交 <编号>")
                return
            result = db.withdraw_submission(self.workspace, tid, qq)
            messages = {
                "ok": f"任务 #{tid} 的提交已撤回，你可以继续修改后重新提交。",
                "not_submitted": "这个任务当前不在待审核状态。",
                "not_claimed": "你没有认领这个任务。",
            }
            await self.client.send_private(qq, messages.get(result, "撤回失败。"))
            return

        if cmd == "取消认领":
            tid = parse_task_id(arg)
            if tid is None:
                await self.client.send_private(qq, "用法：取消认领 <编号>")
                return
            result = db.release_task(self.workspace, tid, qq)
            messages = {
                "ok": f"已取消你对任务 #{tid} 的认领，空位已重新开放。",
                "not_found": "任务不存在。",
                "deleted": "这个任务已经被删除。",
                "not_claimed": "你没有该任务的未完成认领。",
            }
            await self.client.send_private(qq, messages.get(result, "操作失败。"))
            return

        if cmd == "待审核":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以查看待审核提交。")
                return
            rows = db.list_pending_reviews(self.workspace)
            if not rows:
                await self.client.send_private(qq, "当前没有待审核提交。")
            else:
                chunks = ["待审核提交："]
                for r in rows:
                    content = (r.get("submission_text") or "").replace("\n", " ")
                    if len(content) > 160:
                        content = content[:157] + "..."
                    chunks.append(
                        f"#{r['task_id']}｜{r['name']}｜{r['task_type']}｜{r['subject']}\n"
                        f"提交：{content}\n时间：{r['submitted_at']}"
                    )
                chunks.append("\n审核：通过 <编号> <姓名或QQ>\n打回：打回 <编号> <姓名或QQ> <原因>")
                await self.client.send_private(qq, "\n\n".join(chunks))
            return

        if cmd in {"通过", "打回"}:
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以审核任务。")
                return
            parts = arg.split(" ", 2)
            if (cmd == "通过" and len(parts) < 2) or (cmd == "打回" and len(parts) < 3):
                usage = "通过 <任务编号> <姓名或QQ>" if cmd == "通过" else "打回 <任务编号> <姓名或QQ> <原因>"
                await self.client.send_private(qq, f"用法：{usage}")
                return
            tid = parse_task_id(parts[0])
            if tid is None:
                await self.client.send_private(qq, "任务编号无效。")
                return
            status, target = db.resolve_user(self.workspace, parts[1], active=None)
            if status == "not_found":
                await self.client.send_private(qq, "没有找到对应成员。")
                return
            if status == "ambiguous":
                await self.client.send_private(qq, "存在同名成员，请改用 QQ 号：\n" + "\n".join(f"{r['name']}：{r['qq']}" for r in target))
                return
            if cmd == "通过":
                result = db.approve_submission(self.workspace, tid, target["qq"], qq)
                if result == "ok":
                    db.log_action(self.workspace, qq, "审核通过", f"任务 #{tid}｜成员 {target['name']} ({target['qq']})")
                    await self.client.send_private(qq, f"已通过 {target['name']} 在任务 #{tid} 的提交。")
                    try:
                        await self.client.send_private(target["qq"], f"你在任务 #{tid} 的提交已审核通过。")
                    except Exception:
                        pass
                    task = db.get_task(self.workspace, tid)
                    if task and task["status"] == "completed":
                        await self.announce(
                            f"【任务 #{tid} 已完成】\n{task['task_type']}｜{task['subject']}\n"
                            f"参与成员：{'、'.join(p['name'] for p in task['assignees'])}"
                        )
                else:
                    messages = {"not_found": "任务不存在。", "deleted": "任务已删除。", "not_claimed": "该成员没有认领此任务。", "not_submitted": "该成员当前没有待审核提交。"}
                    await self.client.send_private(qq, messages.get(result, "审核失败。"))
            else:
                comment = parts[2].strip()
                result = db.reject_submission(self.workspace, tid, target["qq"], qq, comment)
                if result == "ok":
                    db.log_action(self.workspace, qq, "审核打回", f"任务 #{tid}｜成员 {target['name']} ({target['qq']})｜原因：{comment}")
                    await self.client.send_private(qq, f"已打回 {target['name']} 在任务 #{tid} 的提交。")
                    try:
                        await self.client.send_private(target["qq"], f"你在任务 #{tid} 的提交被打回。\n原因：{comment}\n修改后请重新发送：提交 {tid}")
                    except Exception:
                        pass
                else:
                    messages = {"not_found": "任务不存在。", "deleted": "任务已删除。", "not_claimed": "该成员没有认领此任务。", "not_submitted": "该成员当前没有待审核提交。"}
                    await self.client.send_private(qq, messages.get(result, "打回失败。"))
            return

        if cmd == "操作日志":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以查看操作日志。")
                return
            try:
                limit = int(arg) if arg else 20
            except ValueError:
                limit = 20
            rows = db.list_audit_logs(self.workspace, limit)
            if not rows:
                await self.client.send_private(qq, "当前还没有管理员操作日志。")
            else:
                lines = ["管理员操作日志："]
                for r in rows:
                    detail = f"｜{r['detail']}" if r.get('detail') else ""
                    lines.append(f"{r['created_at']}｜{r['actor_name']}｜{r['action']}{detail}")
                await self.client.send_private(qq, "\n".join(lines))
            return

        if cmd in {"本周统计", "本月统计"}:
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以查看部门统计。")
                return
            kind = "week" if cmd == "本周统计" else "month"
            start_dt = _period_start(kind)
            st = db.period_stats(self.workspace, start_dt.strftime("%Y-%m-%d %H:%M:%S"))
            label = "本周" if kind == "week" else "本月"
            await self.client.send_private(
                qq,
                f"【{label}统计】\n"
                f"新增任务：{st['tasks_created']}\n"
                f"整体完成任务：{st['tasks_completed']}\n"
                f"新增认领份额：{st['assignments_claimed']}\n"
                f"审核完成份额：{st['assignments_completed']}\n"
                f"参与成员：{st['participants']}\n"
                f"当前待审核：{st['pending_reviews']}\n"
                f"当前进行中任务：{st['active_tasks']}"
            )
            return

        if cmd == "成员统计":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以查看成员统计。")
                return
            if not arg:
                await self.client.send_private(qq, "用法：成员统计 <姓名或QQ>")
                return
            status, target = db.resolve_user(self.workspace, arg, active=None)
            if status == "not_found":
                await self.client.send_private(qq, "没有找到对应成员。")
                return
            if status == "ambiguous":
                await self.client.send_private(qq, "存在同名成员，请改用 QQ 号。")
                return
            st = db.member_stats(self.workspace, target["qq"])
            await self.client.send_private(
                qq,
                f"【成员统计】{target['name']}\n"
                f"累计参与：{st['total']}\n"
                f"进行中：{st['claimed']}\n"
                f"待审核：{st['submitted']}\n"
                f"已完成：{st['completed']}"
            )
            return

        if cmd in {"导出数据", "导出CSV"}:
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以导出数据。")
                return
            try:
                path = make_export_xlsx(self.workspace) if cmd == "导出数据" else make_export_csv_zip(self.workspace)
                resp = await self.client.send_private_file(qq, path)
                if not await self.api_ok(resp):
                    await self.client.send_private(qq, f"导出已生成，但文件发送失败。文件保存在：{path}")
                else:
                    await self.client.send_private(qq, f"导出完成：{path.name}")
            except Exception as e:
                await self.client.send_private(qq, f"导出失败：{type(e).__name__}: {e}")
            return

        if cmd == "发布任务":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以发布任务。")
                return
            publish_sessions[qq] = PublishSession()
            await self.client.send_private(qq, "开始发布任务。\n第 1/4 步：请输入【任务类型】。\n发送“取消”可终止。")
            return

        if cmd == "修改任务":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以修改任务。")
                return
            tid = parse_task_id(arg)
            if tid is None:
                await self.client.send_private(qq, "用法：修改任务 <编号>")
                return
            task = db.get_task(self.workspace, tid)
            if not task:
                await self.client.send_private(qq, "没有找到这个任务，或任务已被删除。")
                return
            edit_sessions[qq] = EditSession(task_id=tid)
            await self.client.send_private(
                qq,
                f"当前任务：\n\n{task_summary(task)}\n\n"
                "要修改哪一项？发送：类型 / 学科 / 知识 / 人数\n发送“取消”退出修改。",
            )
            return

        if cmd == "删除任务":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以删除任务。")
                return
            tid = parse_task_id(arg)
            if tid is None:
                await self.client.send_private(qq, "用法：删除任务 <编号>")
                return
            task = db.get_task(self.workspace, tid)
            if not task:
                await self.client.send_private(qq, "没有找到这个任务，或它已经被删除。")
                return
            delete_sessions[qq] = tid
            await self.client.send_private(
                qq,
                f"你准备删除：\n\n{task_summary(task)}\n\n"
                "这是软删除，之后仍可恢复。\n发送“确认”删除；发送“取消”放弃。",
            )
            return

        if cmd == "恢复任务":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以恢复任务。")
                return
            tid = parse_task_id(arg)
            if tid is None:
                await self.client.send_private(qq, "用法：恢复任务 <编号>")
                return
            result = db.restore_task(self.workspace, tid)
            if result == "ok":
                db.log_action(self.workspace, qq, "恢复任务", f"任务 #{tid}")
                await self.announce(f"【任务 #{tid} 已恢复】成员可重新查看并认领空位。")
            messages = {
                "ok": f"任务 #{tid} 已恢复。",
                "not_found": "任务不存在。",
                "not_deleted": "这个任务当前并未被删除。",
            }
            await self.client.send_private(qq, messages.get(result, "恢复失败。"))
            return

        if cmd == "已删除任务":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以查看已删除任务。")
                return
            rows = db.list_deleted_tasks(self.workspace)
            if not rows:
                await self.client.send_private(qq, "当前没有已删除任务。")
            else:
                lines = ["已删除任务："]
                for task in rows:
                    lines.append(f"#{task['id']}｜{task['task_type']}｜{task['subject']}｜删除于 {task['deleted_at']}")
                lines.append("\n恢复：恢复任务 <编号>")
                await self.client.send_private(qq, "\n".join(lines))
            return

        if cmd == "重开":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以重新打开已完成任务。")
                return
            tid = parse_task_id(arg)
            if tid is None:
                await self.client.send_private(qq, "用法：重开 <编号>")
                return
            ok = db.reopen_task(self.workspace, tid)
            if ok:
                db.log_action(self.workspace, qq, "重开任务", f"任务 #{tid}｜保留原参与成员")
                await self.announce(f"【任务 #{tid} 已重开】原参与成员需继续修改并重新提交。")
            await self.client.send_private(
                qq,
                f"任务 #{tid} 已重开：原参与成员全部保留，并恢复为未完成状态。"
                if ok else "操作失败：任务不存在、已删除，或当前不是已完成状态。",
            )
            return

        if cmd == "重新开放":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以重新开放任务。")
                return
            tid = parse_task_id(arg)
            if tid is None:
                await self.client.send_private(qq, "用法：重新开放 <编号>")
                return
            ok = db.reset_task_assignments(self.workspace, tid)
            if ok:
                db.log_action(self.workspace, qq, "重新开放任务", f"任务 #{tid}｜清空全部参与成员")
                await self.announce(f"【任务 #{tid} 重新开放】认领记录已清空，可重新认领。")
            await self.client.send_private(
                qq,
                f"任务 #{tid} 已清空全部参与成员，并重新从 0 人开始认领。"
                if ok else "操作失败：任务不存在或已删除。",
            )
            return

        if cmd == "移除认领":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以移除他人的认领。")
                return
            parts = arg.split(" ", 1)
            if len(parts) != 2:
                await self.client.send_private(qq, "用法：移除认领 <任务编号> <姓名或QQ>")
                return
            tid = parse_task_id(parts[0])
            if tid is None:
                await self.client.send_private(qq, "任务编号无效。")
                return
            status, target = db.resolve_user(self.workspace, parts[1], active=None)
            if status == "not_found":
                await self.client.send_private(qq, "没有找到对应成员。")
                return
            if status == "ambiguous":
                await self.client.send_private(qq, "存在同名成员，请改用 QQ 号：\n" + "\n".join(f"{r['name']}：{r['qq']}" for r in target))
                return
            result = db.remove_assignment(self.workspace, tid, target["qq"])
            if result == "ok":
                db.log_action(self.workspace, qq, "移除认领", f"任务 #{tid}｜成员 {target['name']} ({target['qq']})")
            messages = {
                "ok": f"已从任务 #{tid} 中移除 {target['name']} 的未完成认领。",
                "not_found": "任务不存在。",
                "deleted": "任务已删除。",
                "not_claimed": f"{target['name']} 当前没有该任务的未完成认领。",
            }
            await self.client.send_private(qq, messages.get(result, "操作失败。"))
            return

        if cmd == "所有任务":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以查看全部任务。")
                return
            rows = db.list_all_tasks(self.workspace)
            if not rows:
                await self.client.send_private(qq, "当前还没有任务。")
            else:
                lines = ["最近任务："]
                for task in rows:
                    names = "、".join(p["name"] for p in task["assignees"]) or "—"
                    lines.append(
                        f"#{task['id']}｜{task_state_label(task)}｜{task['task_type']}｜{task['subject']}｜"
                        f"人数 {task['assigned_count']}/{task['required_people']}｜参与：{names}"
                    )
                await self.client.send_private(qq, "\n".join(lines))
            return

        if cmd == "成员列表":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以查看成员列表。")
                return
            rows = db.list_users(self.workspace)
            lines = ["已注册成员："]
            for row in rows:
                lines.append(f"{row['name']}｜{ROLE_LABEL[row['role']]}｜QQ {row['qq']}")
            await self.client.send_private(qq, "\n".join(lines))
            return

        if cmd == "管理员列表":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以查看管理员列表。")
                return
            rows = db.list_admins(self.workspace)
            await self.client.send_private(qq, "负责人 / 管理员：\n" + "\n".join(f"{r['name']}｜{ROLE_LABEL[r['role']]}" for r in rows))
            return

        if cmd == "停用成员":
            if not is_owner(user):
                await self.client.send_private(qq, "只有负责人可以查看停用成员。")
                return
            rows = db.list_inactive_users(self.workspace)
            if not rows:
                await self.client.send_private(qq, "当前没有停用成员。")
            else:
                await self.client.send_private(qq, "停用成员：\n" + "\n".join(f"{r['name']}｜QQ {r['qq']}" for r in rows))
            return

        if cmd in {"管理员添加", "管理员移除", "成员移除", "成员恢复"}:
            if not is_owner(user):
                await self.client.send_private(qq, "只有负责人可以执行成员/管理员管理。")
                return
            if not arg:
                await self.client.send_private(qq, f"用法：{cmd} <姓名或QQ>")
                return

            target_active = False if cmd == "成员恢复" else True
            status, target = db.resolve_user(self.workspace, arg, active=target_active)
            if status == "not_found":
                await self.client.send_private(qq, "没有找到对应成员。")
                return
            if status == "ambiguous":
                await self.client.send_private(qq, "存在同名成员，请改用 QQ 号：\n" + "\n".join(f"{r['name']}：{r['qq']}" for r in target))
                return
            if target["qq"] == self.owner_qq:
                await self.client.send_private(qq, "不能修改负责人自己的成员状态或角色。")
                return

            if cmd == "管理员添加":
                db.set_role(self.workspace, target["qq"], "admin")
                msg = f"已将 {target['name']} 设置为管理员。"
            elif cmd == "管理员移除":
                db.set_role(self.workspace, target["qq"], "member")
                msg = f"已将 {target['name']} 调整为普通成员。"
            elif cmd == "成员移除":
                ok = db.deactivate_user(self.workspace, target["qq"])
                msg = f"已停用成员 {target['name']}；其未完成认领已释放，完成历史仍保留。" if ok else "停用失败。"
            else:
                ok = db.reactivate_user(self.workspace, target["qq"])
                msg = f"已恢复成员 {target['name']}。" if ok else "恢复失败。"
            if (cmd in {"管理员添加", "管理员移除", "成员移除", "成员恢复"}) and ("失败" not in msg):
                db.log_action(self.workspace, qq, cmd, f"{target['name']} ({target['qq']})")
            await self.client.send_private(qq, msg)
            return

        await self.client.send_private(qq, "没有识别这个命令。发送“帮助”查看菜单。")

    async def handle_submit_session(self, qq: str, text: str):
        assert self.client
        task_id = submit_sessions.get(qq)
        if task_id is None:
            return
        if text in {"取消", "取消提交"}:
            submit_sessions.pop(qq, None)
            await self.client.send_private(qq, "已取消本次提交。")
            return
        result = db.submit_task(self.workspace, task_id, qq, text)
        if result == "ok":
            submit_sessions.pop(qq, None)
            await self.client.send_private(
                qq,
                f"任务 #{task_id} 已提交，当前状态：待审核。\n"
                "管理员审核通过后才会计为完成。",
            )

            # 谁发布的任务，就把待审核提醒发给谁。
            # 通知失败不影响成员已经成功提交的状态。
            task = db.get_task(self.workspace, task_id)
            submitter = db.get_user(self.workspace, qq, include_inactive=True)
            if task and task.get("publisher_qq"):
                publisher_qq = str(task["publisher_qq"])
                submitter_name = submitter["name"] if submitter else qq
                content = text if len(text) <= 1200 else text[:1200] + "…"
                notice = (
                    "【有新的任务提交待审核】\n"
                    f"任务 #{task_id}｜{task['task_type']}｜{task['subject']}\n"
                    f"提交人：{submitter_name}（{qq}）\n"
                    f"成果说明：{content}\n\n"
                    "审核命令：\n"
                    f"通过 {task_id} {qq}\n"
                    f"打回 {task_id} {qq} <原因>"
                )
                try:
                    await self.client.send_private(publisher_qq, notice)
                except Exception as e:
                    print(
                        f"[审核提醒] 任务 #{task_id} 通知发布人 {publisher_qq} 失败："
                        f"{type(e).__name__}: {e}"
                    )
            return
        messages = {
            "empty": "成果说明不能为空，请重新输入。",
            "not_found": "任务不存在。",
            "deleted": "任务已被删除。",
            "not_claimed": "你没有认领这个任务。",
            "already_submitted": "这个任务已经提交，正在等待审核。",
            "already_completed": "这个任务已经审核完成。",
        }
        submit_sessions.pop(qq, None)
        await self.client.send_private(qq, messages.get(result, "提交失败。"))

    async def handle_publish_session(self, qq: str, text: str):
        assert self.client
        user = db.get_user(self.workspace, qq)
        if not is_admin(user):
            publish_sessions.pop(qq, None)
            await self.client.send_private(qq, "你的管理员权限已失效，本次发布已取消。")
            return
        if text in {"取消", "取消发布"}:
            publish_sessions.pop(qq, None)
            await self.client.send_private(qq, "已取消发布任务。")
            return

        s = publish_sessions[qq]
        if s.stage == "task_type":
            s.task_type = text
            s.stage = "subject"
            await self.client.send_private(qq, "第 2/4 步：请输入【任务涉及学科】。")
            return
        if s.stage == "subject":
            s.subject = text
            s.stage = "required_knowledge"
            await self.client.send_private(qq, "第 3/4 步：请输入【任务所需知识】。")
            return
        if s.stage == "required_knowledge":
            s.required_knowledge = text
            s.stage = "required_people"
            await self.client.send_private(qq, "第 4/4 步：请输入【人数需求】（正整数，例如 1、2、3）。")
            return
        if s.stage == "required_people":
            try:
                n = int(text)
            except ValueError:
                n = 0
            if n < 1:
                await self.client.send_private(qq, "人数需求必须是大于等于 1 的整数，请重新输入。")
                return
            s.required_people = n
            s.stage = "confirm"
            await self.client.send_private(
                qq,
                "请确认任务信息：\n\n"
                f"任务类型：{s.task_type}\n"
                f"涉及学科：{s.subject}\n"
                f"所需知识：{s.required_knowledge}\n"
                f"人数需求：{s.required_people} 人\n\n"
                "发送“确认”正式发布；发送“取消”终止。",
            )
            return
        if s.stage == "confirm":
            if text != "确认":
                await self.client.send_private(qq, "请发送“确认”发布，或发送“取消”终止。")
                return
            task_id = db.create_task(
                self.workspace,
                s.task_type,
                s.subject,
                s.required_knowledge,
                s.required_people,
                qq,
            )
            publish_sessions.pop(qq, None)
            db.log_action(self.workspace, qq, "发布任务", f"任务 #{task_id}｜{s.task_type}｜{s.subject}｜人数 {s.required_people}")
            task = db.get_task(self.workspace, task_id)
            await self.client.send_private(qq, f"任务发布成功。\n\n{task_summary(task)}")
            announcement = (
                f"【新任务 #{task_id}】\n"
                f"任务类型：{task['task_type']}\n"
                f"涉及学科：{task['subject']}\n"
                f"所需知识：{task['required_knowledge']}\n"
                f"人数需求：{task['required_people']} 人\n"
                f"发布人：{task['publisher_name']}\n\n"
                f"需要接取的成员请私聊我发送：认领 {task_id}"
            )
            try:
                await self.client.send_group(self.workspace, announcement)
            except Exception as e:
                await self.client.send_private(qq, f"任务已保存，但群公告发送失败：{e}")

    async def handle_edit_session(self, qq: str, text: str):
        assert self.client
        user = db.get_user(self.workspace, qq)
        if not is_admin(user):
            edit_sessions.pop(qq, None)
            await self.client.send_private(qq, "你的管理员权限已失效，本次修改已取消。")
            return
        if text in {"取消", "取消修改"}:
            edit_sessions.pop(qq, None)
            await self.client.send_private(qq, "已取消修改任务。")
            return

        s = edit_sessions[qq]
        if s.stage == "choose_field":
            field_map = {
                "类型": "task_type",
                "任务类型": "task_type",
                "学科": "subject",
                "涉及学科": "subject",
                "知识": "required_knowledge",
                "所需知识": "required_knowledge",
                "人数": "required_people",
                "人数需求": "required_people",
            }
            field = field_map.get(text)
            if not field:
                await self.client.send_private(qq, "请选择：类型 / 学科 / 知识 / 人数，或发送“取消”。")
                return
            s.field = field
            s.stage = "value"
            prompts = {
                "task_type": "请输入新的【任务类型】。",
                "subject": "请输入新的【任务涉及学科】。",
                "required_knowledge": "请输入新的【任务所需知识】。",
                "required_people": "请输入新的【人数需求】（正整数）。",
            }
            await self.client.send_private(qq, prompts[field])
            return

        if s.stage == "value":
            result = db.update_task_field(self.workspace, s.task_id, s.field, text)
            if result == "ok":
                task = db.get_task(self.workspace, s.task_id)
                db.log_action(self.workspace, qq, "修改任务", f"任务 #{s.task_id}｜字段 {s.field}｜新值：{text}")
                edit_sessions.pop(qq, None)
                await self.client.send_private(qq, f"修改成功。\n\n{task_summary(task)}")
                await self.announce(f"【任务 #{s.task_id} 信息更新】\n{task['task_type']}｜{task['subject']}\n人数需求：{task['required_people']} 人")
                return
            if result == "below_assigned":
                task = db.get_task(self.workspace, s.task_id)
                await self.client.send_private(
                    qq,
                    f"不能把人数需求改成 {text}：当前已经有 {task['assigned_count']} 人参与。\n"
                    "请先移除多余认领，或者输入不小于当前参与人数的数字。",
                )
                return
            if result == "invalid_people":
                await self.client.send_private(qq, "人数需求必须是大于等于 1 的整数，请重新输入。")
                return
            if result == "empty":
                await self.client.send_private(qq, "内容不能为空，请重新输入。")
                return
            edit_sessions.pop(qq, None)
            await self.client.send_private(qq, "任务不存在或已删除，本次修改已结束。")

    async def handle_delete_session(self, qq: str, text: str):
        assert self.client
        user = db.get_user(self.workspace, qq)
        if not is_admin(user):
            delete_sessions.pop(qq, None)
            await self.client.send_private(qq, "你的管理员权限已失效，本次删除已取消。")
            return
        task_id = delete_sessions[qq]
        if text in {"取消", "取消删除"}:
            delete_sessions.pop(qq, None)
            await self.client.send_private(qq, "已取消删除任务。")
            return
        if text != "确认":
            await self.client.send_private(qq, "请发送“确认”删除，或发送“取消”放弃。")
            return
        result = db.soft_delete_task(self.workspace, task_id)
        delete_sessions.pop(qq, None)
        if result == "ok":
            db.log_action(self.workspace, qq, "删除任务", f"任务 #{task_id}")
            await self.announce(f"【任务 #{task_id} 已撤下】")
        messages = {
            "ok": f"任务 #{task_id} 已删除（软删除）。需要找回时发送：恢复任务 {task_id}",
            "not_found": "任务不存在。",
            "already_deleted": "这个任务已经被删除。",
        }
        await self.client.send_private(qq, messages.get(result, "删除失败。"))


async def main():
    config = load_config()
    bot = NotesDepartmentBot(config)
    await bot.start()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n已停止。")
