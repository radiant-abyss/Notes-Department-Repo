from __future__ import annotations

import asyncio
import csv
import inspect
import json
import re
import subprocess
import time
import uuid
import zipfile
from datetime import datetime, timedelta
from dataclasses import dataclass
from pathlib import Path
from typing import Dict

from openpyxl import Workbook

import websockets

import db
import volunteer

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"

ROLE_LABEL = {
    "owner": "Owner（负责人）",
    "admin": "Admin（管理员）",
    "member": "Member（工作成员）",
}

REGISTRATION_HELP = "你还没有注册。请先加入笔记部工作群，再私聊发送：注册 姓名\n例如：注册 张三\n注册后发送“完善信息”补齐学号、主修学院和辅修学院；没有辅修填“无”。"
MEMBER_HELP = """笔记部 Bot｜成员菜单
【个人资料】
完善信息（按提示填写；中途可发送“取消”）
完善信息 <学号> | <主修学院> | <辅修学院>
未来技术学院填写“未来技术学院-某某学院/学部”，没有辅修填“无”。
我的信息 / 我的身份
【任务】
任务 / 任务 <编号> / 认领 <编号>
我的待办 / 我的任务 / 我的历史
更新进度 <编号>（随后填写说明，不是正式提交）
提交 <编号> / 完成 <编号>（随后填写成果）
撤回提交 <编号> / 取消认领 <编号>
【申请部门发布任务】
申请任务（填写后发送“确认”；批准后公开认领，不自动分配给你）
我的申请 [页码] / 申请详情 <申请编号> / 撤回申请 <申请编号>
【本人志愿时长】
我的志愿时长 [YYYY-MM]
志愿时长明细 [YYYY-MM] [页码]
默认本月，从 2026-10-30 起记录；待登记不等于 0 小时。
发送“帮助”或“菜单”显示当前身份可用的指令。"""
ADMIN_HELP = '🛠️ 日常管理帮助\n\n【发布与查看任务】\n发布任务\n按提示逐项填写，最后确认发布。\n所有任务\n查看部门任务列表。\n\n【审核成果】\n待审核\n查看等待审核的提交。\n通过 <编号> <姓名或QQ>\n打回 <编号> <姓名或QQ> <原因>\n例如：通过 12 张三\n普通管理员审核时不填小时数，时长由高级管理员登记。\n\n【调整任务】\n修改任务 <编号>\n按提示选择要修改的内容。\n重开 <编号>\n重新开放 <编号>\n移除认领 <编号> <姓名或QQ>\n请核对任务和成员后操作。\n\n【删除与恢复】\n删除任务 <编号>\n已删除任务\n恢复任务 <编号>\n删除为软删除，可恢复。\n\n【审批发布申请】\n待审批申请 [页码]\n批准申请 <申请编号>\n驳回申请 <申请编号> <原因>\n详细说明请发送“申请帮助”。\n\n【查看成员】\n成员列表\n信息未齐成员\n管理员列表\n停用成员\n\n【调整成员与管理员】\n管理员添加 <姓名或QQ>\n管理员移除 <姓名或QQ>\n成员移除 <姓名或QQ>\n成员恢复 <姓名或QQ>\n有同名成员时，请使用 QQ 号。\n\n【统计与导出】\n本周统计\n本月统计\n成员统计 <姓名或QQ>\n操作日志 [条数]\n导出数据\n导出CSV\n\n普通管理员不能操作高级管理者，导出和日志不包含部门时长账目。\n尖括号内容需替换为实际值，方括号参数可以省略。'
SENIOR_HELP = '📊 高级管理员帮助\n\n【查询部门时长】\n待登记时长 [YYYY-MM] [页码]\n查看尚未登记时长的任务份额。\n志愿时长统计 [YYYY-MM] [页码]\n查看全体成员的月度统计。\n导出志愿时长 [YYYY-MM]\n生成月报；不会清空历史记录。\n\n【审核时登记】\n通过 <任务编号> <姓名或QQ> <小时数>\n例如：通过 12 张三 2.5\n不填小时数也可审核，新增账目会标为待登记。\n\n【补录时长】\n登记志愿时长 <任务编号> <姓名或QQ> <小时数> [YYYY-MM-DD]\n例如：登记志愿时长 12 张三 2.5 2026-10-30\n省略日期则归入今天；补往月请明确填写日期。\n\n【更正时长】\n更正志愿时长 <记录编号> <新总小时数> <YYYY-MM-DD> <原因>\n例如：更正志愿时长 5 3 2026-10-30 工作量核对后调整\n这里的编号是时长记录编号，不是任务编号。\n填写新总时长，不是追加；填 0 可冲销，修改留痕。\n\n【任免高级管理员】\n时长负责人添加 <姓名或QQ>\n时长负责人移除 <姓名或QQ>\n时长负责人列表\n时长负责人即高级管理员，拥有日常管理和时长管理权限。\n有同名成员时，请使用 QQ 号。\n\n【登记说明】\n小时数最多两位小数。\n待登记不等于 0 小时。\n任务重开或删除不会自动抹除账目，需要明确更正。\n尖括号内容需替换为实际值，方括号参数可以省略。'
ADMIN_COMMANDS = {"发布任务","所有任务","待审核","通过","打回","修改任务","重开","重新开放","删除任务","恢复任务","已删除任务","移除认领","待审批申请","批准申请","驳回申请","成员列表","信息未齐成员","管理员列表","停用成员","管理员添加","管理员移除","成员移除","成员恢复","操作日志","本周统计","本月统计","成员统计","导出数据","导出CSV"}
SENIOR_COMMANDS = {"待登记时长","志愿时长统计","导出志愿时长","登记志愿时长","更正志愿时长","时长负责人添加","时长负责人移除","时长负责人列表"}


def role_label(user):
    if user['role'] == 'owner':
        return ROLE_LABEL['owner']
    if volunteer.has_appointment(user['workspace'], user['qq']):
        return '高级管理员（志愿时长负责人）'
    return ROLE_LABEL[user['role']]


HELP_COMMANDS = {"帮助", "help", "菜单", "任务帮助", "时长帮助", "申请帮助", "管理帮助", "全部帮助"}


def format_detail_help(text: str) -> str:
    """Keep semantic paragraphs; put parallel commands on separate lines."""
    return text.replace(" / ", "\n")


def help_for(user, command="帮助"):
    if not user or not user['active']:
        return REGISTRATION_HELP
    greeting = f"尊敬的 {user['name']} 您好，欢迎使用笔记部Bot！(∠・ω< )⌒☆"
    complete = all(user[key] for key in ("student_number", "major_college", "minor_college"))
    admin = is_admin(user)
    senior = volunteer.can_manage(user['workspace'], user['qq'])
    profile_hint = ('若要修改个人信息，请发送“完善信息”。' if complete
                    else '个人信息尚未填齐，请发送“完善信息”，按提示填写学号、主修学院和辅修学院；没有辅修填“无”。')
    if command in {"帮助", "help", "菜单"}:
        text = (f"{greeting}\n\n"
                "1 📋 我的待办\n2 📂 我的任务\n3 📝 提交任务\n4 ✏️ 更新进度\n"
                "5 ⏱️ 我的志愿时长\n6 👤 我的信息\n7 💡 申请任务\n\n"
                "🔎 想领取新任务？\n发送“任务”查看可领取列表。\n例如领取任务 #12，发送“认领 12”。\n“我的任务”只查看自己已经领取的任务。\n\n"
                + profile_hint + "\n\n📖 详细用法与完整帮助：任务帮助 / 时长帮助 / 申请帮助")
        if admin:
            text += "\n日常管理：管理帮助"
        if senior:
            text += "\n高级管理员帮助：时长帮助"
        return text + ("\n全部指令：全部帮助" if admin else "")
    if command == "全部帮助" and not admin:
        return "请发送“任务帮助”“时长帮助”或“申请帮助”查看详细用法。"
    task_help = '📋 任务帮助\n\n【查看任务】\n发送：任务\n查看可认领的任务。\n发送：任务 12\n查看任务 #12 的详情。\n\n【认领任务】\n发送：认领 12\n认领任务 #12 的一个名额。\n\n【查看自己的任务】\n我的待办：查看需要处理的事项。\n我的任务：查看当前任务。\n我的历史：查看已完成记录。\n\n【更新进度】\n发送：更新进度 12\n随后按提示填写进度说明。\n也可直接发送：更新进度 12 初稿完成\n只更新自己的进度，不是正式提交。\n\n【提交成果】\n发送：提交 12\n随后填写成果说明或链接，等待管理者审核。\n“完成 12”也可以；填写中发送“取消”退出。\n\n【撤回或取消】\n发送：撤回提交 12\n撤回自己尚待审核的提交，修改后再提交。\n发送：取消认领 12\n在任务状态允许时释放自己的名额。\n\n以上 12 都是示例，请换成实际任务编号。'
    hours_help = '⏱️ 本人志愿时长\n\n【查看本月】\n发送：我的志愿时长\n查看已登记小时数和待登记份数。\n\n【查看历史月份】\n例如：我的志愿时长 2026-10\n把月份换成需要查询的月份。\n\n【查看逐笔明细】\n发送：志愿时长明细\n默认查看本月。\n指定月份：志愿时长明细 2026-10\n翻页示例：志愿时长明细 2026-10 2\n\n【统计说明】\n从 2026-10-30 起记录，按自然月统计。\n待登记表示尚未填写时长，不等于 0 小时。\n以上指令只查询本人。'
    admin_help = ADMIN_HELP
    if senior:
        admin_help = admin_help.replace("普通管理员不能操作高级管理者，导出和日志不包含部门时长账目。", "高级管理者可操作全部管理指令；总负责人账号受保护。")
    if command == "任务帮助":
        return task_help
    if command == "申请帮助":
        text = """💡 申请帮助

【申请发布任务】
发送：申请任务
按提示填写，最后发送“确认”。
批准后任务公开供成员认领，不会自动分配给你。

【需要填写的内容】
任务类型、学科、任务内容及所需知识、人数、申请理由。
Bot 会逐项询问，不用一次写完。
填写中发送“取消”可退出。

【查看我的申请】
发送：我的申请
翻页示例：我的申请 2

【查看申请详情】
发送：申请详情 编号
例如：申请详情 3

【撤回申请】
发送：撤回申请 编号
例如：撤回申请 3
只能撤回自己尚未审批的申请。"""
        if admin:
            text += """

🛠️ 申请审批（管理者）

查看待审批列表：
待审批申请 [页码]

批准并发布任务：
批准申请 <申请编号>

驳回并说明原因：
驳回申请 <申请编号> <原因>"""
        return text
    if command == "时长帮助":
        return hours_help + ("\n\n" + SENIOR_HELP if senior else "")
    if command == "管理帮助":
        return format_detail_help(admin_help) if admin else "该指令仅管理者可用。发送“帮助”查看成员指令。"
    profile = ("👤 个人资料\n\n" + profile_hint + "\n我的信息：查看本人资料。\n我的身份：同样可查看本人身份和资料。")
    if not complete:
        profile += "\n\n一次填写格式：\n完善信息 <学号> | <主修学院> | <辅修学院>\n未来技术学院填写“未来技术学院-某某学院/学部”，没有辅修填“无”。"
    return "\n\n".join((greeting, profile, task_help,
                            help_for(user, "申请帮助"), help_for(user, "时长帮助"), admin_help))


def clear_user_sessions(workspace,qq):
    for sessions in (profile_sessions,publish_sessions,edit_sessions,delete_sessions,submit_sessions):
        sessions.pop(qq,None)
    request_sessions.pop((workspace,qq),None)
    progress_sessions.pop((workspace,qq),None)
    menu_sessions.pop((workspace,qq),None)


def redact_volunteer_export(data):
    data = {k:v for k,v in data.items() if not k.startswith('volunteer_')}
    data['audit_logs'] = [r for r in data['audit_logs'] if '志愿时长' not in r['action'] and not r['action'].startswith('时长负责人')]
    return data

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
progress_sessions: Dict[tuple[str, str], int] = {}
menu_sessions: dict[tuple[str, str], tuple[str, float]] = {}
profile_sessions: Dict[str, list[str]] = {}
request_sessions: Dict[tuple[str, str], list[str]] = {}
REQUEST_PROMPTS = (
    "第 1/5 步：请输入任务类型（如笔记整理、习题校对）。",
    "第 2/5 步：请输入涉及学科。",
    "第 3/5 步：请输入具体任务内容及所需知识。",
    "第 4/5 步：请输入人数需求（1 到 100 的整数）。",
    "第 5/5 步：请输入申请部门发布这项任务的理由。",
)
REQUEST_STATUS = {'pending':'待审批','approved':'已批准','rejected':'已驳回','withdrawn':'已撤回'}


def request_summary(row):
    lines = [f"申请 #{row['id']}｜{REQUEST_STATUS[row['status']]}",
             f"申请人：{row['applicant_name']}", f"任务类型：{row['task_type']}",
             f"学科：{row['subject']}", f"任务内容及所需知识：{row['required_knowledge']}",
             f"人数：{row['required_people']}", f"申请理由：{row['reason']}",
             f"申请时间：{row['created_at']}"]
    if row.get('review_comment'):
        lines.append("审批意见：" + row['review_comment'])
    if row.get('reviewed_at'):
        lines.append("处理时间：" + row['reviewed_at'])
    if row.get('task_id'):
        lines.append(f"正式任务 #{row['task_id']}，向所有成员开放认领。")
    return "\n".join(lines)


PROFILE_USAGE = "完善信息 <学号> | <学院（主修）> | <辅修学院>\n例如：完善信息 0012345678 | 未来技术学院-计算机学院 | 无\n没有辅修学院请填“无”。也可只发“完善信息”，按提示逐项填写。"
PROFILE_PROMPTS = ("请输入学号（发送“取消”可退出，旧资料不会改变）：", "请输入学院（主修）。未来技术学院请填写：未来技术学院-某某学院", "请输入辅修学院，没有请填“无”：")


def profile_summary(user):
    return (f"学号：{user['student_number'] or '未填写'}\n"
            f"学院（主修）：{user['major_college'] or '未填写'}\n"
            f"辅修学院：{user['minor_college'] or '未填写'}")


def parse_profile_fields(text, count=3):
    fields = [part.strip() for part in text.replace("｜", "|").split("|")]
    if len(fields) != count or not all(fields):
        raise ValueError("请用 | 分隔各项，且每项不能为空。\n" + PROFILE_USAGE)
    return fields


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        raise RuntimeError("缺少 config.json。请复制 config.example.json 为 config.json 并填写。")
    data = json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))
    required = ["ws_url", "department_group_id", "owner_qq", "owner_name"]
    missing = [k for k in required if not str(data.get(k, "")).strip()]
    if missing:
        raise RuntimeError(f"config.json 缺少配置：{', '.join(missing)}")
    data["department_group_id"] = str(data["department_group_id"])
    data["owner_qq"] = str(data["owner_qq"])
    # OneBot Token 为可选项：留空可兼容旧配置；正式使用建议启用。
    data["onebot_token"] = str(data.get("onebot_token", "")).strip()

    # NapCat watchdog：仅负责“WebSocket 连续失败时尝试拉起 NapCat”。
    # 默认关闭，避免旧配置在没有 restart_napcat.bat 时意外触发。
    data["napcat_watchdog_enabled"] = bool(data.get("napcat_watchdog_enabled", False))
    data["napcat_restart_after_failures"] = max(1, int(data.get("napcat_restart_after_failures", 5)))
    data["napcat_restart_cooldown_seconds"] = max(0, int(data.get("napcat_restart_cooldown_seconds", 300)))
    data["napcat_restart_wait_seconds"] = max(0, int(data.get("napcat_restart_wait_seconds", 20)))
    data["napcat_restart_script"] = str(data.get("napcat_restart_script", "restart_napcat.bat")).strip() or "restart_napcat.bat"
    return data


def clean_text(text: str) -> str:
    text = text or ""
    text = re.sub(r"\[CQ:at,qq=\d+[^\]]*\]", "", text)
    text = text.strip()
    if text.startswith("/"):
        text = text[1:].strip()
    return re.sub(r"\s+", " ", text)


# 这些命令允许“命令参数”之间省略空格。
# 例如：认领6 / 认领 6、提交12 / 提交 12。
# 多参数命令只允许省略“命令”和第一个参数之间的空格，
# 例如：通过6 张三、打回6 张三 原因。
CONCAT_ARG_COMMANDS = tuple(sorted({
    "撤回提交", "取消认领", "重新开放", "修改任务", "删除任务", "恢复任务",
    "移除认领", "成员统计", "管理员添加", "管理员移除", "成员移除", "成员恢复",
    "待登记时长", "时长负责人添加", "时长负责人移除", "时长负责人列表", "我的志愿时长", "志愿时长明细", "志愿时长统计", "导出志愿时长", "登记志愿时长", "更正志愿时长", "申请详情", "撤回申请", "批准申请", "驳回申请", "待审批申请", "我的申请", "操作日志", "认领", "提交", "完成", "通过", "打回", "重开", "注册", "任务", "更新进度",
}, key=len, reverse=True))


def split_command(text: str):
    text = clean_text(text)
    if not text:
        return "", ""

    # 先识别“命令紧贴第一个参数”的写法。按最长命令优先，避免
    # “撤回提交12”之类被较短命令错误识别。
    # 如果命令后本来就是空格，则交给下面的传统解析。
    for command in CONCAT_ARG_COMMANDS:
        if (
            text.startswith(command)
            and len(text) > len(command)
            and not text[len(command)].isspace()
        ):
            return command, text[len(command):].strip()

    # 原有带空格写法保持兼容。
    if " " in text:
        parts = text.split(" ", 1)
        return parts[0], parts[1].strip()

    return text, ""


def is_admin(user) -> bool:
    return bool(user and user["active"] and (user["role"] in {"owner", "admin"} or volunteer.can_manage(user["workspace"], user["qq"])))


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


def progress_summary(assignment: dict) -> str:
    progress = assignment.get("progress_text") or "尚未填写"
    when = assignment.get("progress_updated_at")
    return f"最近进度：{progress}" + (f"\n更新时间：{when}" if when else "")


def todo_summary(user, tasks, qq: str) -> str:
    pending, waiting = [], []
    missing = [label for key, label in (("student_number", "学号"), ("major_college", "学院（主修）"), ("minor_college", "辅修学院")) if not user[key]]
    if missing:
        pending.append("个人资料待补齐：" + "、".join(missing) + "\n发送“完善信息”按提示填写。")
    for task in tasks:
        mine = next((p for p in task["assignees"] if p["qq"] == str(qq)), None)
        if not mine:
            continue
        title = f"#{task['id']}｜{task['task_type']}｜{task['subject']}"
        if mine["status"] == "submitted":
            waiting.append(title + "\n等待审核，无需重复提交。\n" + progress_summary(mine))
        elif mine["status"] == "claimed":
            rejected = bool(mine.get("reviewer_qq") and mine.get("reviewed_at"))
            item = title + ("\n需修改后重新提交" if rejected else "\n进行中，完成后请提交")
            if rejected:
                item += "\n审核意见：" + (mine.get("review_comment") or "需要继续修改")
            pending.append(item + "\n" + progress_summary(mine))
    sections = [f"我的待办：{len(pending)} 项待处理", "\n\n".join(pending) if pending else "目前没有需要你处理的事项。"]
    if waiting:
        sections.extend([f"等待审核：{len(waiting)} 项（不计入待处理）", "\n\n".join(waiting)])
    return "\n\n".join(sections)


def task_summary(task: dict, show_progress: bool = False) -> str:
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
    if show_progress and task.get("assignees"):
        lines.append("成员进度：\n" + "\n\n".join(f"{p['name']}\n{progress_summary(p)}" for p in task["assignees"]))
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


def make_export_xlsx(workspace: str, include_volunteer=True) -> Path:
    data = db.export_data(workspace)
    if not include_volunteer: data = redact_volunteer_export(data)
    export_dir = BASE_DIR / "exports"
    export_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:10]
    path = export_dir / f"notes_department_{stamp}.xlsx"
    wb = Workbook()
    default = wb.active
    wb.remove(default)
    mapping = [
        ("成员", data["users"]),
        ("任务", data["tasks"]),
        ("认领与提交", data["assignments"]),
        ("操作日志", data["audit_logs"]),
        ("任务发布申请", data["task_requests"]),
        ("志愿时长", data.get("volunteer_entries", [])),
        ("时长修改记录", data.get("volunteer_changes", [])),
        ("时长负责人", data.get("volunteer_managers", [])),
    ]
    for title, rows in mapping:
        if not include_volunteer and title in {"志愿时长", "时长修改记录", "时长负责人"}:
            continue
        ws = wb.create_sheet(title)
        _write_sheet(ws, rows)
    wb.save(path)
    return path


def make_export_csv_zip(workspace: str, include_volunteer=True) -> Path:
    data = db.export_data(workspace)
    if not include_volunteer: data = redact_volunteer_export(data)
    export_dir = BASE_DIR / "exports"
    export_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:10]
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
    def __init__(
        self,
        ws_url: str,
        handler,
        access_token: str = "",
        *,
        watchdog_enabled: bool = False,
        restart_after_failures: int = 5,
        restart_cooldown_seconds: int = 300,
        restart_wait_seconds: int = 20,
        restart_script: str = "restart_napcat.bat",
    ):
        self.ws_url = ws_url
        self.handler = handler
        self.access_token = str(access_token or "").strip()
        self.ws = None
        self.pending: Dict[str, asyncio.Future] = {}
        self.send_lock = asyncio.Lock()

        self.watchdog_enabled = bool(watchdog_enabled)
        self.restart_after_failures = max(1, int(restart_after_failures))
        self.restart_cooldown_seconds = max(0, int(restart_cooldown_seconds))
        self.restart_wait_seconds = max(0, int(restart_wait_seconds))
        script_path = Path(restart_script)
        if not script_path.is_absolute():
            script_path = BASE_DIR / script_path
        self.restart_script = script_path.resolve()
        self.consecutive_failures = 0
        self.last_restart_monotonic: float | None = None
        self.restart_attempts = 0
        self.restart_limit_reported = False
        self.last_recovery_status = None
        self.restart_lock = asyncio.Lock()

    def _report_recovery_status(self):
        # The task writes only a state code, never credentials or raw login URLs.
        path = self.restart_script.parent / "napcat_recovery_status.json"
        try:
            if time.time() - path.stat().st_mtime > 180:
                return
            state = json.loads(path.read_text(encoding="utf-8-sig")).get("state")
        except (OSError, ValueError):
            return
        if state == self.last_recovery_status:
            return
        self.last_recovery_status = state
        labels = {
            "starting": "正在启动 NapCat，继续等待连接。",
            "waiting_login": "NapCat 已运行，正在等待 QQ 登录；不会重复启动。",
            "manual_verification": "QQ 需要验证码或安全验证，请在 NapCat 页面完成；不会重复启动。",
            "online_port_ready": "QQ 在线且 OneBot 端口已就绪，等待 Bot 实际连接。",
            "online_port_missing": "QQ 在线，但 OneBot 端口未就绪，请检查 NapCat 网络配置。",
            "existing_unverified": "检测到已有服务但无法确认状态，保留进程，避免重复启动。",
            "timeout": "本次恢复等待超时，Bot 仍会重连，并按间隔进行有限恢复尝试。",
            "script_error": "恢复脚本执行失败，请检查计划任务权限及本地凭据。",
        }
        if state in labels:
            print("[Watchdog] " + labels[state])

    async def _maybe_restart_napcat(self) -> bool:
        if not self.watchdog_enabled:
            return False
        self._report_recovery_status()
        if self.consecutive_failures < self.restart_after_failures:
            return False
        async with self.restart_lock:
            if self.restart_attempts >= 3:
                if not self.restart_limit_reported:
                    print("[Watchdog] 本次故障已尝试恢复 3 次，暂停自动拉起；仍持续重连。请检查 NapCat 登录页面。")
                    self.restart_limit_reported = True
                return False
            now = time.monotonic()
            # Give the scheduled task up to 90 seconds to finish. Preserve the
            # user's configured minimum cooldown and back off on later attempts.
            delay = max(120, self.restart_cooldown_seconds) * max(1, self.restart_attempts)
            if self.last_restart_monotonic is not None and now - self.last_restart_monotonic < delay:
                return False
            self.restart_attempts += 1
            self.last_restart_monotonic = now
            if not self.restart_script.exists():
                print(f"[Watchdog] 找不到恢复脚本：{self.restart_script.name}")
                return False
            print(f"[Watchdog] 发起第 {self.restart_attempts}/3 次恢复检查：{self.restart_script.name}")
            try:
                proc = await asyncio.create_subprocess_exec(
                    "cmd.exe", "/c", str(self.restart_script),
                    cwd=str(self.restart_script.parent),
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                )
                try:
                    code = await asyncio.wait_for(proc.wait(), timeout=15)
                except asyncio.TimeoutError:
                    proc.kill()
                    await proc.wait()
                    print("[Watchdog] 触发计划任务超时，继续重连。")
                    return False
            except Exception as e:
                print(f"[Watchdog] 触发恢复失败：{type(e).__name__}")
                return False
            if code != 0:
                print(f"[Watchdog] 计划任务触发失败，退出码 {code}。")
            else:
                print("[Watchdog] 已请求恢复检查，尚未确认登录成功；继续重连。")
            return False

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
                    self.consecutive_failures = 0
                    self.restart_attempts = 0
                    self.restart_limit_reported = False
                    self.last_restart_monotonic = None
                    self.last_recovery_status = None
                    print("[NapCat] OneBot WebSocket 已连接。")
                    await self.recv_loop(ws)
                    raise ConnectionError("NapCat closed the WebSocket")
            except asyncio.CancelledError:
                raise
            except Exception as e:
                self.consecutive_failures += 1
                print(f"[NapCat] 连接中断：{type(e).__name__}: {e}")
                if self.watchdog_enabled:
                    print(
                        f"[Watchdog] 当前连续失败：{self.consecutive_failures}/"
                        f"{self.restart_after_failures}"
                    )
            finally:
                self.ws = None
                for fut in self.pending.values():
                    if not fut.done():
                        fut.set_exception(ConnectionError("NapCat WebSocket 已断开"))
                self.pending.clear()

            restarted = await self._maybe_restart_napcat()
            if restarted:
                continue
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

    async def send_message_checked(self, action: str, params: dict):
        response = await self.call_api(action, params)
        if response.get("status") == "ok" and response.get("retcode") == 0:
            return response
        # Do not dump response bodies: they may contain message content or credentials.
        detail = str(response.get("message", "")) + str(response.get("wording", ""))
        reason = ("QQ 内核报告网络连接异常（1006514）" if "1006514" in detail
                  else "发送超时，未确认送达" if "Timeout" in detail
                  else "NapCat 拒绝发送或未确认成功")
        code = response.get("retcode")
        code = code if isinstance(code, int) else "unknown"
        error = f"{action} 失败：retcode={code}，{reason}。WebSocket 已连接不代表消息已送达。"
        print("[消息发送失败] " + error)
        raise RuntimeError(error)

    async def send_private(self, qq: str, text: str):
        return await self.send_message_checked("send_private_msg", {"user_id": int(qq), "message": text})

    async def send_group(self, group_id: str, text: str):
        return await self.send_message_checked("send_group_msg", {"group_id": int(group_id), "message": text})

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
            watchdog_enabled=self.config.get("napcat_watchdog_enabled", False),
            restart_after_failures=self.config.get("napcat_restart_after_failures", 5),
            restart_cooldown_seconds=self.config.get("napcat_restart_cooldown_seconds", 300),
            restart_wait_seconds=self.config.get("napcat_restart_wait_seconds", 20),
            restart_script=self.config.get("napcat_restart_script", "restart_napcat.bat"),
        )
        print(f"[Bot] 工作群：{self.workspace}")
        print(f"[Bot] Owner：{self.owner_name} ({self.owner_qq})")
        print("[安全] OneBot Token：" + ("已启用" if self.config.get("onebot_token") else "未启用"))
        if self.config.get("napcat_watchdog_enabled", False):
            print(
                "[Watchdog] 已启用：连续失败 "
                f"{self.config.get('napcat_restart_after_failures', 5)} 次后尝试重启 NapCat；"
                f"冷却 {self.config.get('napcat_restart_cooldown_seconds', 300)} 秒；"
                "同一次故障最多检查恢复 3 次，间隔递增；已有实例不重复启动。"
            )
        else:
            print("[Watchdog] 未启用。")
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

        member = db.get_user(self.workspace,str(event.get('user_id')))
        if not member:
            await self.client.send_group(self.workspace,"请先私聊我发送：注册 姓名。已停用的成员请联系管理者恢复。")
            return
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

        user = db.get_user(self.workspace, qq)
        inactive = db.get_user(self.workspace, qq, include_inactive=True) if not user else None
        if inactive and not inactive['active']:
            clear_user_sessions(self.workspace,qq)
            await self.client.send_private(qq,"你的成员资格已停用，请联系管理者恢复；不能通过重新注册自行恢复。")
            return
        if not user:
            clear_user_sessions(self.workspace,qq)
            if cmd != "注册":
                await self.client.send_private(qq,REGISTRATION_HELP)
                return
        if cmd in HELP_COMMANDS:
            key = (self.workspace, qq)
            menu_sessions.pop(key, None)
            busy = (any(qq in sessions for sessions in (profile_sessions, publish_sessions, edit_sessions, delete_sessions, submit_sessions))
                    or key in request_sessions or key in progress_sessions)
            reply = help_for(user, cmd)
            if cmd in {"帮助", "help", "菜单"}:
                if busy:
                    reply += "\n你正在填写内容；请先完成，或发送“取消”后重新发送“帮助”选择菜单。"
                else:
                    menu_sessions[key] = ("choose", time.monotonic() + 300)
            await self.client.send_private(qq, reply)
            return
        if user and cmd in SENIOR_COMMANDS and not volunteer.can_manage(self.workspace,qq):
            await self.client.send_private(qq,"该指令仅总负责人和高级管理员可用。发送“帮助”查看你的指令。")
            return
        if user and cmd in ADMIN_COMMANDS and not is_admin(user):
            await self.client.send_private(qq,"该指令仅管理者可用。发送“帮助”查看成员指令。")
            return

        if qq in profile_sessions:
            if text == "取消":
                profile_sessions.pop(qq, None)
                await self.client.send_private(qq, "已取消补填，原资料和任务记录未改变。")
                return
            values = profile_sessions[qq]
            candidate = values + [text]
            probe = candidate + ["占位"] * (3 - len(candidate))
            try:
                db.validate_user_profile(*probe)
                if len(candidate) == 3:
                    updated = db.update_user_profile(self.workspace, qq, *candidate)
                    profile_sessions.pop(qq, None)
                    await self.client.send_private(qq, "信息已保存，已有任务记录保留。\n" + profile_summary(updated))
                else:
                    profile_sessions[qq] = candidate
                    await self.client.send_private(qq, PROFILE_PROMPTS[len(candidate)])
            except ValueError as exc:
                await self.client.send_private(qq, str(exc) + "\n" + PROFILE_PROMPTS[len(values)])
            return

        if (self.workspace, qq) in request_sessions:
            await self.handle_request_session(qq, text)
            return

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
        if (self.workspace, qq) in progress_sessions:
            await self.handle_progress_session(qq, text)
            return

        if qq in submit_sessions:
            await self.handle_submit_session(qq, text)
            return

        if await self.handle_menu_session(qq, text):
            return

        if cmd == "注册":
            if not arg:
                await self.client.send_private(qq, "用法：注册 <姓名>\n例如：注册 张三")
                return
            if not await self.is_group_member(qq):
                await self.client.send_private(qq, "注册失败：你目前不在配置的笔记部工作群中。")
                return
            try:
                if "|" in arg or "｜" in arg:
                    name, student, major, minor = parse_profile_fields(arg, 4)
                    user = db.register_user(self.workspace, qq, name, self.owner_qq, student_number=student, major_college=major, minor_college=minor)
                else:
                    user = db.register_user(self.workspace, qq, arg, self.owner_qq)
            except ValueError as exc:
                await self.client.send_private(qq, str(exc))
                return
            message = f"注册成功。\n姓名：{user['name']}\n身份：{role_label(user)}\n" + profile_summary(user)
            if any(not user[key] for key in ("student_number", "major_college", "minor_college")):
                message += "\n请发送“完善信息”补齐资料；不用重新注册，原任务记录会保留。"
            await self.client.send_private(qq, message)
            return

        user = db.get_user(self.workspace, qq)
        if not user:
            inactive = db.get_user(self.workspace, qq, include_inactive=True)
            if inactive and not inactive["active"]:
                await self.client.send_private(qq, "你的成员资格目前处于停用状态，请联系负责人。")
            else:
                await self.client.send_private(qq, "你还没有注册。请发送：注册 <姓名>")
            return

        if await self.handle_volunteer_command(user, qq, cmd, arg):
            return

        if await self.handle_request_command(user, qq, cmd, arg):
            return

        if cmd == "完善信息":
            if not arg:
                profile_sessions[qq] = []
                await self.client.send_private(qq, PROFILE_PROMPTS[0])
                return
            try:
                updated = db.update_user_profile(self.workspace, qq, *parse_profile_fields(arg))
            except ValueError as exc:
                await self.client.send_private(qq, str(exc))
                return
            await self.client.send_private(qq, "信息已保存，已有任务记录保留。\n" + profile_summary(updated))
            return

        if cmd in {"我的身份", "我的信息"}:
            await self.client.send_private(qq, f"姓名：{user['name']}\n身份：{role_label(user)}\nQQ：{qq}\n" + profile_summary(user))
            return

        if cmd == "任务":
            if arg:
                tid = parse_task_id(arg)
                if tid is None:
                    await self.client.send_private(qq, "用法：任务 <编号>")
                    return
                task = db.get_task(self.workspace, tid)
                await self.client.send_private(qq, task_summary(task, show_progress=is_admin(user)) if task else "没有找到这个任务。")
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
                    chunks.append(f"\n例如查看任务 #{rows[0]['id']}：任务 {rows[0]['id']}\n认领该任务：认领 {rows[0]['id']}\n请把编号换成你想领取的任务编号。")
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

        if cmd == "我的待办":
            await self.client.send_private(qq, todo_summary(user, db.list_my_active_tasks(self.workspace, qq), qq))
            return

        if cmd == "更新进度":
            parts = arg.split(maxsplit=1)
            tid = parse_task_id(parts[0]) if parts else None
            if tid is None:
                await self.client.send_private(qq, "用法：更新进度 <编号>，例如：更新进度 12。随后按提示填写说明。也可直接发送：更新进度 12 初稿完成。")
                return
            if len(parts) == 2:
                await self.save_progress(qq, tid, parts[1])
                return
            task = db.get_task(self.workspace, tid, include_deleted=True)
            assignment = db.get_assignment(self.workspace, tid, qq)
            if not task or task['deleted_at'] is not None:
                await self.client.send_private(qq, "任务不存在或已删除，不能更新进度。")
                return
            if not assignment:
                await self.client.send_private(qq, "你没有认领这个任务，只能更新自己的进度。")
                return
            if task['status'] == 'completed' or assignment['status'] == 'completed':
                await self.client.send_private(qq, "该任务或你的任务份额已经完成，不能更新进度。")
                return
            if assignment['status'] == 'submitted':
                await self.client.send_private(qq, f"正在等待审核；如需修改，请先发送：撤回提交 {tid}")
                return
            progress_sessions[(self.workspace, qq)] = tid
            await self.client.send_private(qq, f"请发送任务 #{tid} 的进度说明（1000 字以内）。\n发送“取消”可退出；更新进度不是正式提交。")
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
                    mine = next((p for p in task["assignees"] if p["qq"] == qq), None)
                    if mine:
                        chunks[-1] += "\n" + ("状态：等待审核" if mine["status"] == "submitted" else "状态：进行中") + "\n" + progress_summary(mine)
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
                try:
                    result = db.approve_submission(self.workspace, tid, target["qq"], qq, parts[2] if len(parts) == 3 else None)
                except ValueError as exc:
                    await self.client.send_private(qq, str(exc))
                    return
                if result == "ok":
                    await self.notify_volunteer(tid, target["qq"], "成果审核通过，请核对时长或补齐待登记记录。")
                    db.log_action(self.workspace, qq, "审核通过", f"任务 #{tid}｜成员 {target['name']} ({target['qq']})")
                    await self.client.send_private(qq, f"已通过 {target['name']} 在任务 #{tid} 的提交。" + (f"\n本次登记 {volunteer.hours(volunteer.parse_hours(parts[2]))} 小时。" if len(parts)==3 else "\n未填写时长：新记录待登记；已有账目保留，请核对时长明细。"))
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
            rows = db.list_audit_logs(self.workspace, limit, include_volunteer=volunteer.can_manage(self.workspace,qq))
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
                path = make_export_xlsx(self.workspace, volunteer.can_manage(self.workspace,qq)) if cmd == "导出数据" else make_export_csv_zip(self.workspace, volunteer.can_manage(self.workspace,qq))
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

        if cmd in {"成员列表", "信息未齐成员"}:
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以查看成员列表。")
                return
            rows = db.list_users(self.workspace)
            if cmd == "信息未齐成员":
                rows = [row for row in rows if any(not row[key] for key in ("student_number", "major_college", "minor_college"))]
                await self.client.send_private(qq, "信息未齐成员：\n" + ("\n".join(f"{row['name']}｜QQ {row['qq']}" for row in rows) or "所有有效成员均已填写完整。"))
                return
            lines = ["已注册成员："]
            for row in rows:
                lines.append(f"{row['name']}｜{role_label(row)}｜QQ {row['qq']}\n" + profile_summary(row))
            await self.client.send_private(qq, "\n".join(lines))
            return

        if cmd == "管理员列表":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人和管理员可以查看管理员列表。")
                return
            rows = db.list_admins(self.workspace)
            await self.client.send_private(qq, "负责人 / 管理员：\n" + "\n".join(f"{r['name']}｜{role_label(r)}" for r in rows))
            return

        if cmd == "停用成员":
            if not is_admin(user):
                await self.client.send_private(qq, "只有负责人可以查看停用成员。")
                return
            rows = db.list_inactive_users(self.workspace)
            if not rows:
                await self.client.send_private(qq, "当前没有停用成员。")
            else:
                await self.client.send_private(qq, "停用成员：\n" + "\n".join(f"{r['name']}｜QQ {r['qq']}" for r in rows))
            return

        if cmd in {"管理员添加", "管理员移除", "成员移除", "成员恢复"}:
            if not is_admin(user):
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
            if volunteer.has_appointment(self.workspace,target['qq']) and not volunteer.can_manage(self.workspace,qq):
                await self.client.send_private(qq,"普通管理员不能修改高级管理者的身份或成员状态。")
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

    async def handle_menu_session(self, qq: str, text: str) -> bool:
        key = (self.workspace, qq)
        session = menu_sessions.get(key)
        if session is None:
            return False
        stage, deadline = session
        if time.monotonic() > deadline:
            menu_sessions.pop(key, None)
            if text.isdecimal():
                await self.client.send_private(qq, "菜单已过期，请重新发送“帮助”。")
                return True
            return False
        if text == "取消":
            menu_sessions.pop(key, None)
            await self.client.send_private(qq, "已退出菜单选择。")
            return True
        if not text.isdecimal():
            menu_sessions.pop(key, None)
            return False
        if stage == "choose":
            choice = int(text) if len(text) <= 2 else 0
            if choice not in range(1, 8):
                await self.client.send_private(qq, "请回复 1～7，或发送“取消”退出。")
                return True
            if choice in {3, 4}:
                action = "提交" if choice == 3 else "更新进度"
                menu_sessions[key] = (action, time.monotonic() + 300)
                await self.client.send_private(qq, f"请发送要{action}的任务编号，例如：12。\n不知道编号可发送“我的任务”查看，再重新选择；发送“取消”退出。")
                return True
            action = {1: "我的待办", 2: "我的任务", 5: "我的志愿时长", 6: "我的信息", 7: "申请任务"}[choice]
        else:
            if len(text) > 18 or int(text) <= 0:
                await self.client.send_private(qq, "请输入有效的任务编号，或发送“取消”退出。")
                return True
            action = f"{stage} {int(text)}"
        menu_sessions.pop(key, None)
        await self.handle_private({"user_id": qq, "raw_message": action})
        return True

    async def handle_progress_session(self, qq: str, text: str):
        key = (self.workspace, qq)
        tid = progress_sessions[key]
        if text in {"取消", "取消更新"}:
            progress_sessions.pop(key, None)
            await self.client.send_private(qq, "已取消更新进度，原进度保留。")
            return
        result = await self.save_progress(qq, tid, text)
        if result not in {"empty", "too_long"}:
            progress_sessions.pop(key, None)

    async def save_progress(self, qq: str, tid: int, text: str):
        result = db.update_task_progress(self.workspace, tid, qq, text)
        messages = {
            "ok": f"任务 #{tid} 的个人进度已更新。\n{text}\n这不是正式提交；完成后请发送：提交 {tid}",
            "empty": "进度说明不能为空。",
            "too_long": "进度说明请控制在 1000 字以内。",
            "inactive": "当前成员资格无效，请联系负责人。",
            "not_found": "任务不存在。",
            "deleted": "任务已删除，不能更新进度。",
            "not_claimed": "你没有认领这个任务，只能更新自己的进度。",
            "submitted": f"该任务份额正在等待审核；如需继续修改，请先发送：撤回提交 {tid}",
            "completed": "该任务或你的任务份额已经完成，不能更新进度。",
        }
        await self.client.send_private(qq, messages.get(result, "更新失败，请稍后再试。"))
        return result

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
            task = db.get_task(self.workspace, task_id)
            submitter = db.get_user(self.workspace, qq, include_inactive=True)
            content = text if len(text) <= 1200 else text[:1200] + "…"
            name = submitter['name'] if submitter else qq
            notice = (f"【有新的任务提交待审核】\n任务 #{task_id}\n提交人：{name}（{qq}）\n"
                      f"成果说明：{content}\n\n审核：通过 {task_id} {qq} <小时数>\n"
                      f"暂不登记时长：通过 {task_id} {qq}\n打回 {task_id} {qq} <原因>")
            recipients = [self.owner_qq]
            publisher = str(task.get('publisher_qq') or '') if task else ''
            if publisher and publisher not in recipients:
                recipients.append(publisher)
            failed = []
            for recipient in recipients:
                if not await self.request_notice(recipient, notice):
                    failed.append(recipient)
            warning = ""
            if self.owner_qq in failed:
                warning = "\n总负责人提醒发送失败，请联系负责人查看“待审核”；成果已保存，不用重复提交。"
            elif failed:
                warning = "\n总负责人提醒已发送，但发布人提醒失败；成果已保存。"
            await self.request_notice(qq, f"任务 #{task_id} 已提交，当前状态：待审核。管理员审核通过后才计为完成。" + warning)
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

    async def notify_volunteer(self, task_id, target, action):
        recipients = list(dict.fromkeys([self.owner_qq] + [r['qq'] for r in volunteer.managers(self.workspace)]))
        text = "【志愿时长提醒】\n" + action
        if task_id is not None:
            with db.connect() as conn:
                row=conn.execute("SELECT id,units,service_date FROM volunteer_entries WHERE workspace=? AND task_id=? AND qq=?",(self.workspace,task_id,str(target))).fetchone()
            if row:
                value="待登记" if row['units'] is None else volunteer.hours(row['units'])+" 小时"
                text+=f"\n任务 #{task_id}｜成员 QQ {target}｜时长记录 #{row['id']}\n{row['service_date']}：{value}"
                if row['units'] is None:text+=f"\n补录：登记志愿时长 {task_id} {target} <小时数> {row['service_date']}"
        for recipient in recipients:
            if not await self.request_notice(recipient,text):
                print("[志愿时长] 通知失败，已保存的账目不回滚；请查询月报。")

    async def handle_volunteer_command(self, user, qq, cmd, arg):
        member_commands = {"我的志愿时长", "志愿时长明细"}
        admin_commands = {"待登记时长", "志愿时长统计", "导出志愿时长", "登记志愿时长", "更正志愿时长"}
        role_commands = {"时长负责人添加", "时长负责人移除", "时长负责人列表"}
        if cmd in role_commands:
            if cmd == "时长负责人列表":
                people=volunteer.managers(self.workspace)
                await self.client.send_private(qq,"志愿时长负责人：\n" + ("\n".join(f"{p['name']}（{p['qq']}）" for p in people) or "暂未任命；请由总负责人处理。"))
                return True
            if not volunteer.can_manage(self.workspace,qq):
                await self.client.send_private(qq,"只有总负责人和高级管理员可以任命或移除志愿时长负责人。")
                return True
            status,target=db.resolve_user(self.workspace,arg,active=None)
            if status!='ok':
                await self.client.send_private(qq,"用法："+cmd+" <姓名或QQ>；有同名时请使用 QQ。")
                return True
            try:
                volunteer.set_manager(self.workspace,qq,target['qq'],cmd=="时长负责人添加")
                await self.client.send_private(qq,"已更新志愿时长负责人设置。")
                await self.request_notice(target['qq'],"高级管理者已"+("授予你高级管理员职责，可使用任务管理、时长管理及高级任命指令。发送“帮助”查看指令。" if cmd=="时长负责人添加" else "移除你的志愿时长管理职责。"))
            except ValueError as exc:await self.client.send_private(qq,str(exc))
            return True
        if cmd not in member_commands | admin_commands:
            return False
        if cmd in admin_commands and not volunteer.can_manage(self.workspace,qq):
            await self.client.send_private(qq, "只有总负责人或高级管理员可以操作全体志愿时长。")
            return True
        try:
            parts = arg.split()
            if cmd == "登记志愿时长":
                if len(parts) not in (3,4) or parse_task_id(parts[0]) is None:
                    raise ValueError("用法：登记志愿时长 <任务编号> <姓名或QQ> <小时数> [YYYY-MM-DD]")
                status,target = db.resolve_user(self.workspace,parts[1],active=None)
                if status != 'ok':
                    raise ValueError("成员不存在或有同名，请使用准确的 QQ 号。")
                volunteer.register(self.workspace,int(parts[0].lstrip('#')),target['qq'],qq,parts[2],parts[3] if len(parts)==4 else None)
                await self.notify_volunteer(parse_task_id(parts[0]),target["qq"],"已补录志愿时长。")
                await self.client.send_private(qq,"志愿时长已登记。" + ("已归入指定日期所在月份。" if len(parts)==4 else "已归入今天所在月份。"))
                return True
            if cmd == "更正志愿时长":
                pieces=arg.split(maxsplit=3)
                if len(pieces)!=4 or not pieces[0].isdigit():
                    raise ValueError("用法：更正志愿时长 <记录编号> <新总小时数> <YYYY-MM-DD> <原因>")
                volunteer.correct(self.workspace,int(pieces[0]),qq,pieces[1],pieces[2],pieces[3])
                await self.notify_volunteer(None,None,f"志愿时长记录 #{pieces[0]} 已更正；可导出相关月份核对修改记录。")
                await self.client.send_private(qq,"时长及归属日期已更正，原值与原因保留在修改记录中。")
                return True
            month=parts[0] if parts else None
            paged=cmd in {"志愿时长统计","志愿时长明细","待登记时长"}
            if len(parts)>(2 if paged else 1):
                raise ValueError(f"用法：{cmd} [YYYY-MM]" + (" [页码]" if paged else ""))
            page=int(parts[1]) if len(parts)==2 else 1
            if page<1 or page>1000000:raise ValueError("页码应为正整数。")
            if cmd == "导出志愿时长":
                path=volunteer.export_month(self.workspace,month,BASE_DIR)
                try:
                    resp=await self.client.send_private_file(qq,path)
                    if not await self.api_ok(resp):raise RuntimeError("file upload failed")
                except Exception:
                    await self.client.send_private(qq,f"月报已生成，但发送失败。文件保存在：{path}")
                else:
                    await self.client.send_private(qq,f"已导出：{path.name}。历史数据仍保留，未清零。")
                return True
            m,start,end,people,details=volunteer.report(self.workspace,month,qq if cmd in member_commands else None)
            total=sum(r['units'] for r in people)
            pending=sum(r['pending'] for r in people)
            lines=[f"{m} 志愿时长｜从 {start} 起",f"已登记：{volunteer.hours(total)} 小时；待登记：{pending} 份（不计入已登记时长）"]
            if cmd=="我的志愿时长":
                lines.append(f"查看逐笔记录：志愿时长明细 {m}")
            elif cmd in {"志愿时长明细","待登记时长"}:
                if cmd=="待登记时长":details=[r for r in details if r["units"] is None]
                rows=details[(page-1)*10:page*10]
                lines.append(f"第 {page} 页，每页 10 条")
                for r in rows:
                    value='待登记' if r['units'] is None else volunteer.hours(r['units'])+' 小时'
                    lines.append(f"记录 #{r['id']}｜任务 #{r['task_id']}｜{r['name']}（{r['qq']}）｜{r['service_date']}｜{value}")
                if not rows:lines.append("本页没有记录。")
                if len(details)>page*10:lines.append(f"下一页：{cmd} {m} {page+1}")
            else:
                rows=people[(page-1)*20:page*20]
                lines.append(f"第 {page} 页，每页 20 人")
                for r in rows:lines.append(f"{r['name']}（{r['qq']}）：{volunteer.hours(r['units'])} 小时；待登记 {r['pending']} 份")
                if not rows:lines.append("本页没有成员。")
                if len(people)>page*20:lines.append(f"下一页：{cmd} {m} {page+1}")
                lines.append(f"查看记录编号和完整明细：导出志愿时长 {m}")
            await self.client.send_private(qq,"\n".join(lines))
        except ValueError as exc:
            await self.client.send_private(qq,str(exc))
        return True

    async def request_notice(self, qq, text):
        try:
            await self.client.send_private(qq, text)
            return True
        except Exception as exc:
            print(f"[私聊通知] 发送失败：{type(exc).__name__}；已保存的数据仍保留。")
            return False

    async def handle_request_session(self, qq, text):
        key = (self.workspace, qq)
        if not db.get_user(self.workspace, qq):
            request_sessions.pop(key, None)
            await self.client.send_private(qq, "成员资格已失效，本次填写已取消。")
            return
        if text == "取消":
            request_sessions.pop(key, None)
            await self.client.send_private(qq, "已取消填写，未提交任务申请。")
            return
        values = request_sessions[key]
        if len(values) < 5:
            try:
                value = db.validate_task_request_field(len(values), text)
            except ValueError as exc:
                await self.client.send_private(qq, str(exc))
                return
            values.append(value)
            if len(values) < 5:
                await self.client.send_private(qq, REQUEST_PROMPTS[len(values)])
            else:
                await self.client.send_private(qq,
                    f"请确认申请：\n类型：{values[0]}\n学科：{values[1]}\n任务内容及所需知识：{values[2]}\n"
                    f"人数：{values[3]}\n理由：{values[4]}\n\n"
                    "发送“确认”提交审批，或“取消”。批准后部门公开发布任务，不自动由你认领。")
            return
        if text != "确认":
            await self.client.send_private(qq, "请发送“确认”提交申请，或“取消”退出。")
            return
        try:
            rid = db.create_task_request(self.workspace, qq, values)
        except ValueError as exc:
            request_sessions.pop(key, None)
            await self.client.send_private(qq, str(exc))
            return
        # Drop the session before network IO so a repeated confirmation cannot resubmit it.
        request_sessions.pop(key, None)
        row = db.get_task_request(self.workspace, rid)
        failed = False
        for admin in db.list_admins(self.workspace):
            ok = await self.request_notice(admin['qq'], "【新的任务发布申请】\n" + request_summary(row) +
                f"\n\n批准申请 {rid}\n驳回申请 {rid} 原因")
            failed = failed or not ok
        await self.request_notice(qq, f"申请 #{rid} 已保存，等待负责人或管理员审批。尚未发布为任务。\n发送“我的申请”查看状态。" +
            ("\n部分审批提醒发送失败，请联系管理员查看“待审批申请”；不用重复提交。" if failed else ""))

    async def handle_request_command(self, user, qq, cmd, arg):
        commands = {"申请任务","我的申请","申请详情","撤回申请","待审批申请","批准申请","驳回申请"}
        if cmd not in commands:
            return False
        if cmd in {"待审批申请","批准申请","驳回申请"} and not is_admin(user):
            await self.client.send_private(qq, "只有负责人和管理员可以审批任务发布申请。")
            return True
        if cmd == "申请任务":
            request_sessions[(self.workspace, qq)] = []
            await self.client.send_private(qq, "开始申请部门发布任务。批准后公开认领，不自动分配给你。\n" + REQUEST_PROMPTS[0] + "\n发送“取消”可退出。")
            return True
        if cmd in {"我的申请","待审批申请"}:
            try:
                page = int(arg) if arg else 1
                if not 1 <= page <= 1000000:
                    raise ValueError()
            except ValueError:
                await self.client.send_private(qq, f"用法：{cmd} [正整数页码]")
                return True
            rows = db.list_task_requests(self.workspace, qq if cmd == "我的申请" else None, cmd == "待审批申请", page)
            chunks = [f"{cmd}｜第 {page} 页（每页 10 条）"]
            for row in rows:
                chunks.append(f"申请 #{row['id']}｜{REQUEST_STATUS[row['status']]}｜{row['applicant_name']}｜{row['task_type']}｜{row['subject']}" +
                    (f"｜正式任务 #{row['task_id']}" if row['task_id'] else ""))
            chunks.append("查看完整内容：申请详情 <申请编号>" if rows else "本页没有申请。")
            if len(rows) == 10:
                chunks.append(f"下一页：{cmd} {page+1}")
            await self.client.send_private(qq, "\n\n".join(chunks))
            return True
        parts = arg.split(maxsplit=1)
        rid = parse_task_id(parts[0]) if parts else None
        if rid is None or (cmd == "驳回申请" and len(parts) < 2) or (cmd != "驳回申请" and len(parts) != 1):
            await self.client.send_private(qq, f"用法：{cmd} <申请编号>" + (" <原因>" if cmd == "驳回申请" else ""))
            return True
        row = db.get_task_request(self.workspace, rid)
        if not row or (not is_admin(user) and row['applicant_qq'] != qq):
            await self.client.send_private(qq, "没有找到你可以查看的申请。")
            return True
        if cmd == "申请详情":
            await self.client.send_private(qq, request_summary(row))
            return True
        decision = {"批准申请":"approved","驳回申请":"rejected","撤回申请":"withdrawn"}[cmd]
        result, tid = db.decide_task_request(self.workspace, rid, qq, decision, parts[1] if len(parts) > 1 else "")
        if result != "ok":
            messages = {"forbidden":"没有权限执行该操作；撤回只能由申请人本人操作。",
                "not_found":"申请不存在。", "already_processed":"申请已处理，不能重复审批或撤回。",
                "reason_required":"驳回时必须填写原因。", "reason_too_long":"驳回原因不能超过 1000 字。"}
            await self.client.send_private(qq, messages.get(result,"操作未完成。"))
            return True
        updated = db.get_task_request(self.workspace, rid)
        await self.request_notice(qq, "处理已保存。\n" + request_summary(updated))
        if decision != 'withdrawn' and row['applicant_qq'] != qq:
            if not await self.request_notice(row['applicant_qq'], "你的任务发布申请已有结果：\n" + request_summary(updated)):
                await self.request_notice(qq, "审批已保存，但申请人私聊通知失败；对方可用“我的申请”查询。")
        if tid:
            task = db.get_task(self.workspace, tid)
            announcement = (f"【新任务 #{tid}】\n任务类型：{task['task_type']}\n涉及学科：{task['subject']}\n"
                f"所需知识：{task['required_knowledge']}\n人数需求：{task['required_people']} 人\n发布人：{task['publisher_name']}\n\n"
                f"需要接取的成员请私聊我发送：认领 {tid}")
            try:
                await self.client.send_group(self.workspace, announcement)
            except Exception as exc:
                print(f"[任务申请] 任务 #{tid} 公告失败：{type(exc).__name__}")
                await self.request_notice(qq, f"任务 #{tid} 已创建并开放认领，但群公告发送失败。请手动通知成员，不要重复批准。")
        return True

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
