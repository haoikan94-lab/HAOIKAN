import json
import os
from datetime import datetime, timedelta, time
import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, JobQueue
from zoneinfo import ZoneInfo

# ================== 配置区 ==================
TOKEN = os.getenv("BOT_TOKEN")

if not TOKEN:
    raise ValueError("❌ 请在 Railway 设置环境变量 BOT_TOKEN")

print(f"✅ Bot Token 已成功加载 | 长度: {len(TOKEN)}")

# ================== Railway Volume 持久化配置 ==================
DATA_PATH = os.getenv("DATA_PATH", "/data")

if not os.path.exists(DATA_PATH):
    os.makedirs(DATA_PATH)

DATA_FILE = os.path.join(DATA_PATH, "group_attendance.json")
EXCEL_FOLDER = os.path.join(DATA_PATH, "excel_files")

if not os.path.exists(EXCEL_FOLDER):
    os.makedirs(EXCEL_FOLDER)

# ================== 北京时间设置 ==================
TZ = ZoneInfo("Asia/Shanghai")

def beijing_now():
    return datetime.now(TZ)

def beijing_date_str():
    return beijing_now().strftime("%Y-%m-%d")

def beijing_time_str():
    return beijing_now().strftime("%H:%M:%S")

# ================== ACTIONS ==================
ACTIONS = {
    "1": {"name": "第一班上班", "time": "09:00", "is_work": True,  "type": "work", "early_allowed": "08:00"},
    "2": {"name": "第一班下班", "time": "15:00", "is_work": False, "type": "work", "max_time": "15:59"},
    "3": {"name": "第二班上班", "time": "17:00", "is_work": True,  "type": "work", "early_allowed": "16:00"},
    "4": {"name": "第二班下班", "time": "23:00", "is_work": False, "type": "work", "max_time": "23:59"},
    "5": {"name": "开始休息",   "time": None,    "is_work": False, "type": "rest_start"},
    "6": {"name": "结束休息",   "time": None,    "is_work": False, "type": "rest_end"},
}

# ================== 数据读写 ==================
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# ================== 管理员配置 ==================
def get_admins(chat_id: str):
    data = load_data()
    return data.setdefault(chat_id, {}).setdefault("admins", [])

def save_admin(chat_id: str, admin_id: str):
    data = load_data()
    admins = data.setdefault(chat_id, {}).setdefault("admins", [])
    if admin_id not in admins:
        admins.append(admin_id)
        save_data(data)
        return True
    return False

def remove_admin(chat_id: str, admin_id: str):
    data = load_data()
    admins = data.setdefault(chat_id, {}).setdefault("admins", [])
    if admin_id in admins:
        admins.remove(admin_id)
        save_data(data)
        return True
    return False

# ================== 自动清理 ==================
async def cleanup_old_data(context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    cutoff_date = (beijing_now() - timedelta(days=90)).strftime("%Y-%m-%d")
    cleaned = 0
    for chat_id in list(data.keys()):
        for user_id in list(data[chat_id].get("users", {}).keys()):
            user_info = data[chat_id]["users"][user_id]
            if "records" in user_info:
                old_dates = [d for d in list(user_info["records"].keys()) if d < cutoff_date]
                for d in old_dates:
                    del user_info["records"][d]
                    cleaned += 1
    if cleaned > 0:
        save_data(data)
        print(f"🧹 已清理 {cleaned} 条90天前旧记录")

# ================== 获取群主 ==================
async def get_group_owner(context: ContextTypes.DEFAULT_TYPE, chat_id: int):
    try:
        admins = await context.bot.get_chat_administrators(chat_id)
        for admin in admins:
            if admin.status == "creator":
                return admin.user.id
    except Exception as e:
        print(f"获取群主失败 {chat_id}: {e}")
    return None

# ================== 自动发送日报表 ==================
async def send_daily_report(context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    today_str = beijing_date_str()
    
    print(f"📊 开始生成 {today_str} 全群日报表...")

    for chat_id_str, chat_data in data.items():
        chat_id = int(chat_id_str)
        recipients = set()
        
        owner_id = await get_group_owner(context, chat_id)
        if owner_id:
            recipients.add(owner_id)
        
        admins = get_admins(chat_id_str)
        recipients.update(int(uid) for uid in admins)
        
        if not recipients:
            print(f"⚠️ 群 {chat_id} 没有接收人，跳过")
            continue

        rows = []
        users = chat_data.get("users", {})
        has_record = False

        for user_info in users.values():
            records = user_info.get("records", {}).get(today_str, [])
            if not records:
                continue

            has_record = True
            rest_count = sum(1 for r in records if r.get("type") == "rest_end")
            total_rest = sum(r.get("rest_minutes", 0) for r in records if r.get("type") == "rest_end")

            for r in records:
                if r.get("type") in ["rest_start", "rest_end"]:
                    continue
                rows.append({
                    "姓名": user_info.get("name", "未知"),
                    "日期": today_str,
                    "班次": ACTIONS.get(r.get("action"), {}).get("name", r.get("action")),
                    "打卡时间": r["time"],
                    "状态": r.get("display", ""),
                    "迟到分钟": r.get("late_min", 0),
                    "休息次数": rest_count,
                    "总休息分钟": total_rest
                })

        filename = f"全群打卡日报_{today_str}.xlsx"
        filepath = os.path.join(EXCEL_FOLDER, filename)

        try:
            if rows:
                pd.DataFrame(rows).to_excel(filepath, index=False)
                caption = f"📊 **{today_str} 全群打卡日报** 已生成\n共 {len(set(r['姓名'] for r in rows))} 人有打卡记录"
            else:
                empty_df = pd.DataFrame(columns=["姓名","日期","班次","打卡时间","状态","迟到分钟","休息次数","总休息分钟"])
                empty_df.to_excel(filepath, index=False)
                caption = f"📊 **{today_str} 全群打卡日报**\n\n今日暂无打卡记录"

            sent_count = 0
            for recipient_id in recipients:
                try:
                    with open(filepath, 'rb') as f:
                        await context.bot.send_document(
                            chat_id=recipient_id,
                            document=f,
                            filename=filename,
                            caption=caption,
                            parse_mode="Markdown"
                        )
                    sent_count += 1
                except Exception as e:
                    print(f"发送日报失败 {recipient_id} (群{chat_id}): {e}")

            print(f"✅ 已成功发送 {today_str} 日报表给群 {chat_id} 的 {sent_count}/{len(recipients)} 位接收人")
        except Exception as e:
            print(f"❌ 生成或发送日报表失败 (群{chat_id}): {e}")

# ================== 辅助函数 ==================
def get_late_minutes(expected_time: str) -> int:
    if not expected_time:
        return 0
    now = beijing_now()
    expected = datetime.strptime(expected_time, "%H:%M").replace(
        year=now.year, month=now.month, day=now.day, tzinfo=TZ)
    if now > expected:
        return int((now - expected).total_seconds() / 60)
    return 0

def is_valid_checkin_time(shift: str) -> tuple[bool, str]:
    action = ACTIONS.get(shift)
    if not action:
        return True, ""
    
    now = beijing_now()
    now_time = now.strftime("%H:%M")

    if shift == "1" and now_time > "15:00":
        return False, "❌ 15:00 之后不能再打第一班上班（1）"
    if shift == "3" and now_time > "23:00":
        return False, "❌ 23:00 之后不能再打第二班上班（3）"

    if "early_allowed" in action:
        early_time = datetime.strptime(action["early_allowed"], "%H:%M").replace(
            year=now.year, month=now.month, day=now.day, tzinfo=TZ)
        if now < early_time:
            return False, f"❌ {action['name']} 最早 {action['early_allowed']} 才能打卡！"

    if "max_time" in action:
        max_time = datetime.strptime(action["max_time"], "%H:%M").replace(
            year=now.year, month=now.month, day=now.day, tzinfo=TZ)
        if now > max_time:
            return False, f"❌ {action['name']} 必须在 {action['max_time']} 前打卡！"

    return True, ""

def calculate_rest_duration(start_time_str: str, end_time_str: str) -> int:
    try:
        start = datetime.strptime(start_time_str, "%H:%M:%S")
        end = datetime.strptime(end_time_str, "%H:%M:%S")
        if end < start:
            end += timedelta(days=1)
        return int((end - start).total_seconds() / 60)
    except:
        return 0

# ================== 核心打卡函数 ==================
async def daka(update: Update, context: ContextTypes.DEFAULT_TYPE, shift: str):
    chat_id = str(update.effective_chat.id)
    user = update.effective_user
    user_id = str(user.id)

    # 所有打卡都自动注册
    await auto_register(update, context)

    now = beijing_now()
    date_str = beijing_date_str()
    time_str = beijing_time_str()

    valid, error_msg = is_valid_checkin_time(shift)
    if not valid:
        await update.message.reply_text(error_msg)
        return

    data = load_data()
    chat_data = data.setdefault(chat_id, {"registered": {}, "users": {}})
    
    user_data = chat_data["users"].setdefault(user_id, {"name": user.full_name, "records": {}})
    records = user_data["records"].setdefault(date_str, [])

    is_resting = any(r.get("type") == "rest_start" and "rest_minutes" not in r for r in records)

    if shift == "5" and is_resting:
        await update.message.reply_text("⏳ 你当前正在休息中，请先输入 **6** 结束休息后再开始新休息")
        return
    if shift == "6" and not is_resting:
        await update.message.reply_text("⚠️ 请先输入 **5** 开始休息，才能输入 6 结束休息")
        return

    if shift in ["1", "2", "3", "4"]:
        if any(r.get("action") == shift for r in records):
            action_name = ACTIONS.get(shift, {}).get("name", shift)
            await update.message.reply_text(f"⚠️ 今天已经打过 **{action_name}** 了，不能重复打卡")
            return

    action_info = ACTIONS.get(shift, {"name": f"未知{shift}", "type": "unknown"})
    late_min = get_late_minutes(action_info.get("time")) if action_info.get("is_work") else 0

    display = action_info["name"]
    rest_minutes = 0

    if shift == "6" and is_resting:
        for r in reversed(records):
            if r.get("type") == "rest_start" and "rest_minutes" not in r:
                rest_minutes = calculate_rest_duration(r["time"], time_str)
                r["rest_minutes"] = rest_minutes
                display = f"结束休息（休息{rest_minutes}分钟）"
                break

    if shift == "5":
        display = "开始休息"

    record = {
        "time": time_str,
        "action": shift,
        "display": display,
        "late_min": late_min,
        "type": action_info.get("type", "unknown")
    }
    if rest_minutes > 0:
        record["rest_minutes"] = rest_minutes

    records.append(record)
    save_data(data)

    emoji = "⚠️" if late_min > 0 else "✅"
    late_text = f"（迟到{late_min}分钟）" if late_min > 0 else ""
    await update.message.reply_text(
        f"{emoji} **{user.full_name}** {display}{late_text}\n日期：{date_str}\n时间：{time_str}",
        parse_mode="Markdown"
    )

async def text_daka(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    mapping = {
        "1": "1", "上班": "1", "上午": "1",
        "2": "2", "下班1": "2", "下1": "2", "下班": "2",
        "3": "3", "下午上班": "3", "上班2": "3",
        "4": "4", "下班2": "4", "下2": "4",
        "5": "5", "休息": "5", "开始休息": "5",
        "6": "6", "结束休息": "6", "回岗": "6", "结束": "6"
    }
    if text in mapping:
        await daka(update, context, mapping[text])

async def auto_register(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = str(update.effective_chat.id)
    user_id = str(update.effective_user.id)
    user_name = update.effective_user.full_name

    data = load_data()
    registered = data.setdefault(chat_id, {}).setdefault("registered", {})
    
    if user_id not in registered:
        registered[user_id] = user_name
        data[chat_id].setdefault("users", {}).setdefault(user_id, {"name": user_name, "records": {}})
        save_data(data)
        await update.message.reply_text(f"✅ **{user_name}** 自动注册成功！", parse_mode="Markdown")

# ================== 其他命令 ==================
async def register(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = str(update.effective_chat.id)
    user_id = str(update.effective_user.id)
    user_name = update.effective_user.full_name

    data = load_data()
    chat_data = data.setdefault(chat_id, {"registered": {}, "users": {}})
    if user_id not in chat_data["registered"]:
        chat_data["registered"][user_id] = user_name
        chat_data["users"].setdefault(user_id, {"name": user_name, "records": {}})
        save_data(data)
        await update.message.reply_text(f"✅ **{user_name}** 注册成功！", parse_mode="Markdown")
    else:
        await update.message.reply_text("✅ 你已经注册过了。")

async def add_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    chat_id = str(update.effective_chat.id)
    owner_id = await get_group_owner(context, int(chat_id))
    if owner_id and owner_id != update.effective_user.id:
        await update.message.reply_text("⚠️ 仅群主可添加/删除管理员")
        return

    if not context.args:
        await update.message.reply_text("用法：/addadmin 用户ID")
        return

    target_id = context.args[0].strip()
    if save_admin(chat_id, target_id):
        await update.message.reply_text(f"✅ 已添加管理员：`{target_id}`", parse_mode="Markdown")
    else:
        await update.message.reply_text("✅ 该用户已是管理员")

async def del_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    chat_id = str(update.effective_chat.id)
    owner_id = await get_group_owner(context, int(chat_id))
    if owner_id and owner_id != update.effective_user.id:
        await update.message.reply_text("⚠️ 仅群主可添加/删除管理员")
        return

    if not context.args:
        await update.message.reply_text("用法：/deladmin 用户ID")
        return

    target_id = context.args[0].strip()
    if remove_admin(chat_id, target_id):
        await update.message.reply_text(f"✅ 已删除管理员：`{target_id}`", parse_mode="Markdown")
    else:
        await update.message.reply_text("❌ 该用户不是管理员")

async def adminlist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    chat_id = str(update.effective_chat.id)
    admins = get_admins(chat_id)
    owner_id = await get_group_owner(context, int(chat_id))
    
    text = "📋 **本群日报表接收人列表**\n\n"
    if owner_id:
        text += f"👑 群主: `{owner_id}`\n"
    text += f"📌 指定管理员（{len(admins)}人）:\n"
    for i, aid in enumerate(admins, 1):
        text += f"{i}. `{aid}`\n"
    await update.message.reply_text(text, parse_mode="Markdown")

async def myrecord(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.type != "private":
        await update.message.reply_text("此命令仅支持私聊使用")
        return

    user_id = str(update.effective_user.id)
    data = load_data()
    records_found = False
    text = f"📋 **{update.effective_user.full_name}** 的打卡记录\n\n"

    for chat_id, chat_data in data.items():
        user_data = chat_data.get("users", {}).get(user_id)
        if not user_data or not user_data.get("records"):
            continue
        records_found = True
        text += f"**群组ID：** `{chat_id}`\n"
        for date, recs in sorted(user_data["records"].items(), reverse=True)[:15]:
            text += f"**{date}**\n"
            for r in recs:
                if r.get("type") in ["rest_start", "rest_end"]:
                    continue
                late = f"（迟到{r.get('late_min',0)}分）" if r.get("late_min", 0) > 0 else ""
                text += f"• {r['display']}{late} {r['time']}\n"
        text += "\n"

    if not records_found:
        await update.message.reply_text("暂无打卡记录")
    else:
        await update.message.reply_text(text, parse_mode="Markdown")

async def is_admin(update: Update) -> bool:
    if update.effective_chat.type == "private":
        return True
    try:
        member = await update.effective_chat.get_member(update.effective_user.id)
        return member.status in ["administrator", "creator"]
    except:
        return False

async def registered_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = str(update.effective_chat.id)
    data = load_data()
    registered = data.get(chat_id, {}).get("registered", {})
    if not registered:
        await update.message.reply_text("📋 本群暂无已注册人员。")
        return
    text = f"📋 **本群已注册人员名单**（共 {len(registered)} 人）\n\n"
    for i, (uid, name) in enumerate(registered.items(), 1):
        text += f"{i}. {name} (ID: `{uid}`)\n"
    await update.message.reply_text(text, parse_mode="Markdown")

async def deluser(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 此命令仅限管理员使用")
        return
    if not context.args:
        await update.message.reply_text("用法：/deluser @用户名 或 用户ID")
        return

    chat_id = str(update.effective_chat.id)
    target = context.args[0].strip()
    data = load_data()
    registered = data.get(chat_id, {}).get("registered", {})
    target_id = None

    if target.startswith('@'):
        search_name = target[1:].lower().replace(" ", "")
        for uid, name in registered.items():
            if name.lower().replace(" ", "") == search_name:
                target_id = uid
                break
    elif target in registered:
        target_id = target

    if target_id and target_id in registered:
        name = registered.pop(target_id)
        data.get(chat_id, {}).get("users", {}).pop(target_id, None)
        save_data(data)
        await update.message.reply_text(f"✅ 已删除用户：**{name}**", parse_mode="Markdown")
    else:
        await update.message.reply_text("❌ 未找到该用户。")

async def delete_record(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    if not context.args:
        await update.message.reply_text("用法：/del YYYY-MM-DD")
        return
    date_to_del = context.args[0].strip()
    chat_id = str(update.effective_chat.id)
    data = load_data()
    count = 0
    for user_data in data.get(chat_id, {}).get("users", {}).values():
        if date_to_del in user_data.get("records", {}):
            del user_data["records"][date_to_del]
            count += 1
    save_data(data)
    await update.message.reply_text(f"✅ 已删除 {date_to_del} 的所有记录（影响 {count} 人）")

async def todayexcel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    chat_id = str(update.effective_chat.id)
    data = load_data()
    today_str = beijing_date_str()
    filename = f"全群打卡_{today_str}.xlsx"
    filepath = os.path.join(EXCEL_FOLDER, filename)

    rows = []
    for user_info in data.get(chat_id, {}).get("users", {}).values():
        if today_str not in user_info.get("records", {}):
            continue
        records = user_info["records"][today_str]
        rest_count = sum(1 for r in records if r.get("type") == "rest_end")
        total_rest = sum(r.get("rest_minutes", 0) for r in records if r.get("type") == "rest_end")

        for r in records:
            if r.get("type") in ["rest_start", "rest_end"]:
                continue
            rows.append({
                "姓名": user_info["name"],
                "日期": today_str,
                "班次": ACTIONS.get(r.get("action"), {}).get("name", r.get("action")),
                "打卡时间": r["time"],
                "状态": r.get("display", ""),
                "迟到分钟": r.get("late_min", 0),
                "休息次数": rest_count,
                "总休息分钟": total_rest
            })

    if rows:
        pd.DataFrame(rows).to_excel(filepath, index=False)
        with open(filepath, 'rb') as f:
            await update.message.reply_document(
                f, 
                filename=filename, 
                caption=f"✅ {today_str} 全群打卡报表"
            )
    else:
        await update.message.reply_text("今日暂无打卡记录")

async def monthexcel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    chat_id = str(update.effective_chat.id)
    data = load_data()
    current_month = beijing_now().strftime("%Y-%m")
    filename = f"全群打卡_{current_month}.xlsx"
    filepath = os.path.join(EXCEL_FOLDER, filename)

    rows = []
    for user_info in data.get(chat_id, {}).get("users", {}).values():
        for date_str, records in user_info.get("records", {}).items():
            if not date_str.startswith(current_month):
                continue
            rest_count = sum(1 for r in records if r.get("type") == "rest_end")
            total_rest = sum(r.get("rest_minutes", 0) for r in records if r.get("type") == "rest_end")
            for r in records:
                if r.get("type") in ["rest_start", "rest_end"]:
                    continue
                rows.append({
                    "姓名": user_info["name"],
                    "日期": date_str,
                    "班次": ACTIONS.get(r.get("action"), {}).get("name", r.get("action")),
                    "打卡时间": r["time"],
                    "状态": r.get("display", ""),
                    "迟到分钟": r.get("late_min", 0),
                    "休息次数": rest_count,
                    "总休息分钟": total_rest
                })

    if rows:
        pd.DataFrame(rows).to_excel(filepath, index=False)
        with open(filepath, 'rb') as f:
            await update.message.reply_document(
                f, 
                filename=filename, 
                caption=f"✅ {current_month} 全群打卡月报表"
            )
    else:
        await update.message.reply_text(f"{current_month} 暂无打卡记录")

async def absent(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    chat_id = str(update.effective_chat.id)
    today_str = beijing_date_str()
    data = load_data()
    registered = data.get(chat_id, {}).get("registered", {})
    users = data.get(chat_id, {}).get("users", {})

    incomplete = []
    main_shifts = {"1", "2", "3", "4"}

    for uid, name in registered.items():
        records = users.get(uid, {}).get("records", {}).get(today_str, [])
        done = {r["action"] for r in records if r["action"] in main_shifts}
        if done != main_shifts:
            done_str = ",".join(sorted(done)) if done else "无"
            incomplete.append(f"{name} → 已打: {done_str}")

    if not incomplete:
        await update.message.reply_text(f"🎉 今天所有人均已完成全部打卡！")
    else:
        text = f"📋 **今日未完成全部打卡人员** ({len(incomplete)}/{len(registered)})\n\n"
        for i, item in enumerate(incomplete, 1):
            text += f"{i}. {item}\n"
        await update.message.reply_text(text, parse_mode="Markdown")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "飞机的代号确定下来了就不要再改了，否则会打卡记录失败\n\n"
        "第一班上班打1，下班打2。第二班上班打3，下班打4，离开工位休息打5，回来打6。\n\n"
        "上下班打卡的，迟到早退相同，10分钟内扣50，1小时内扣100，1小时外按旷工扣200.超过1秒也算迟到跟早退。漏打每次100（下班打卡有效时间1小时）.\n"
        "严禁互相打卡与飞机定时发送。互相打卡两个人各扣300，定时发送扣600.\n"
        "如果遇到没有信号的缘故，或者帮公司做其他事情没办法及时打卡的，找公司组长或副组长证明补打卡后原因写上。其余不管是加班聊客户或者其他原因的也算迟到早退。\n\n"
        "私聊发送 /myrecord 可查询个人记录\n"
    )

# ================== 主程序 ==================
def main():
    app = Application.builder().token(TOKEN).build()
    job_queue: JobQueue = app.job_queue

    job_queue.run_daily(cleanup_old_data, time=time(23, 59, 59, tzinfo=TZ))
    job_queue.run_daily(send_daily_report, time=time(23, 59, 59, tzinfo=TZ))

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("register", register))
    app.add_handler(CommandHandler("registered", registered_list))
    app.add_handler(CommandHandler("deluser", deluser))
    app.add_handler(CommandHandler("todayexcel", todayexcel))
    app.add_handler(CommandHandler("monthexcel", monthexcel))
    app.add_handler(CommandHandler("absent", absent))
    app.add_handler(CommandHandler("del", delete_record))
    app.add_handler(CommandHandler("myrecord", myrecord))
    app.add_handler(CommandHandler("addadmin", add_admin))
    app.add_handler(CommandHandler("deladmin", del_admin))
    app.add_handler(CommandHandler("adminlist", adminlist))

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_daka))

    print("🤖 打卡机器人已成功启动 | 每日报表自动发送功能已启用（北京时间23:59:59）")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
