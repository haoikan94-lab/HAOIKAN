import json
import os
import asyncio
from datetime import datetime, timedelta, time
import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, JobQueue
from zoneinfo import ZoneInfo
import time as time_module
import shutil

# ================== 配置区 ==================
TOKEN = os.getenv("BOT_TOKEN")
if not TOKEN:
    raise ValueError("❌ 请在 Railway 设置环境变量 BOT_TOKEN")

print(f"✅ Bot Token 已成功加载 | 长度: {len(TOKEN)}")

# ================== Railway Volume 配置 ==================
DATA_PATH = os.getenv("DATA_PATH", "/data")
os.makedirs(DATA_PATH, exist_ok=True)

DATA_FILE = os.path.join(DATA_PATH, "group_attendance.json")
EXCEL_FOLDER = os.path.join(DATA_PATH, "excel_files")
os.makedirs(EXCEL_FOLDER, exist_ok=True)

# ================== 北京时间 ==================
TZ = ZoneInfo("Asia/Shanghai")

def beijing_now():
    return datetime.now(TZ)

def beijing_date_str(dt=None):
    if dt is None:
        dt = beijing_now()
    return dt.strftime("%Y-%m-%d")

# ================== 新日期边界逻辑：05:00 为分界点 ==================
def get_attendance_date(now=None):
    """获取考勤所属日期（05:00 之后为当天）"""
    if now is None:
        now = beijing_now()
    if now.hour < 5:
        return beijing_date_str(now - timedelta(days=1))
    return beijing_date_str(now)

def get_record_date(shift: str, now=None) -> str:
    """根据打卡类型获取正确的记录日期 - 已修复 Shift 4 跨天逻辑"""
    if now is None:
        now = beijing_now()
    
    base_date = get_attendance_date(now)
    
    # Shift 4（第二班下班）在 05:00 之前打卡时，始终属于前一天
    if shift == "4" and now.hour < 5:
        return beijing_date_str(now - timedelta(days=1))
    
    return base_date

# ================== 前一天考勤日期（用于05:30自动日报） ==================
def get_previous_attendance_date(now=None) -> str:
    if now is None:
        now = beijing_now()
    return beijing_date_str(now - timedelta(days=1))

def get_report_date_for_daily() -> str:
    return get_previous_attendance_date(beijing_now())

# ================== 时间有效性检查 ==================
def is_valid_checkin_time(shift: str, now: datetime = None) -> tuple[bool, str]:
    """打卡时间有效性检查"""
    if shift not in {"1", "2", "3", "4"}:
        return True, ""
    
    if now is None:
        now = beijing_now()
    
    current_time = now.time()
    
    if shift == "1":      # 第一班上班
        if current_time < time(13, 0):
            return False, "⚠️ 第一班上班需在 **13:00之后** 打卡"
    elif shift == "2":    # 第一班下班
        if current_time >= time(20, 0):
            return False, "⚠️ 第一班下班需在 **20:00之前** 打卡"
    elif shift == "3":    # 第二班上班
        if current_time < time(20, 0):
            return False, "⚠️ 第二班上班需在 **20:00之后** 打卡"
    elif shift == "4":    # 第二班下班
        if current_time >= time(5, 0):
            return False, "⚠️ 第二班下班需在 **05:00之前** 打卡（00:00-05:00）"
    
    return True, ""

def is_valid_rest_time(shift: str) -> tuple[bool, str]:
    if shift not in {"5", "7"}:
        return True, ""
    
    now = beijing_now()
    current_time = now.time()
    
    if time(14, 0) <= current_time < time(19, 0):
        return True, ""
    if current_time >= time(21, 0) or current_time < time(5, 0):
        return True, ""
    
    return False, "⚠️ 休息/暂离（5或7）只能在以下工作时段打卡：\n• 第一班 14:00-19:00\n• 第二班 21:00-05:00"

def calculate_rest_duration(start_time_str: str, end_time_str: str) -> int:
    """计算休息/暂离时长（分钟），支持跨天"""
    try:
        fmt = "%H:%M:%S"
        start = datetime.strptime(start_time_str, fmt)
        end = datetime.strptime(end_time_str, fmt)
        
        if end < start:
            end += timedelta(days=1)
        
        delta = end - start
        return int(delta.total_seconds() / 60)
    except Exception as e:
        print(f"⚠️ 计算休息时长失败: {e}")
        return 0


# ================== 状态判断工具函数 ==================
def is_currently_on_duty(records: list) -> bool:
    if not records:
        return False

    shift1_active = False
    shift2_active = False

    for r in records:
        act = r.get("action")
        if act == "1":
            shift1_active = True
        elif act == "2":
            shift1_active = False
        elif act == "3":
            shift2_active = True
        elif act == "4":
            shift2_active = False

    return shift1_active or shift2_active


def has_started_work_today(records: list) -> bool:
    """当天是否至少打过一次上班卡"""
    return any(r.get("action") in {"1", "3"} for r in records)

def get_late_minutes(expected: str, shift: str = None, now: datetime = None) -> int:
    if not expected or shift not in {"1", "3"}:
        return 0
    
    if now is None:
        now = beijing_now()
    
    try:
        exp_hm = datetime.strptime(expected, "%H:%M").time()
        expected_dt = now.replace(hour=exp_hm.hour, minute=exp_hm.minute, 
                                second=0, microsecond=0)
        return max(0, int((now - expected_dt).total_seconds() / 60))
    except:
        return 0


# ================== DataManager - 并发安全强化 ==================
class DataManager:
    def __init__(self):
        self._data: dict = {}
        self._last_mtime = 0
        self._last_save = 0
        self._dirty = False
        self._global_lock = asyncio.Lock()
        self._chat_locks: dict[str, asyncio.Lock] = {}
        self._save_task = None
        self._migrated = False

    def _get_chat_lock(self, chat_id: str):
        if chat_id not in self._chat_locks:
            self._chat_locks[chat_id] = asyncio.Lock()
        return self._chat_locks[chat_id]

    def _file_mtime(self) -> float:
        try:
            return os.path.getmtime(DATA_FILE) if os.path.exists(DATA_FILE) else 0
        except:
            return 0

    def load(self, force: bool = False) -> dict:
        current_mtime = self._file_mtime()
        if force or current_mtime > self._last_mtime or not self._data:
            if os.path.exists(DATA_FILE):
                try:
                    with open(DATA_FILE, "r", encoding="utf-8") as f:
                        self._data = json.load(f)
                    print(f"📥 数据已从磁盘加载 | 群组: {len(self._data)}")
                except Exception as e:
                    print(f"❌ 加载数据失败: {e}")
                    self._data = {}
            else:
                self._data = {}
            
            self._last_mtime = current_mtime
            self._dirty = False
            
            if not self._migrated:
                self._migrate_historical_data()
                self._migrated = True
        return self._data

    def _migrate_historical_data(self):
        print("🔄 开始执行历史数据日期迁移（05:00分界 + Shift4 跨天）...")
        migrated_count = 0
        changed = False
        
        for chat_id, chat_data in self._data.items():
            users = chat_data.get("users", {})
            for user_id, user_info in users.items():
                records = user_info.get("records", {})
                new_records: dict[str, list] = {}
                
                for old_date, rec_list in list(records.items()):
                    for rec in rec_list:
                        action = rec.get("action")
                        time_str = rec.get("time", "00:00:00")
                        
                        try:
                            rec_time = datetime.strptime(time_str, "%H:%M:%S").time()
                            dummy_dt = datetime.strptime(old_date, "%Y-%m-%d").replace(
                                hour=rec_time.hour, minute=rec_time.minute, 
                                second=rec_time.second, tzinfo=TZ
                            )
                            new_date = get_record_date(action, dummy_dt)
                            
                            if new_date not in new_records:
                                new_records[new_date] = []
                            
                            if not any(
                                r.get("time") == rec.get("time") and 
                                r.get("action") == action and
                                r.get("display") == rec.get("display")
                                for r in new_records.get(new_date, [])
                            ):
                                new_records[new_date].append(rec.copy())
                                if new_date != old_date:
                                    migrated_count += 1
                                    changed = True
                        except Exception:
                            if old_date not in new_records:
                                new_records[old_date] = []
                            new_records[old_date].append(rec.copy())
                
                user_info["records"] = new_records
        
        if migrated_count > 0 or changed:
            self._dirty = True
            print(f"✅ 历史数据迁移完成，共调整 {migrated_count} 条记录")
        else:
            print("✅ 历史数据无需迁移或已完成")

    async def aload(self, force: bool = False) -> dict:
        return await asyncio.to_thread(self.load, force)

    async def save(self, immediate: bool = False):
        async with self._global_lock:
            if not self._dirty and not immediate:
                return
            try:
                temp_file = DATA_FILE + ".tmp"
                backup_file = DATA_FILE + ".bak"

                with open(temp_file, "w", encoding="utf-8") as f:
                    json.dump(self._data, f, ensure_ascii=False, indent=2)

                if os.path.exists(DATA_FILE):
                    shutil.copy2(DATA_FILE, backup_file)

                os.replace(temp_file, DATA_FILE)

                self._last_mtime = self._file_mtime()
                self._last_save = time_module.time()
                self._dirty = False
                print(f"💾 数据已安全保存 | 群组: {len(self._data)} | {beijing_now()}")
            except Exception as e:
                print(f"❌ 保存失败: {e}")

    async def _delayed_save(self):
        await asyncio.sleep(3)
        await self.save()

    async def get_chat_data(self, chat_id: str):
        async with self._get_chat_lock(chat_id):
            await self.aload()
            return self._data.setdefault(chat_id, {
                "registered": {},
                "users": {},
                "admins": []
            })

    async def update_chat_data(self, chat_id: str, chat_data: dict):
        async with self._get_chat_lock(chat_id):
            await self.aload()
            self._data[chat_id] = chat_data
            self._dirty = True
            if not self._save_task or self._save_task.done():
                self._save_task = asyncio.create_task(self._delayed_save())

    async def force_save(self):
        await self.save(immediate=True)

    async def cleanup_old_data(self):
        async with self._global_lock:
            await self.aload(force=True)
            cutoff = (beijing_now() - timedelta(days=90)).strftime("%Y-%m-%d")
            cleaned = 0
            for chat_id in list(self._data.keys()):
                for user_id in list(self._data[chat_id].get("users", {}).keys()):
                    records = self._data[chat_id]["users"][user_id].get("records", {})
                    for d in list(records.keys()):
                        if d < cutoff:
                            del records[d]
                            cleaned += 1
            if cleaned > 0:
                self._dirty = True
                print(f"🧹 已清理 {cleaned} 条旧记录")
                await self.force_save()


data_manager = DataManager()

# ================== ACTIONS ==================
ACTIONS = {
    "1": {"name": "第一班上班", "time": "14:00", "is_work": True,  "type": "work"},
    "2": {"name": "第一班下班", "time": "19:00", "is_work": False, "type": "work"},
    "3": {"name": "第二班上班", "time": "21:00", "is_work": True,  "type": "work"},
    "4": {"name": "第二班下班", "time": "04:00", "is_work": False, "type": "work"},
    "5": {"name": "开始休息",       "time": None, "is_work": False, "type": "rest_start"},
    "6": {"name": "结束休息",       "time": None, "is_work": False, "type": "rest_end"},
    "7": {"name": "工作原因暂离座位", "time": None, "is_work": False, "type": "work_rest_start"},
    "8": {"name": "作业结束回到座位", "time": None, "is_work": False, "type": "work_rest_end"},
}

# ================== 休息超时提醒（新增功能） ==================
async def check_rest_timeout(context: ContextTypes.DEFAULT_TYPE):
    """检查用户休息是否超时未结束（仅针对打5未打6）"""
    job_data = context.job.data
    chat_id = job_data["chat_id"]
    user_id = job_data["user_id"]
    start_time = job_data["start_time"]
    
    try:
        chat_data = await data_manager.get_chat_data(str(chat_id))
        user_records = chat_data.get("users", {}).get(user_id, {}).get("records", {})
        
        today = get_attendance_date(beijing_now())
        records = user_records.get(today, [])
        
        still_resting = any(
            r.get("type") == "rest_start" and "rest_minutes" not in r 
            for r in records
        )
        
        if still_resting:
            await context.bot.send_message(
                chat_id=user_id,
                text=f"⚠️ **休息超时提醒**\n\n"
                     f"您已在 **{start_time}** 开始休息，\n"
                     f"已超过 **60分钟** 仍未结束休息（未打6）。\n\n"
                     f"请尽快回复 **6** 结束休息。",
                parse_mode="Markdown"
            )
            print(f"✅ 已发送休息超时提醒给用户 {user_id}")
        else:
            print(f"ℹ️ 用户 {user_id} 休息已结束，无需提醒")
    except Exception as e:
        print(f"❌ 休息提醒任务执行失败 (用户{user_id}): {e}")


# ================== 报表生成 ==================
def build_daily_report_rows(chat_data: dict, report_date: str):
    registered = chat_data.get("registered", {})
    users = chat_data.get("users", {})
    rows = []
    
    for user_id, user_name in registered.items():
        user_info = users.get(user_id, {"name": user_name, "records": {}})
        records = user_info.get("records", {}).get(report_date, [])

        shifts = {r.get("action"): r for r in records if r.get("action") in {"1", "2", "3", "4"}}

        # ================== 修复后的休息统计 ==================
        total_rest = 0
        rest_count = 0
        total_work_rest = 0
        work_rest_count = 0

        for r in records:
            if r.get("rest_minutes"):
                minutes = r.get("rest_minutes", 0)
                rtype = r.get("type")
                if rtype == "rest_start":          # 普通休息（5-6）
                    total_rest += minutes
                    rest_count += 1
                elif rtype == "work_rest_start":   # 工作原因暂离（7-8）
                    total_work_rest += minutes
                    work_rest_count += 1

        # ================== 迟到等其他逻辑保持不变 ==================
        late1 = shifts.get("1", {}).get("late_min", 0)
        late2 = shifts.get("3", {}).get("late_min", 0)

        missing = set('1234') - set(shifts.keys())
        status = "正常" if not missing else f"缺卡: {','.join(sorted(missing))}"

        rows.append({
            "姓名": user_name,
            "日期": report_date,
            "第一班上班": shifts.get("1", {}).get("time", "缺卡"),
            "第一班下班": shifts.get("2", {}).get("time", "缺卡"),
            "第二班上班": shifts.get("3", {}).get("time", "缺卡"),
            "第二班下班": shifts.get("4", {}).get("time", "缺卡"),
            "第一班迟到": late1,
            "第二班迟到": late2,
            "休息次数": rest_count,
            "总休息分钟": total_rest,
            "工作原因休息次数": work_rest_count,
            "工作原因总休息分钟": total_work_rest,
            "状态": status,
        })
    
    rows.sort(key=lambda x: x["姓名"])
    return rows


def build_month_report_rows(chat_data: dict, current_month: str):
    rows = []
    for user_id, user_name in chat_data.get("registered", {}).items():
        user_records = chat_data.get("users", {}).get(user_id, {}).get("records", {})
        for date_str in sorted(d for d in user_records if d.startswith(current_month)):
            daily = build_daily_report_rows(
                {"registered": {user_id: user_name}, 
                 "users": {user_id: user_records}}, 
                date_str)
            rows.extend(daily)
    
    rows.sort(key=lambda x: (x["姓名"], x["日期"]))
    return rows


def cleanup_old_excels():
    """清理3天前的Excel文件"""
    try:
        now = beijing_now()
        for f in os.listdir(EXCEL_FOLDER):
            if f.endswith(".xlsx"):
                path = os.path.join(EXCEL_FOLDER, f)
                file_mtime = datetime.fromtimestamp(os.path.getmtime(path), tz=TZ)
                if (now - file_mtime).days >= 3:
                    try:
                        os.remove(path)
                        print(f"🗑️ 已清理过期Excel: {f}")
                    except Exception as e:
                        print(f"清理文件失败 {f}: {e}")
    except Exception as e:
        print(f"清理Excel失败: {e}")


# ================== 核心打卡函数 ==================
async def daka(update: Update, context: ContextTypes.DEFAULT_TYPE, shift: str):
    chat_id_str = str(update.effective_chat.id)
    user = update.effective_user
    user_id = str(user.id)

    await auto_register(update, context)

    now = beijing_now()
    time_str = now.strftime("%H:%M:%S")
    date_str = get_record_date(shift, now)

    valid, msg = is_valid_checkin_time(shift, now)
    if not valid:
        await update.message.reply_text(msg)
        return

    if shift in ["5", "7"]:
        valid_rest, rest_msg = is_valid_rest_time(shift)
        if not valid_rest:
            await update.message.reply_text(rest_msg)
            return

    chat_data = await data_manager.get_chat_data(chat_id_str)

    user_data = chat_data["users"].setdefault(user_id, {"name": user.full_name, "records": {}})
    records = user_data["records"].setdefault(date_str, [])

    is_on_duty = is_currently_on_duty(records)
    is_resting = any(r.get("type") == "rest_start" and "rest_minutes" not in r for r in records)
    is_work_resting = any(r.get("type") == "work_rest_start" and "rest_minutes" not in r for r in records)

    if shift in ["5", "7"]:
        if not is_on_duty:
            await update.message.reply_text(
                "⚠️ **当前不在上班状态**，无法开始休息或暂离！\n\n"
                "请先打上班卡（1 或 3）后再操作。"
            )
            return
        if not has_started_work_today(records):
            await update.message.reply_text("⚠️ 必须先打上班卡（1 或 3）才能开始休息/暂离")
            return

    if shift in ["5", "7"] and (is_resting or is_work_resting):
        await update.message.reply_text("⏳ 当前正在休息中，请先结束再开始新休息")
        return
    if shift == "6" and not is_resting:
        await update.message.reply_text("⚠️ 请先输入5开始休息")
        return
    if shift == "8" and not is_work_resting:
        await update.message.reply_text("⚠️ 请先输入7工作原因暂离")
        return

    if shift in ["1","2","3","4"] and any(r.get("action") == shift for r in records):
        await update.message.reply_text(f"⚠️ {date_str} 已打过 {ACTIONS[shift]['name']}")
        return

    # ================== 记录打卡 ==================
    action = ACTIONS.get(shift, {"name": shift, "type": "unknown"})
    late = get_late_minutes(action.get("time"), shift, now)

    display = action["name"]
    rest_min = 0

    if shift in ["6", "8"]:
        target = "rest_start" if shift == "6" else "work_rest_start"
        for r in reversed(records):
            if r.get("type") == target and "rest_minutes" not in r:
                rest_min = calculate_rest_duration(r["time"], time_str)
                r["rest_minutes"] = rest_min
                display = f"{action['name']}（{rest_min}分钟）"
                break

    records.append({
        "time": time_str,
        "action": shift,
        "display": display,
        "late_min": late,
        "type": action.get("type")
    })

    # ================== 新增：5号休息60分钟未结束提醒 ==================
    if shift == "5":
        job_name = f"rest_timeout_{chat_id_str}_{user_id}_{time_str}"
        context.job_queue.run_once(
            callback=check_rest_timeout,
            when=3600,                    # 60分钟
            data={
                "chat_id": int(chat_id_str),
                "user_id": user_id,
                "start_time": time_str
            },
            name=job_name
        )
        print(f"⏰ 已为用户 {user_id} 启动休息60分钟提醒任务")

    await data_manager.update_chat_data(chat_id_str, chat_data)

    emoji = "⚠️" if late > 0 else "✅"
    late_txt = f"（迟到{late}分钟）" if late > 0 else ""
    await update.message.reply_text(
        f"{emoji} **{user.full_name}** {display}{late_txt}\n日期：{date_str}\n时间：{time_str}",
        parse_mode="Markdown"
    )


# ================== 消息处理 ==================
async def text_daka(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip().lower()
    mapping = {
        "1":"1","上班":"1","上午":"1",
        "2":"2","下班":"2","下班1":"2","下1":"2",
        "3":"3","下午上班":"3","上班2":"3",
        "4":"4","下班2":"4","下2":"4",
        "5":"5","休息":"5","开始休息":"5",
        "6":"6","结束休息":"6","回岗":"6",
        "7":"7","暂离":"7","离开":"7","工作原因休息":"7",
        "8":"8","回到座位":"8","回座位":"8",
    }
    if text in mapping:
        await daka(update, context, mapping[text])


async def auto_register(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id_str = str(update.effective_chat.id)
    user_id = str(update.effective_user.id)
    name = update.effective_user.full_name

    chat_data = await data_manager.get_chat_data(chat_id_str)
    if user_id not in chat_data["registered"]:
        chat_data["registered"][user_id] = name
        chat_data["users"].setdefault(user_id, {"name": name, "records": {}})
        await data_manager.update_chat_data(chat_id_str, chat_data)
        await update.message.reply_text(f"✅ **{name}** 自动注册成功！", parse_mode="Markdown")


# ================== 管理员命令 ==================
async def get_group_owner(context: ContextTypes.DEFAULT_TYPE, chat_id: int):
    try:
        admins = await context.bot.get_chat_administrators(chat_id)
        for a in admins:
            if a.status == "creator":
                return a.user.id
    except:
        pass
    return None

async def is_admin(update: Update) -> bool:
    if update.effective_chat.type == "private":
        return True
    try:
        member = await update.effective_chat.get_member(update.effective_user.id)
        return member.status in ["administrator", "creator"]
    except:
        return False

async def add_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    chat_id_str = str(update.effective_chat.id)
    owner = await get_group_owner(context, int(chat_id_str))
    if owner and owner != update.effective_user.id:
        await update.message.reply_text("⚠️ 仅群主可操作管理员")
        return
    if not context.args:
        await update.message.reply_text("用法: /addadmin <用户ID>")
        return
    target = context.args[0].strip()
    chat_data = await data_manager.get_chat_data(chat_id_str)
    if target not in chat_data["admins"]:
        chat_data["admins"].append(target)
        await data_manager.update_chat_data(chat_id_str, chat_data)
        await data_manager.force_save()
        await update.message.reply_text(f"✅ 已添加管理员 `{target}`", parse_mode="Markdown")
    else:
        await update.message.reply_text("✅ 该用户已是管理员")

async def del_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    chat_id_str = str(update.effective_chat.id)
    owner = await get_group_owner(context, int(chat_id_str))
    if owner and owner != update.effective_user.id:
        await update.message.reply_text("⚠️ 仅群主可操作管理员")
        return
    if not context.args:
        await update.message.reply_text("用法: /deladmin <用户ID>")
        return
    target = context.args[0].strip()
    chat_data = await data_manager.get_chat_data(chat_id_str)
    if target in chat_data["admins"]:
        chat_data["admins"].remove(target)
        await data_manager.update_chat_data(chat_id_str, chat_data)
        await data_manager.force_save()
        await update.message.reply_text(f"✅ 已删除管理员 `{target}`", parse_mode="Markdown")
    else:
        await update.message.reply_text("❌ 该用户不是管理员")

async def adminlist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    chat_id_str = str(update.effective_chat.id)
    chat_data = await data_manager.get_chat_data(chat_id_str)
    owner = await get_group_owner(context, int(chat_id_str))
    text = "📋 **日报表接收人**\n\n"
    if owner:
        text += f"👑 群主: `{owner}`\n"
    text += f"📌 指定管理员 ({len(chat_data.get('admins',[]))}人):\n"
    for i, aid in enumerate(chat_data.get("admins", []), 1):
        text += f"{i}. `{aid}`\n"
    await update.message.reply_text(text, parse_mode="Markdown")

async def deluser(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    if not context.args:
        await update.message.reply_text("用法: /deluser <用户ID 或 @用户名>")
        return
    chat_id_str = str(update.effective_chat.id)
    target = context.args[0].strip()
    chat_data = await data_manager.get_chat_data(chat_id_str)
    target_id = None

    if target.startswith('@'):
        name_search = target[1:].lower()
        for uid, name in chat_data["registered"].items():
            if name.lower() == name_search:
                target_id = uid
                break
    elif target in chat_data["registered"]:
        target_id = target

    if target_id and target_id in chat_data["registered"]:
        name = chat_data["registered"].pop(target_id)
        chat_data["users"].pop(target_id, None)
        await data_manager.update_chat_data(chat_id_str, chat_data)
        await data_manager.force_save()
        await update.message.reply_text(f"✅ 已删除用户：**{name}**", parse_mode="Markdown")
    else:
        await update.message.reply_text("❌ 未找到该用户")

async def delete_record(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    
    if not context.args:
        await update.message.reply_text("用法: `/del YYYY-MM-DD`", parse_mode="Markdown")
        return
    
    date_to_del = context.args[0].strip()
    if len(date_to_del) != 10 or date_to_del[4] != '-' or date_to_del[7] != '-':
        await update.message.reply_text("❌ 日期格式错误！请使用 `YYYY-MM-DD` 格式")
        return

    chat_id_str = str(update.effective_chat.id)
    chat_data = await data_manager.get_chat_data(chat_id_str)
    
    count = 0
    affected_users = []
    
    for user_id, user_info in list(chat_data.get("users", {}).items()):
        records_dict = user_info.get("records", {})
        if date_to_del in records_dict:
            del records_dict[date_to_del]
            count += 1
            user_name = chat_data.get("registered", {}).get(user_id, user_id)
            affected_users.append(user_name)
    
    if count == 0:
        await update.message.reply_text(f"ℹ️ 日期 **{date_to_del}** 没有找到任何打卡记录", parse_mode="Markdown")
        return
    
    await data_manager.update_chat_data(chat_id_str, chat_data)
    await data_manager.force_save()
    
    user_list = ", ".join(affected_users[:6])
    if len(affected_users) > 6:
        user_list += f" 等共 {len(affected_users)} 人"
    
    await update.message.reply_text(
        f"✅ **删除成功**\n\n📅 日期：**{date_to_del}**\n👥 影响用户：**{count}** 人\n用户：{user_list}",
        parse_mode="Markdown"
    )

# ================== 报表命令 ==================
async def todayexcel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    chat_id_str = str(update.effective_chat.id)
    today = get_attendance_date(beijing_now())
    chat_data = await data_manager.get_chat_data(chat_id_str)
    rows = build_daily_report_rows(chat_data, today)

    filename = f"全群打卡_{today}.xlsx"
    filepath = os.path.join(EXCEL_FOLDER, filename)

    cols = ["姓名","日期","第一班上班","第一班下班","第二班上班","第二班下班",
            "第一班迟到","第二班迟到","休息次数","总休息分钟",
            "工作原因休息次数","工作原因总休息分钟","状态"]
    df = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
    df.to_excel(filepath, index=False)

    with open(filepath, 'rb') as f:
        await update.message.reply_document(f, filename=filename, caption=f"✅ {today} 全群打卡报表")


async def monthexcel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    chat_id_str = str(update.effective_chat.id)
    month = beijing_now().strftime("%Y-%m")
    chat_data = await data_manager.get_chat_data(chat_id_str)
    rows = build_month_report_rows(chat_data, month)

    filename = f"全群打卡_{month}.xlsx"
    filepath = os.path.join(EXCEL_FOLDER, filename)

    cols = ["姓名","日期","第一班上班","第一班下班","第二班上班","第二班下班",
            "第一班迟到","第二班迟到","休息次数","总休息分钟",
            "工作原因休息次数","工作原因总休息分钟","状态"]
    df = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
    df.to_excel(filepath, index=False)

    with open(filepath, 'rb') as f:
        await update.message.reply_document(f, filename=filename, 
                                          caption=f"✅ {month} 月报表\n共 {len(rows)} 条记录")


async def absent(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_admin(update):
        await update.message.reply_text("⚠️ 仅限管理员使用")
        return
    chat_id_str = str(update.effective_chat.id)
    today = get_attendance_date(beijing_now())
    chat_data = await data_manager.get_chat_data(chat_id_str)
    registered = chat_data.get("registered", {})
    users = chat_data.get("users", {})

    incomplete = []
    for uid, name in registered.items():
        records = users.get(uid, {}).get("records", {}).get(today, [])
        done = {r["action"] for r in records if r.get("action") in "1234"}
        if done != {"1","2","3","4"}:
            incomplete.append(f"{name} → 已打: {','.join(sorted(done)) if done else '无'}")

    if not incomplete:
        await update.message.reply_text("🎉 今天所有人均已完成全部打卡！")
    else:
        text = f"📋 **今日未完成打卡人员** ({len(incomplete)}/{len(registered)})\n\n"
        text += "\n".join(f"{i+1}. {item}" for i, item in enumerate(incomplete))
        await update.message.reply_text(text, parse_mode="Markdown")


# ================== 自动日报 ==================
async def send_daily_report(context: ContextTypes.DEFAULT_TYPE):
    cleanup_old_excels()
    report_date = get_report_date_for_daily()
    print(f"📊 开始生成自动日报 | 日期: {report_date}")

    all_data = await data_manager.aload(force=True)
    
    success_count = 0
    fail_count = 0
    total_groups = len(all_data)

    for chat_id_str, chat_data in all_data.items():
        chat_id = int(chat_id_str)
        recipients = set()
        owner = await get_group_owner(context, chat_id)
        if owner:
            recipients.add(owner)
        recipients.update(int(uid) for uid in chat_data.get("admins", []))

        if not recipients:
            continue

        try:
            fresh_chat_data = await data_manager.get_chat_data(chat_id_str)
            rows = build_daily_report_rows(fresh_chat_data, report_date)

            filename = f"全群打卡日报_{report_date}_{chat_id_str}.xlsx"
            filepath = os.path.join(EXCEL_FOLDER, filename)

            cols = ["姓名","日期","第一班上班","第一班下班","第二班上班","第二班下班",
                    "第一班迟到","第二班迟到","休息次数","总休息分钟",
                    "工作原因休息次数","工作原因总休息分钟","状态"]
            
            df = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
            df.to_excel(filepath, index=False)

            caption = f"📊 **{report_date} 全群日报**（05:00~次日05:00）\n👥 总注册: {len(rows)} 人"

            sent_to = 0
            for rid in recipients:
                for attempt in range(3):
                    try:
                        with open(filepath, 'rb') as f:
                            await context.bot.send_document(
                                rid, f, 
                                filename=filename.replace(f"_{chat_id_str}", ""),
                                caption=caption, 
                                parse_mode="Markdown"
                            )
                        sent_to += 1
                        break
                    except:
                        if attempt < 2:
                            await asyncio.sleep(2 ** attempt)
            if sent_to > 0:
                success_count += 1
        except Exception as e:
            fail_count += 1
            print(f"处理群 {chat_id} 日报异常: {e}")

    print(f"📨 自动日报任务完成 | 成功: {success_count}/{total_groups}")


# ================== 其他命令 ==================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "飞机的代号确定下来了就不要再改了，否则会打卡记录失败\n\n"
        "第一班上班打1，下班打2。第二班上班打3，下班打4，离开工位休息打5，回来打6。\n\n"
        "上下班打卡的，迟到早退相同，10分钟内扣50，1小时内扣100，1小时外按旷工扣200.超过1秒也算迟到跟早退。漏打每次100（下班打卡有效时间1小时）.\n"
        "严禁互相打卡与飞机定时发送。互相打卡两个人各扣300，定时发送扣600.\n"
        "如果遇到没有信号的缘故，或者帮公司做其他事情没办法及时打卡的，找公司组长或副组长证明补打卡后原因写上。其余不管是加班聊客户或者其他原因的也算迟到早退。\n\n"
        "私聊机器人发送 /myrecord 可查询个人记录\n"
    )

async def registered_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id_str = str(update.effective_chat.id)
    chat_data = await data_manager.get_chat_data(chat_id_str)
    registered = chat_data.get("registered", {})
    if not registered:
        await update.message.reply_text("📋 本群暂无注册人员。")
        return
    text = f"📋 **本群已注册人员**（{len(registered)}人）\n\n"
    for i, (uid, name) in enumerate(registered.items(), 1):
        text += f"{i}. {name} (`{uid}`)\n"
    await update.message.reply_text(text, parse_mode="Markdown")

async def myrecord(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.type != "private":
        await update.message.reply_text("此命令仅支持私聊使用")
        return
    user_id = str(update.effective_user.id)
    data = await data_manager.aload()
    text = f"📋 **{update.effective_user.full_name}** 打卡记录\n\n"
    found = False
    for chat_id, cdata in data.items():
        urec = cdata.get("users", {}).get(user_id, {}).get("records", {})
        if not urec: continue
        found = True
        text += f"**群 {chat_id}**\n"
        for date in sorted(urec.keys(), reverse=True)[:15]:
            recs = urec[date]
            if not recs: continue
            text += f"**{date}**\n"
            for r in recs:
                late = f"（迟到{r.get('late_min',0)}分）" if r.get("late_min") else ""
                text += f"• {r.get('display', r.get('action'))}{late} {r['time']}\n"
            text += "\n"
    await update.message.reply_text(text if found else "暂无记录", parse_mode="Markdown")

async def today_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    now = beijing_now()
    att_date = get_attendance_date(now)
    report_date = get_report_date_for_daily()
    await update.message.reply_text(
        f"🕒 当前北京时间：**{now.strftime('%Y-%m-%d %H:%M:%S')}**\n"
        f"📅 当前考勤日期：**{att_date}**\n"
        f"📊 今日05:30将发送的日报日期：**{report_date}**",
        parse_mode="Markdown"
    )

# ================== 主程序 ==================
def main():
    data_manager.load(force=True)
    
    from telegram.ext import Defaults
    
    app = Application.builder() \
        .token(TOKEN) \
        .defaults(Defaults(tzinfo=TZ)) \
        .build()
    
    jq: JobQueue = app.job_queue

    jq.run_daily(send_daily_report, time(5, 30, 0))
    jq.run_daily(data_manager.cleanup_old_data, time(6, 10, 0))

    handlers = [
        CommandHandler("start", start),
        CommandHandler("register", lambda u,c: auto_register(u,c)),
        CommandHandler("registered", registered_list),
        CommandHandler("myrecord", myrecord),
        CommandHandler("addadmin", add_admin),
        CommandHandler("deladmin", del_admin),
        CommandHandler("adminlist", adminlist),
        CommandHandler("deluser", deluser),
        CommandHandler("del", delete_record),
        CommandHandler("todayexcel", todayexcel),
        CommandHandler("monthexcel", monthexcel),
        CommandHandler("absent", absent),
        CommandHandler("today", today_cmd),
    ]
    for h in handlers:
        app.add_handler(h)

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_daka))

    print("🚀 打卡机器人已完全启动（啊原的第5个版本 - 5.9）")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
