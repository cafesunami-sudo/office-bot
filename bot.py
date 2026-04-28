from aiogram import Bot, Dispatcher
from aiogram.types import Message, ReplyKeyboardMarkup, KeyboardButton, FSInputFile
from aiogram.filters import Command
from docx import Document
from datetime import datetime, timedelta
import asyncio
import os
import json
import re

TOKEN = "8690185918:AAGBg90wZ7Wyea8-JBI4aSPKPUS-OUIhuVk"

ADMIN_CHAT_ID = 137602775
GROUP_CHAT_ID = -1003555153586

ADMIN_USERS = [137602775]
REPORT_ONLY_USERS = [390687401]

REMIND_TIMES = ["09:30"]
START_REMIND_TIME = "17:00"

EMPLOYEES_FILE = "employees/sotrudniki.docx"
READY_FOLDER = "ready"
HISTORY_FILE = "history.json"
HOLIDAYS_FILE = "holidays.json"
REMINDERS_SENT_FILE = "reminders_sent.json"

TEMPLATES = {
    "🌴 Полный отпуск": "templates/otpusk_full.docx",
    "🧩 Часть отпуска": "templates/otpusk_part.docx",
    "📌 Оставшийся отпуск": "templates/otpusk_rest.docx",
    "📚 Учебный отпуск": "templates/study_leave.docx",
    "📝 БС с периода по период": "templates/bs_range.docx",
    "📅 БС на один день": "templates/bs_one.docx",
    "💍 Мат помощь (свадьба)": "templates/mat_wedding.docx",
    "👶 Мат помощь (ребенок)": "templates/mat_child.docx"
}

TRIP_REPORT_TEMPLATE = "templates/trip/business_trip_report.docx"
TRIP_CERTIFICATE_TEMPLATE = "templates/trip/business_trip_certificate.docx"
TRIP_TYPE = "✈️ Командировка"
TRIP_REGION_SUFFIX = "подразделения по Хорезмской области"

SICK_LEAVE_TYPE = "🏥 Больничный"
BS_LEAVE_TYPES = ["📝 БС с периода по период", "📅 БС на один день"]
BS_RANGE_TYPE = "📝 БС с периода по период"

MANUAL_TYPES = [
    "🌴 Полный отпуск",
    "🧩 Часть отпуска",
    "📌 Оставшийся отпуск",
    "📚 Учебный отпуск",
    "📝 БС с периода по период",
    "📅 БС на один день",
    "🏥 Больничный"
]

bot = Bot(token=TOKEN)
dp = Dispatcher()

data = {}
state = {}
temp_search = {}


def make_keyboard(buttons, cols=2):
    keyboard = []
    row = []

    for btn in buttons:
        row.append(KeyboardButton(text=btn))
        if len(row) == cols:
            keyboard.append(row)
            row = []

    if row:
        keyboard.append(row)

    return ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)


def is_allowed(chat_id):
    return chat_id in ADMIN_USERS or chat_id in REPORT_ONLY_USERS


def is_admin(chat_id):
    return chat_id in ADMIN_USERS


def is_report_only(chat_id):
    return chat_id in REPORT_ONLY_USERS


def get_menu(chat_id):
    if is_admin(chat_id):
        return menu
    if is_report_only(chat_id):
        return report_only_menu
    return None


def normalize_date(text):
    """Приводит дату к единому виду ДД.ММ.ГГГГ.
    Например: 3.05.2026 -> 03.05.2026, 03.5.2026 -> 03.05.2026.
    """
    text = str(text).strip()
    match = re.fullmatch(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", text)
    if not match:
        return text

    day, month, year = match.groups()
    try:
        dt = datetime(int(year), int(month), int(day))
        return dt.strftime("%d.%m.%Y")
    except:
        return text


def is_valid_date(text):
    text = normalize_date(text)
    if not re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", text):
        return False
    try:
        datetime.strptime(normalize_date(text), "%d.%m.%Y")
        return True
    except:
        return False

def load_holidays():
    if not os.path.exists(HOLIDAYS_FILE):
        return set()
    try:
        with open(HOLIDAYS_FILE, "r", encoding="utf-8") as f:
            return set(json.load(f))
    except:
        return set()


HOLIDAYS = load_holidays()


def get_return_to_work_date(end_date_text):
    d = datetime.strptime(end_date_text, "%d.%m.%Y") + timedelta(days=1)
    while d.weekday() >= 5 or d.strftime("%Y-%m-%d") in HOLIDAYS:
        d += timedelta(days=1)
    return d.strftime("%d.%m.%Y")


def load_sent_reminders():
    if not os.path.exists(REMINDERS_SENT_FILE):
        return []
    try:
        with open(REMINDERS_SENT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return []


def save_sent_reminder(key):
    sent = load_sent_reminders()
    if key not in sent:
        sent.append(key)
    with open(REMINDERS_SENT_FILE, "w", encoding="utf-8") as f:
        json.dump(sent, f, ensure_ascii=False, indent=2)


def month_name(date_text):
    months = {
        "01": "января", "02": "февраля", "03": "марта", "04": "апреля",
        "05": "мая", "06": "июня", "07": "июля", "08": "августа",
        "09": "сентября", "10": "октября", "11": "ноября", "12": "декабря"
    }
    dt = datetime.strptime(date_text, "%d.%m.%Y")
    return months[dt.strftime("%m")]


def format_date_text(date_text):
    dt = datetime.strptime(date_text, "%d.%m.%Y")
    return f"«{dt.strftime('%d')}» {month_name(date_text)} {dt.year} года"


def format_date_text_without_year_word(date_text):
    dt = datetime.strptime(date_text, "%d.%m.%Y")
    return f"«{dt.strftime('%d')}» {month_name(date_text)} {dt.year}"


def format_date_range_start(date_text):
    dt = datetime.strptime(date_text, "%d.%m.%Y")
    return f"«{dt.strftime('%d')}» {month_name(date_text)} {dt.year}"


def format_date_range_end(date_text):
    dt = datetime.strptime(date_text, "%d.%m.%Y")
    return f"«{dt.strftime('%d')}» {month_name(date_text)} {dt.year} года"


def format_fio_short(fio):
    parts = fio.split()
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    return fio


def type_uz(t):
    if t == "📌 Оставшийся отпуск":
        return "Mehnat ta’tilining qolgan qismi"
    if t in ["🌴 Полный отпуск", "🧩 Часть отпуска"]:
        return "Mehnat ta’tili"
    if t in ["📝 БС с периода по период", "📅 БС на один день"]:
        return "Ish haqi saqlanmagan ta’til"
    if t == "📚 Учебный отпуск":
        return "O‘quv ta’tili"
    if t == "🏥 Больничный":
        return "Kasallik ta’tili"
    if t in ["💍 Мат помощь (свадьба)", "👶 Мат помощь (ребенок)"]:
        return "Moddiy yordam"
    return t


def start_phrase_uz(t):
    if t == "📌 Оставшийся отпуск":
        return "mehnat ta’tilining qolgan qismiga chiqadi"
    if t in ["🌴 Полный отпуск", "🧩 Часть отпуска"]:
        return "ta’tilga chiqadi"
    if t in ["📝 БС с периода по период", "📅 БС на один день"]:
        return "ish haqi saqlanmagan ta’tilga chiqadi"
    if t == "📚 Учебный отпуск":
        return "o‘quv ta’tiliga chiqadi"
    if t == "🏥 Больничный":
        return "kasallik ta’tiliga chiqadi"
    return "ta’tilga chiqadi"


def load_employees():
    doc = Document(EMPLOYEES_FILE)
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]


def search_employees_by_text(text, employees):
    q = str(text).lower().strip()

    if len(q) < 2:
        return None

    return [
        e for e in employees
        if any(part.lower().startswith(q) for part in e.split())
    ]


def fio_startswith_text(fio, text):
    q = str(text).lower().strip()

    if len(q) < 2:
        return False

    return any(part.lower().startswith(q) for part in str(fio).split())


def load_history():
    if not os.path.exists(HISTORY_FILE):
        return []
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return []


def save_history_full(history):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


def normalize_created_at(value):
    if not value:
        return datetime.now().strftime("%d.%m.%Y %H:%M:%S")

    value = str(value).strip()
    formats = ["%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%d.%m.%Y"]

    for fmt in formats:
        try:
            return datetime.strptime(value, fmt).strftime("%d.%m.%Y %H:%M:%S")
        except:
            pass

    return datetime.now().strftime("%d.%m.%Y %H:%M:%S")


def normalize_history_record(record):
    fixed = dict(record)

    fixed["fio"] = fixed.get("fio", "")
    fixed["position"] = fixed.get("position", "")
    fixed["project"] = fixed.get("project", "")
    fixed["type"] = fixed.get("type", "")
    fixed["start"] = normalize_date(fixed.get("start", "")) if fixed.get("start") else ""
    fixed["end"] = normalize_date(fixed.get("end", "")) if fixed.get("end") else ""

    periods = fixed.get("periods")
    if isinstance(periods, list) and periods:
        clean_periods = []
        for p in periods:
            if isinstance(p, dict) and p.get("start") and p.get("end"):
                clean_periods.append({"start": normalize_date(p.get("start", "")), "end": normalize_date(p.get("end", ""))})

        if clean_periods:
            fixed["periods"] = clean_periods
            fixed["start"] = clean_periods[0]["start"]
            fixed["end"] = clean_periods[-1]["end"]
            fixed["days"] = str(calc_total_days(clean_periods))
        elif "periods" in fixed:
            del fixed["periods"]

    if not fixed.get("days") and fixed.get("start") and fixed.get("end"):
        try:
            fixed["days"] = str(days_between(fixed["start"], fixed["end"]))
        except:
            fixed["days"] = ""
    else:
        fixed["days"] = str(fixed.get("days", ""))

    if fixed.get("end"):
        try:
            fixed["return_date"] = get_return_to_work_date(fixed["end"])
        except:
            fixed["return_date"] = fixed.get("return_date", "")
    else:
        fixed["return_date"] = fixed.get("return_date", "")

    fixed["created_at"] = normalize_created_at(fixed.get("created_at"))

    if fixed.get("extended_at"):
        fixed["extended_at"] = normalize_created_at(fixed.get("extended_at"))

    return fixed


def history_unique_key(record):
    periods = record.get("periods")
    periods_text = ""
    if isinstance(periods, list):
        periods_text = json.dumps(periods, ensure_ascii=False, sort_keys=True)

    return (
        record.get("fio", ""),
        record.get("type", ""),
        record.get("start", ""),
        record.get("end", ""),
        record.get("days", ""),
        record.get("position", ""),
        record.get("project", ""),
        periods_text,
    )


def normalize_history_file():
    history = load_history()
    if not isinstance(history, list):
        history = []

    fixed_history = []
    seen = set()

    for record in history:
        if not isinstance(record, dict):
            continue

        fixed = normalize_history_record(record)
        key = history_unique_key(fixed)

        if key in seen:
            continue

        seen.add(key)
        fixed_history.append(fixed)

    save_history_full(fixed_history)
    return fixed_history


def days_between(start_text, end_text):
    start_date = datetime.strptime(start_text, "%d.%m.%Y")
    end_date = datetime.strptime(end_text, "%d.%m.%Y")
    return (end_date - start_date).days + 1


def calc_total_days(periods):
    total = 0
    for p in periods:
        try:
            total += days_between(p.get("start", ""), p.get("end", ""))
        except:
            pass
    return total


def get_periods_from_record(record):
    periods = record.get("periods")
    if isinstance(periods, list) and periods:
        return periods

    if record.get("start") and record.get("end"):
        return [{
            "start": record.get("start"),
            "end": record.get("end")
        }]

    return []


def format_periods(periods):
    lines = []
    for i, p in enumerate(periods, 1):
        lines.append(f"{i}. {p.get('start')} — {p.get('end')}")
    return "\n".join(lines)


def is_same_record(a, b):
    return (
        a.get("fio") == b.get("fio") and
        a.get("type") == b.get("type") and
        a.get("start") == b.get("start") and
        a.get("end") == b.get("end") and
        a.get("created_at") == b.get("created_at")
    )


def find_active_sick_records_by_fio(fio):
    today = datetime.now().date()
    history = load_history()
    return [
        r for r in history
        if r.get("fio") == fio
        and r.get("type") == SICK_LEAVE_TYPE
        and is_active_record(r, today)
    ]


def find_active_sick_records_by_text(text):
    today = datetime.now().date()
    history = load_history()
    return [
        r for r in history
        if r.get("type") == SICK_LEAVE_TYPE
        and fio_startswith_text(r.get("fio", ""), text)
        and is_active_record(r, today)
    ]


def find_active_bs_records_by_fio(fio):
    today = datetime.now().date()
    history = load_history()
    return [
        r for r in history
        if r.get("fio") == fio
        and r.get("type") in BS_LEAVE_TYPES
        and is_active_record(r, today)
    ]


def find_active_bs_records_by_text(text):
    today = datetime.now().date()
    history = load_history()
    return [
        r for r in history
        if r.get("type") in BS_LEAVE_TYPES
        and fio_startswith_text(r.get("fio", ""), text)
        and is_active_record(r, today)
    ]




def find_previous_part_leave_records_by_fio(fio):
    history = load_history()
    return [
        r for r in history
        if r.get("fio") == fio
        and r.get("type") in ["🧩 Часть отпуска", "📌 Оставшийся отпуск"]
    ]


def format_leave_records_for_message(records):
    msg = ""
    for i, r in enumerate(records, 1):
        msg += f"{i}. {r.get('type')} — {r.get('start')} по {r.get('end')}"
        if r.get("days"):
            msg += f" ({r.get('days')} kun)"
        msg += "\n"
    return msg

def update_history_record(old_record, new_record):
    history = load_history()
    new_history = []
    replaced = False

    for r in history:
        if not replaced and is_same_record(r, old_record):
            new_history.append(new_record)
            replaced = True
        else:
            new_history.append(r)

    if not replaced:
        new_history.append(new_record)

    save_history_full(new_history)


def build_sick_extended_record(old_record, new_start, new_end):
    old_periods = get_periods_from_record(old_record)
    periods = old_periods + [{
        "start": new_start,
        "end": new_end
    }]

    total_days = calc_total_days(periods)
    return_date = get_return_to_work_date(new_end)

    new_record = dict(old_record)
    new_record["start"] = periods[0].get("start", "")
    new_record["end"] = new_end
    new_record["days"] = str(total_days)
    new_record["return_date"] = return_date
    new_record["periods"] = periods
    new_record["extended_at"] = datetime.now().strftime("%d.%m.%Y %H:%M:%S")

    return new_record


def build_bs_extended_record(old_record, new_start, new_end):
    old_periods = get_periods_from_record(old_record)
    periods = old_periods + [{
        "start": new_start,
        "end": new_end
    }]

    total_days = calc_total_days(periods)
    return_date = get_return_to_work_date(new_end)

    new_record = dict(old_record)
    new_record["type"] = BS_RANGE_TYPE
    new_record["start"] = periods[0].get("start", "")
    new_record["end"] = new_end
    new_record["days"] = str(total_days)
    new_record["return_date"] = return_date
    new_record["periods"] = periods
    new_record["extended_at"] = datetime.now().strftime("%d.%m.%Y %H:%M:%S")

    return new_record


def save_history(d):
    return_date = ""
    if d.get("end"):
        return_date = get_return_to_work_date(d["end"])

    record = {
        "fio": d.get("fio"),
        "position": d.get("pos", ""),
        "project": d.get("project", ""),
        "type": d.get("type"),
        "start": normalize_date(d.get("start", "")) if d.get("start") else "",
        "end": normalize_date(d.get("end", "")) if d.get("end") else "",
        "days": d.get("days", ""),
        "return_date": return_date,
        "created_at": datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    }

    if d.get("periods"):
        record["periods"] = d.get("periods")

    history = load_history()
    history.append(record)
    save_history_full(history)


def replace_text(doc, rep):
    for p in doc.paragraphs:
        text = p.text
        for k, v in rep.items():
            text = text.replace(k, v)
        if p.runs:
            p.runs[0].text = text
            for r in p.runs[1:]:
                r.text = ""

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    text = p.text
                    for k, v in rep.items():
                        text = text.replace(k, v)
                    if p.runs:
                        p.runs[0].text = text
                        for r in p.runs[1:]:
                            r.text = ""


def create_doc(d):
    doc = Document(TEMPLATES[d["type"]])

    rep = {
        "{{FIO}}": format_fio_short(d["fio"]),
        "{{POSITION}}": d.get("pos", ""),
        "{{PROJECT}}": d.get("project", ""),
        "{{TODAY}}": format_date_text_without_year_word(datetime.now().strftime("%d.%m.%Y"))
    }

    if "start" in d and "end" in d:
        rep["{{DATE_START}}"] = format_date_range_start(d["start"])
        rep["{{DATE_END}}"] = format_date_range_end(d["end"])
    elif "start" in d:
        rep["{{DATE_START}}"] = format_date_text(d["start"])

    replace_text(doc, rep)

    os.makedirs(READY_FOLDER, exist_ok=True)
    path = os.path.join(READY_FOLDER, f"{d['fio']}.docx")
    doc.save(path)
    return path




def month_name_capitalized(date_text):
    return month_name(date_text).capitalize()


def trip_position_text(position):
    pos = str(position).strip()

    if pos == "Инженер программист":
        pos = "Инженер-программист"

    if TRIP_REGION_SUFFIX.lower() not in pos.lower():
        pos = f"{pos} {TRIP_REGION_SUFFIX}"

    return pos


def safe_filename(text):
    text = re.sub(r"[^А-Яа-яA-Za-z0-9_ -]", "", str(text))
    text = text.strip().replace(" ", "_")
    return text or "file"


def create_trip_docs(d):
    os.makedirs(READY_FOLDER, exist_ok=True)

    start = d["start"]
    end = d["end"]
    now_text = datetime.now().strftime("%d.%m.%Y")

    start_dt = datetime.strptime(start, "%d.%m.%Y")
    end_dt = datetime.strptime(end, "%d.%m.%Y")
    now_dt = datetime.strptime(now_text, "%d.%m.%Y")

    employees_lines = []
    for emp in d.get("employees", []):
        employees_lines.append(f"{emp.get('fio')} - {trip_position_text(emp.get('pos', ''))}")

    common_rep = {
        "{DAY_NOW}": now_dt.strftime("%d"),
        "{MONTH_NOW}": month_name_capitalized(now_text),
        "{YEAR_NOW}": str(now_dt.year),
        "{DAY_FROM}": start_dt.strftime("%d"),
        "{DAY_TO}": end_dt.strftime("%d"),
        "{MONTH}": month_name_capitalized(start),
        "{YEAR}": str(start_dt.year),
        "{EMPLOYEES_LIST}": "\n".join(employees_lines),
    }

    report_doc = Document(TRIP_REPORT_TEMPLATE)
    replace_text(report_doc, common_rep)
    report_path = os.path.join(READY_FOLDER, f"Командировка_рапорт_{start_dt.strftime('%d_%m_%Y')}_{end_dt.strftime('%d_%m_%Y')}.docx")
    report_doc.save(report_path)

    paths = [report_path]

    for emp in d.get("employees", []):
        cert_doc = Document(TRIP_CERTIFICATE_TEMPLATE)
        rep = {
            "{FIO}": emp.get("fio", ""),
            "{POSITION}": trip_position_text(emp.get("pos", "")),
            "{DAY_FROM}": start_dt.strftime("%d"),
            "{DAY_TO}": end_dt.strftime("%d"),
            "{MONTH}": month_name_capitalized(start),
            "{YEAR}": str(start_dt.year),
        }
        replace_text(cert_doc, rep)
        cert_path = os.path.join(READY_FOLDER, f"Командировочное_удостоверение_{safe_filename(emp.get('fio', ''))}.docx")
        cert_doc.save(cert_path)
        paths.append(cert_path)

    return paths


def finish_and_send(chat_id):
    path = create_doc(data[chat_id])
    save_history(data[chat_id])
    return path


async def send_group_message(text):
    try:
        await bot.send_message(chat_id=GROUP_CHAT_ID, text=text)
    except Exception as e:
        print("GROUP SEND ERROR:", e)


async def notify_application_created(d):
    return_date = ""
    if d.get("end"):
        return_date = get_return_to_work_date(d["end"])

    msg = (
        f"📄 ARIZA TAYYORLANDI\n\n"
        f"👤 Xodim: {d.get('fio')}\n"
        f"💼 Lavozim: {d.get('pos', '')}\n"
        f"📌 Loyiha: {d.get('project', '')}\n\n"
        f"📝 Ariza turi: {type_uz(d.get('type'))}\n"
    )

    if d.get("start") and d.get("end"):
        msg += f"📅 Muddat: {d.get('start')} — {d.get('end')}\n"
    elif d.get("start"):
        msg += f"📅 Sana: {d.get('start')}\n"

    if d.get("days"):
        msg += f"⏳ Davomiyligi: {d.get('days')} kun\n"

    if d.get("type") == "📌 Оставшийся отпуск":
        msg += "\nℹ️ Xodim mehnat ta’tilining qolgan qismi bo‘yicha ariza yozdi.\n"

    if return_date:
        msg += f"📌 Ishga chiqish sanasi: {return_date}\n"

    msg += "\n✅ Ariza bot orqali shakllantirildi."

    await send_group_message(msg)


async def notify_sick_extended(old_record, new_record, added_start, added_end):
    periods = get_periods_from_record(new_record)
    return_date = new_record.get("return_date", "")

    msg = (
        f"🏥 KASALLIK TA’TILI UZAYTIRILDI\n\n"
        f"👤 Xodim: {new_record.get('fio')}\n"
        f"💼 Lavozim: {new_record.get('position', '')}\n"
        f"📌 Loyiha: {new_record.get('project', '')}\n\n"
        f"📅 Avvalgi muddat: {old_record.get('start')} — {old_record.get('end')}\n"
        f"➕ Qo‘shilgan muddat: {added_start} — {added_end}\n\n"
        f"📋 Umumiy davrlar:\n{format_periods(periods)}\n\n"
        f"⏳ Jami davomiyligi: {new_record.get('days')} kun\n"
        f"📌 Ishga chiqish sanasi: {return_date}\n\n"
        f"ℹ️ Agar oxirgi kun dam olish yoki bayram kuniga to‘g‘ri kelsa, ishga chiqish sanasi keyingi ish kuniga o‘tkazildi."
    )

    await send_group_message(msg)


async def notify_bs_extended(old_record, new_record, added_start, added_end):
    periods = get_periods_from_record(new_record)
    return_date = new_record.get("return_date", "")

    msg = (
        f"📝 ISH HAQI SAQLANMAGAN TA’TIL UZAYTIRILDI\n\n"
        f"👤 Xodim: {new_record.get('fio')}\n"
        f"💼 Lavozim: {new_record.get('position', '')}\n"
        f"📌 Loyiha: {new_record.get('project', '')}\n\n"
        f"📅 Avvalgi muddat: {old_record.get('start')} — {old_record.get('end')}\n"
        f"➕ Qo‘shilgan muddat: {added_start} — {added_end}\n\n"
        f"📋 Umumiy davrlar:\n{format_periods(periods)}\n\n"
        f"⏳ Jami davomiyligi: {new_record.get('days')} kun\n"
        f"📌 Ishga chiqish sanasi: {return_date}\n\n"
        f"ℹ️ Agar oxirgi kun dam olish yoki bayram kuniga to‘g‘ri kelsa, ishga chiqish sanasi keyingi ish kuniga o‘tkazildi."
    )

    await send_group_message(msg)


def employee_keyboard(employees):
    buttons = employees + ["🏠 Старт"]
    return make_keyboard(buttons, cols=2)


def is_active_record(record, today):
    try:
        start = datetime.strptime(record.get("start", ""), "%d.%m.%Y").date()
        end = datetime.strptime(record.get("end", ""), "%d.%m.%Y").date()
        return start <= today <= end
    except:
        return False


def build_report():
    today = datetime.now().date()
    history = load_history()

    groups = {
        "🌴 В отпуске": [],
        "📝 В БС": [],
        "📚 В учебном отпуске": [],
        "🏥 На больничном": []
    }

    for r in history:
        if not is_active_record(r, today):
            continue

        t = r.get("type", "")

        if t in ["🌴 Полный отпуск", "🧩 Часть отпуска", "📌 Оставшийся отпуск"]:
            groups["🌴 В отпуске"].append(r)
        elif t in ["📝 БС с периода по период", "📅 БС на один день"]:
            groups["📝 В БС"].append(r)
        elif t == "📚 Учебный отпуск":
            groups["📚 В учебном отпуске"].append(r)
        elif t == SICK_LEAVE_TYPE:
            groups["🏥 На больничном"].append(r)

    msg = f"📊 Отчет на {datetime.now().strftime('%d.%m.%Y')}\n\n"

    for title, records in groups.items():
        msg += f"{title}:\n"
        if not records:
            msg += "Нет сотрудников\n\n"
            continue

        for i, r in enumerate(records, 1):
            msg += f"{i}. {r.get('fio')}\n"
            msg += f"   Проект: {r.get('project', '')}\n"
            msg += f"   Тип: {r.get('type')}\n"
            if r.get("periods"):
                msg += "   Периоды:\n"
                for p in get_periods_from_record(r):
                    msg += f"   - {p.get('start')} по {p.get('end')}\n"
            else:
                msg += f"   С: {r.get('start')} по {r.get('end')}\n"
            if r.get("days"):
                msg += f"   Дней: {r.get('days')}\n"
            if r.get("return_date"):
                msg += f"   Выход: {r.get('return_date')}\n"
            msg += "\n"

    return msg


async def reminder_loop():
    while True:
        now = datetime.now()
        now_date = now.strftime("%d.%m.%Y")
        now_time = now.strftime("%H:%M")
        tomorrow_date = (now + timedelta(days=1)).strftime("%d.%m.%Y")

        history = load_history()
        sent = load_sent_reminders()

        if now_time == START_REMIND_TIME:
            for r in history:
                if r.get("start") == tomorrow_date:
                    key = f"start_{r.get('fio')}_{r.get('start')}_{r.get('type')}"

                    if key in sent:
                        continue

                    msg = (
                        f"🔔 ERTAGA TA’TIL BOSHLANADI\n\n"
                        f"👤 Xodim: {r.get('fio')}\n"
                        f"📌 Loyiha: {r.get('project', '')}\n\n"
                        f"📝 Ta’til turi: {type_uz(r.get('type'))}\n"
                        f"📅 Muddat: {r.get('start')} — {r.get('end')}\n"
                        f"⏳ Davomiyligi: {r.get('days')} kun\n"
                        f"📌 Ishga chiqish sanasi: {r.get('return_date', '')}\n\n"
                        f"ℹ️ Xodim ertadan boshlab {start_phrase_uz(r.get('type'))}."
                    )

                    await send_group_message(msg)
                    save_sent_reminder(key)

        if now_time in REMIND_TIMES:
            for r in history:
                if r.get("return_date") == now_date:
                    key = f"return_{r.get('fio')}_{now_date}_{now_time}_{r.get('type')}"

                    if key in sent:
                        continue

                    msg = (
                        f"✅ BUGUN ISHGA CHIQADI\n\n"
                        f"👤 Xodim: {r.get('fio')}\n"
                        f"📌 Loyiha: {r.get('project', '')}\n\n"
                        f"📝 Ta’til turi: {type_uz(r.get('type'))}\n"
                        f"📅 Ishga chiqish sanasi: {r.get('return_date', '')}\n\n"
                        f"ℹ️ Xodim bugundan ish faoliyatini davom ettiradi."
                    )

                    await send_group_message(msg)
                    save_sent_reminder(key)

        await asyncio.sleep(30)


menu = make_keyboard([
    "📄 Создать заявление",
    "✈️ Командировка",
    "➕ Добавить запись вручную",
    "🏥 Больничный",
    "📊 Отчет",
    "🗑 Удалить запись",
    "📜 История",
    "🏠 Старт"
], cols=2)

report_only_menu = make_keyboard([
    "📊 Отчет",
    "🏠 Старт"
], cols=2)

history_menu = make_keyboard([
    "🔍 Поиск сотрудника",
    "📋 Полный список сотрудников",
    "🏠 Старт"
], cols=2)

pos_menu = make_keyboard([
    "Инженер программист",
    "Программист",
    "🏠 Старт"
], cols=2)

confirm_delete_menu = make_keyboard([
    "✅ Да, удалить",
    "❌ Нет, отменить"
], cols=2)

manual_type_menu = make_keyboard(MANUAL_TYPES + ["🏠 Старт"], cols=2)


@dp.message(Command("start"))
async def start(m: Message):
    chat_id = m.chat.id

    if chat_id == GROUP_CHAT_ID:
        return

    if not is_allowed(chat_id):
        await m.answer("⛔ Доступ запрещен")
        return

    state[chat_id] = "menu"
    await m.answer("Меню", reply_markup=get_menu(chat_id))


@dp.message()
async def handler(m: Message):
    chat_id = m.chat.id
    text = m.text

    if chat_id == GROUP_CHAT_ID:
        return

    if not is_allowed(chat_id):
        await m.answer("⛔ Доступ запрещен")
        return

    if text == "🏠 Старт":
        state[chat_id] = "menu"
        await m.answer("Меню", reply_markup=get_menu(chat_id))
        return

    if text == "📊 Отчет":
        await m.answer(build_report(), reply_markup=get_menu(chat_id))
        return

    if is_report_only(chat_id):
        await m.answer("У вас доступ только к отчету.", reply_markup=report_only_menu)
        return

    if text == "✈️ Командировка":
        data[chat_id] = {"employees": []}
        state[chat_id] = "trip_count"
        await m.answer("Сколько сотрудников едет в командировку? Напиши цифрой, например: 5")
        return

    if state.get(chat_id) == "trip_count":
        if not text.isdigit():
            await m.answer("Напиши количество сотрудников цифрой.")
            return

        count = int(text)
        if count < 1 or count > 20:
            await m.answer("Количество должно быть от 1 до 20.")
            return

        data[chat_id] = {"employees": [], "trip_count": count, "current_employee": 1}
        state[chat_id] = "trip_search_emp"
        await m.answer(f"Сотрудник 1 из {count}. Напиши минимум 2 буквы фамилии или имени.")
        return

    if state.get(chat_id) == "trip_search_emp":
        employees = load_employees()
        found = search_employees_by_text(text, employees)

        if found is None:
            await m.answer("Напиши минимум 2 буквы фамилии или имени.")
            return

        if not found:
            await m.answer("Сотрудник не найден. Напиши другую часть ФИО.")
            return

        temp_search[chat_id] = found
        state[chat_id] = "trip_choose_emp"
        await m.answer("Выбери сотрудника:", reply_markup=employee_keyboard(found))
        return

    if state.get(chat_id) == "trip_choose_emp":
        if text not in temp_search.get(chat_id, []):
            await m.answer("Выбери сотрудника только из списка кнопок.")
            return

        data[chat_id]["selected_trip_fio"] = text
        state[chat_id] = "trip_pos"
        await m.answer("Должность:", reply_markup=pos_menu)
        return

    if state.get(chat_id) == "trip_pos":
        if text not in ["Инженер программист", "Программист"]:
            await m.answer("Выбери должность из списка.")
            return

        data[chat_id]["employees"].append({
            "fio": data[chat_id].get("selected_trip_fio", ""),
            "pos": text
        })

        current = len(data[chat_id]["employees"])
        total = data[chat_id].get("trip_count", 1)

        if current < total:
            next_num = current + 1
            state[chat_id] = "trip_search_emp"
            await m.answer(
                f"Сотрудник {next_num} из {total}. Напиши минимум 2 буквы фамилии или имени.",
                reply_markup=make_keyboard(["🏠 Старт"], cols=1)
            )
            return

        state[chat_id] = "trip_start_date"
        await m.answer("Введи дату начала командировки ДД.ММ.ГГГГ", reply_markup=make_keyboard(["🏠 Старт"], cols=1))
        return

    if state.get(chat_id) == "trip_start_date":
        if not is_valid_date(text):
            await m.answer("Ошибка. Введи дату начала в формате ДД.ММ.ГГГГ")
            return

        data[chat_id]["start"] = normalize_date(text)
        state[chat_id] = "trip_end_date"
        await m.answer("Введи дату конца командировки ДД.ММ.ГГГГ")
        return

    if state.get(chat_id) == "trip_end_date":
        if not is_valid_date(text):
            await m.answer("Ошибка. Введи дату конца в формате ДД.ММ.ГГГГ")
            return

        start_date = datetime.strptime(data[chat_id]["start"], "%d.%m.%Y")
        end_date = datetime.strptime(normalize_date(text), "%d.%m.%Y")

        if end_date < start_date:
            await m.answer("Дата конца не может быть раньше даты начала. Введи дату конца заново.")
            return

        data[chat_id]["end"] = normalize_date(text)
        paths = create_trip_docs(data[chat_id])

        for path in paths:
            await m.answer_document(FSInputFile(path))

        state[chat_id] = "menu"
        await m.answer("Командировка готова ✅", reply_markup=menu)
        return

    if text == "➕ Добавить запись вручную":
        state[chat_id] = "manual_search_emp"
        await m.answer("Напиши ФИО или часть ФИО сотрудника")
        return

    if state.get(chat_id) == "manual_search_emp":
        employees = load_employees()
        found = search_employees_by_text(text, employees)

        if found is None:
            await m.answer("Напиши минимум 2 буквы фамилии или имени.")
            return

        if not found:
            await m.answer("Сотрудник не найден. Напиши другую часть ФИО.")
            return

        temp_search[chat_id] = found
        state[chat_id] = "manual_choose_emp"
        await m.answer("Выбери сотрудника:", reply_markup=employee_keyboard(found))
        return

    if state.get(chat_id) == "manual_choose_emp":
        if text not in temp_search.get(chat_id, []):
            await m.answer("Выбери сотрудника только из списка кнопок.")
            return

        data[chat_id] = {"fio": text, "pos": ""}
        state[chat_id] = "manual_type"
        await m.answer("Выбери тип записи:", reply_markup=manual_type_menu)
        return

    if state.get(chat_id) == "manual_type":
        if text not in MANUAL_TYPES:
            await m.answer("Выбери тип записи из списка.")
            return

        data[chat_id]["type"] = text
        state[chat_id] = "manual_pos"
        await m.answer("Должность:", reply_markup=pos_menu)
        return

    if state.get(chat_id) == "manual_pos":
        if text not in ["Инженер программист", "Программист"]:
            await m.answer("Выбери должность из списка.")
            return

        data[chat_id]["pos"] = text
        state[chat_id] = "manual_project"
        await m.answer("Напиши название проекта")
        return

    if state.get(chat_id) == "manual_project":
        data[chat_id]["project"] = text
        state[chat_id] = "manual_start"
        await m.answer("Введи дату начала ДД.ММ.ГГГГ")
        return

    if state.get(chat_id) == "manual_start":
        if not is_valid_date(text):
            await m.answer("Ошибка. Введи дату в формате ДД.ММ.ГГГГ")
            return

        data[chat_id]["start"] = normalize_date(text)

        if data[chat_id]["type"] == "📅 БС на один день":
            data[chat_id]["end"] = normalize_date(text)
            data[chat_id]["days"] = "1"
            save_history(data[chat_id])
            state[chat_id] = "menu"
            await m.answer("Запись вручную добавлена ✅", reply_markup=menu)
            return

        state[chat_id] = "manual_end"
        await m.answer("Введи дату конца ДД.ММ.ГГГГ")
        return

    if state.get(chat_id) == "manual_end":
        if not is_valid_date(text):
            await m.answer("Ошибка. Введи дату конца в формате ДД.ММ.ГГГГ")
            return

        start_date = datetime.strptime(data[chat_id]["start"], "%d.%m.%Y")
        end_date = datetime.strptime(normalize_date(text), "%d.%m.%Y")

        if end_date < start_date:
            await m.answer("Дата конца не может быть раньше даты начала. Введи дату конца заново.")
            return

        days = (end_date - start_date).days + 1
        data[chat_id]["end"] = normalize_date(text)
        data[chat_id]["days"] = str(days)

        save_history(data[chat_id])
        state[chat_id] = "menu"
        await m.answer("Запись вручную добавлена ✅", reply_markup=menu)
        return

    if text == "🗑 Удалить запись":
        state[chat_id] = "delete_search_emp"
        await m.answer("Напиши ФИО или часть ФИО сотрудника")
        return

    if state.get(chat_id) == "delete_search_emp":
        employees = load_employees()
        found = search_employees_by_text(text, employees)

        if found is None:
            await m.answer("Напиши минимум 2 буквы фамилии или имени.")
            return

        if not found:
            await m.answer("Сотрудник не найден. Напиши другую часть ФИО.")
            return

        temp_search[chat_id] = found
        state[chat_id] = "delete_choose_emp"
        await m.answer("Выбери сотрудника:", reply_markup=employee_keyboard(found))
        return

    if state.get(chat_id) == "delete_choose_emp":
        if text not in temp_search.get(chat_id, []):
            await m.answer("Выбери сотрудника только из списка кнопок.")
            return

        history = load_history()
        records = [r for r in history if r.get("fio") == text]

        if not records:
            state[chat_id] = "menu"
            await m.answer("У этого сотрудника нет записей в истории.", reply_markup=menu)
            return

        data[chat_id] = {
            "delete_fio": text,
            "delete_records": records
        }

        buttons = []
        msg = f"🗑 Записи сотрудника:\n\n👤 {text}\n\n"

        for i, r in enumerate(records, 1):
            msg += f"{i}. {r.get('type')}\n"
            msg += f"   Проект: {r.get('project', '')}\n"
            if r.get("periods"):
                msg += "   Периоды:\n"
                for p in get_periods_from_record(r):
                    msg += f"   - {p.get('start')} по {p.get('end')}\n"
            else:
                msg += f"   С: {r.get('start')} по {r.get('end')}\n"
            if r.get("days"):
                msg += f"   Дней: {r.get('days')}\n"
            if r.get("return_date"):
                msg += f"   Выход: {r.get('return_date')}\n"
            msg += f"   Создан: {r.get('created_at')}\n\n"
            buttons.append(str(i))

        buttons.append("🏠 Старт")

        state[chat_id] = "delete_choose_record"
        await m.answer(
            msg + "Выбери номер записи для удаления:",
            reply_markup=make_keyboard(buttons, cols=3)
        )
        return

    if state.get(chat_id) == "delete_choose_record":
        if not text.isdigit():
            await m.answer("Выбери номер записи кнопкой.")
            return

        idx = int(text) - 1
        records = data[chat_id].get("delete_records", [])

        if idx < 0 or idx >= len(records):
            await m.answer("Неверный номер записи.")
            return

        record = records[idx]
        data[chat_id]["delete_record"] = record

        msg = "Ты точно хочешь удалить эту запись?\n\n"
        msg += f"👤 {record.get('fio')}\n"
        msg += f"Проект: {record.get('project', '')}\n"
        msg += f"Тип: {record.get('type')}\n"
        msg += f"С: {record.get('start')} по {record.get('end')}\n"
        msg += f"Дней: {record.get('days')}\n"
        msg += f"Выход: {record.get('return_date')}\n"

        state[chat_id] = "delete_confirm"
        await m.answer(msg, reply_markup=confirm_delete_menu)
        return

    if state.get(chat_id) == "delete_confirm":
        if text == "❌ Нет, отменить":
            state[chat_id] = "menu"
            await m.answer("Удаление отменено.", reply_markup=menu)
            return

        if text != "✅ Да, удалить":
            await m.answer("Выбери: ✅ Да, удалить или ❌ Нет, отменить")
            return

        record_to_delete = data[chat_id].get("delete_record")
        history = load_history()

        new_history = []
        deleted = False

        for r in history:
            if not deleted and r == record_to_delete:
                deleted = True
                continue
            new_history.append(r)

        save_history_full(new_history)

        state[chat_id] = "menu"
        await m.answer("Запись удалена ✅", reply_markup=menu)
        return

    if text == "🏥 Больничный":
        state[chat_id] = "sick_search_emp"
        await m.answer("Напиши ФИО или часть ФИО сотрудника")
        return

    if state.get(chat_id) == "sick_search_emp":
        if len(str(text).lower().strip()) < 2:
            await m.answer("Напиши минимум 2 буквы фамилии или имени.")
            return

        active_sick_records = find_active_sick_records_by_text(text)

        if len(active_sick_records) == 1:
            old_record = active_sick_records[0]
            data[chat_id] = {
                "fio": old_record.get("fio"),
                "type": SICK_LEAVE_TYPE,
                "old_sick_record": old_record
            }
            state[chat_id] = "sick_extend_start_date"

            msg = (
                f"✅ Найдена активная запись больничного:\n\n"
                f"👤 {old_record.get('fio')}\n"
                f"Проект: {old_record.get('project', '')}\n"
            )

            if old_record.get("periods"):
                msg += "Периоды:\n"
                for p in get_periods_from_record(old_record):
                    msg += f"- {p.get('start')} по {p.get('end')}\n"
            else:
                msg += f"С: {old_record.get('start')} по {old_record.get('end')}\n"

            msg += (
                f"Дней: {old_record.get('days')}\n"
                f"Выход: {old_record.get('return_date')}\n\n"
                f"Теперь введи дату, С КОТОРОЙ продлеваем больничный: ДД.ММ.ГГГГ"
            )

            await m.answer(msg)
            return

        if len(active_sick_records) > 1:
            found = []
            for r in active_sick_records:
                fio = r.get("fio")
                if fio and fio not in found:
                    found.append(fio)

            temp_search[chat_id] = found
            state[chat_id] = "sick_choose_emp"
            await m.answer("Найдено несколько активных больничных. Выбери сотрудника:", reply_markup=employee_keyboard(found))
            return

        employees = load_employees()
        found = search_employees_by_text(text, employees)

        if found is None:
            await m.answer("Напиши минимум 2 буквы фамилии или имени.")
            return

        if not found:
            await m.answer("Сотрудник не найден. Напиши другую часть ФИО.")
            return

        if len(found) == 1:
            fio = found[0]
            data[chat_id] = {
                "fio": fio,
                "type": SICK_LEAVE_TYPE,
                "pos": ""
            }
            state[chat_id] = "sick_project"
            await m.answer(
                f"Активный больничный по сотруднику {fio} не найден.\n"
                f"Создаем новую запись больничного. Напиши название проекта"
            )
            return

        temp_search[chat_id] = found
        state[chat_id] = "sick_choose_emp"
        await m.answer("Активный больничный не найден. Выбери сотрудника для новой записи:", reply_markup=employee_keyboard(found))
        return

    if state.get(chat_id) == "sick_choose_emp":
        if text not in temp_search.get(chat_id, []):
            await m.answer("Выбери сотрудника только из списка кнопок.")
            return

        active_sick_records = find_active_sick_records_by_fio(text)

        if active_sick_records:
            old_record = active_sick_records[0]
            data[chat_id] = {
                "fio": text,
                "type": SICK_LEAVE_TYPE,
                "old_sick_record": old_record
            }
            state[chat_id] = "sick_extend_start_date"

            msg = (
                f"✅ Найдена активная запись больничного:\n\n"
                f"👤 {old_record.get('fio')}\n"
                f"Проект: {old_record.get('project', '')}\n"
            )

            if old_record.get("periods"):
                msg += "Периоды:\n"
                for p in get_periods_from_record(old_record):
                    msg += f"- {p.get('start')} по {p.get('end')}\n"
            else:
                msg += f"С: {old_record.get('start')} по {old_record.get('end')}\n"

            msg += (
                f"Дней: {old_record.get('days')}\n"
                f"Выход: {old_record.get('return_date')}\n\n"
                f"Теперь введи дату, С КОТОРОЙ продлеваем больничный: ДД.ММ.ГГГГ"
            )

            await m.answer(msg)
            return

        data[chat_id] = {
            "fio": text,
            "type": SICK_LEAVE_TYPE,
            "pos": ""
        }
        state[chat_id] = "sick_project"
        await m.answer("Активный больничный не найден. Создаем новую запись. Напиши название проекта")
        return

    if state.get(chat_id) == "sick_extend_start_date":
        if not is_valid_date(text):
            await m.answer("Ошибка. Введи дату начала продления в формате ДД.ММ.ГГГГ")
            return

        old_record = data[chat_id].get("old_sick_record")
        old_end = datetime.strptime(old_record.get("end"), "%d.%m.%Y")
        new_start = datetime.strptime(normalize_date(text), "%d.%m.%Y")

        if new_start <= old_end:
            await m.answer(
                f"Дата продления должна быть после текущей даты окончания больничного: {old_record.get('end')}.\n"
                f"Введи дату начала продления заново."
            )
            return

        data[chat_id]["extend_start"] = normalize_date(text)
        state[chat_id] = "sick_extend_end_date"
        await m.answer("Введи дату, ДО КОТОРОЙ продлеваем больничный: ДД.ММ.ГГГГ")
        return

    if state.get(chat_id) == "sick_extend_end_date":
        if not is_valid_date(text):
            await m.answer("Ошибка. Введи дату конца продления в формате ДД.ММ.ГГГГ")
            return

        start_date = datetime.strptime(data[chat_id]["extend_start"], "%d.%m.%Y")
        end_date = datetime.strptime(normalize_date(text), "%d.%m.%Y")

        if end_date < start_date:
            await m.answer("Дата конца не может быть раньше даты начала. Введи дату конца заново.")
            return

        old_record = data[chat_id].get("old_sick_record")
        new_record = build_sick_extended_record(
            old_record,
            data[chat_id]["extend_start"],
            normalize_date(text)
        )

        update_history_record(old_record, new_record)
        await notify_sick_extended(
            old_record,
            new_record,
            data[chat_id]["extend_start"],
            normalize_date(text)
        )

        state[chat_id] = "menu"
        await m.answer(
            f"Больничный продлен ✅\n\n"
            f"👤 {new_record.get('fio')}\n"
            f"Всего дней больничного: {new_record.get('days')}\n"
            f"Выход на работу: {new_record.get('return_date')}",
            reply_markup=menu
        )
        return

    if state.get(chat_id) == "sick_project":
        data[chat_id]["project"] = text
        state[chat_id] = "sick_start_date"
        await m.answer("Введи дату начала больничного ДД.ММ.ГГГГ")
        return

    if state.get(chat_id) == "sick_start_date":
        if not is_valid_date(text):
            await m.answer("Ошибка. Введи дату в формате ДД.ММ.ГГГГ")
            return

        data[chat_id]["start"] = normalize_date(text)
        state[chat_id] = "sick_end_date"
        await m.answer("Введи примерную дату конца больничного ДД.ММ.ГГГГ")
        return

    if state.get(chat_id) == "sick_end_date":
        if not is_valid_date(text):
            await m.answer("Ошибка. Введи дату конца в формате ДД.ММ.ГГГГ")
            return

        start_date = datetime.strptime(data[chat_id]["start"], "%d.%m.%Y")
        end_date = datetime.strptime(normalize_date(text), "%d.%m.%Y")

        if end_date < start_date:
            await m.answer("Дата конца не может быть раньше даты начала. Введи дату конца заново.")
            return

        days = (end_date - start_date).days + 1
        data[chat_id]["end"] = normalize_date(text)
        data[chat_id]["days"] = str(days)
        data[chat_id]["periods"] = [{
            "start": data[chat_id]["start"],
            "end": normalize_date(text)
        }]

        save_history(data[chat_id])
        state[chat_id] = "menu"
        await m.answer("Больничный сохранен в историю.", reply_markup=menu)
        return

    if state.get(chat_id) == "bs_extend_start_date":
        if not is_valid_date(text):
            await m.answer("Ошибка. Введи дату начала продления БС в формате ДД.ММ.ГГГГ")
            return

        old_record = data[chat_id].get("old_bs_record")
        old_end = datetime.strptime(old_record.get("end"), "%d.%m.%Y")
        new_start = datetime.strptime(normalize_date(text), "%d.%m.%Y")

        if new_start <= old_end:
            await m.answer(
                f"Дата продления должна быть после текущей даты окончания БС: {old_record.get('end')}.\n"
                f"Введи дату начала продления заново."
            )
            return

        data[chat_id]["extend_start"] = normalize_date(text)
        state[chat_id] = "bs_extend_end_date"
        await m.answer("Введи дату, ДО КОТОРОЙ продлеваем БС: ДД.ММ.ГГГГ")
        return

    if state.get(chat_id) == "bs_extend_end_date":
        if not is_valid_date(text):
            await m.answer("Ошибка. Введи дату конца продления БС в формате ДД.ММ.ГГГГ")
            return

        start_date = datetime.strptime(data[chat_id]["extend_start"], "%d.%m.%Y")
        end_date = datetime.strptime(normalize_date(text), "%d.%m.%Y")

        if end_date < start_date:
            await m.answer("Дата конца не может быть раньше даты начала. Введи дату конца заново.")
            return

        old_record = data[chat_id].get("old_bs_record")
        new_record = build_bs_extended_record(
            old_record,
            data[chat_id]["extend_start"],
            normalize_date(text)
        )

        update_history_record(old_record, new_record)
        await notify_bs_extended(
            old_record,
            new_record,
            data[chat_id]["extend_start"],
            normalize_date(text)
        )

        state[chat_id] = "menu"
        await m.answer(
            f"БС продлен ✅\n\n"
            f"👤 {new_record.get('fio')}\n"
            f"Всего дней БС: {new_record.get('days')}\n"
            f"Выход на работу: {new_record.get('return_date')}",
            reply_markup=menu
        )
        return

    if text == "📄 Создать заявление":
        state[chat_id] = "search_emp"
        await m.answer("Напиши ФИО или часть ФИО сотрудника")
        return

    if state.get(chat_id) == "search_emp":
        employees = load_employees()
        found = search_employees_by_text(text, employees)

        if found is None:
            await m.answer("Напиши минимум 2 буквы фамилии или имени.")
            return

        if not found:
            await m.answer("Сотрудник не найден. Напиши другую часть ФИО.")
            return

        temp_search[chat_id] = found
        state[chat_id] = "choose_emp"
        await m.answer("Выбери сотрудника:", reply_markup=employee_keyboard(found))
        return

    if state.get(chat_id) == "choose_emp":
        if text not in temp_search.get(chat_id, []):
            await m.answer("Выбери сотрудника только из списка кнопок.")
            return

        data[chat_id] = {"fio": text}
        state[chat_id] = "type"

        await m.answer(
            "Выбери тип заявления:",
            reply_markup=make_keyboard(list(TEMPLATES.keys()) + ["🏠 Старт"], cols=2)
        )
        return

    if state.get(chat_id) == "type":
        if text not in TEMPLATES:
            await m.answer("Выбери тип заявления из списка.")
            return

        data[chat_id]["type"] = text

        if text == "🧩 Часть отпуска":
            previous_part_records = find_previous_part_leave_records_by_fio(data[chat_id].get("fio"))

            if previous_part_records:
                data[chat_id]["type"] = "📌 Оставшийся отпуск"

                msg = (
                    "ℹ️ У этого сотрудника уже есть запись по части отпуска.\n\n"
                    f"👤 {data[chat_id].get('fio')}\n"
                    f"{format_leave_records_for_message(previous_part_records)}\n"
                    "✅ Тип заявления автоматически изменен на: 📌 Оставшийся отпуск"
                )

                await m.answer(msg)

        if text in BS_LEAVE_TYPES:
            active_bs_records = find_active_bs_records_by_fio(data[chat_id].get("fio"))

            if active_bs_records:
                old_record = active_bs_records[0]
                data[chat_id] = {
                    "fio": old_record.get("fio"),
                    "type": BS_RANGE_TYPE,
                    "old_bs_record": old_record
                }
                state[chat_id] = "bs_extend_start_date"

                msg = (
                    f"✅ Найдена активная запись БС:\n\n"
                    f"👤 {old_record.get('fio')}\n"
                    f"Проект: {old_record.get('project', '')}\n"
                )

                if old_record.get("periods"):
                    msg += "Периоды:\n"
                    for p in get_periods_from_record(old_record):
                        msg += f"- {p.get('start')} по {p.get('end')}\n"
                else:
                    msg += f"С: {old_record.get('start')} по {old_record.get('end')}\n"

                msg += (
                    f"Дней: {old_record.get('days')}\n"
                    f"Выход: {old_record.get('return_date')}\n\n"
                    f"Теперь введи дату, С КОТОРОЙ продлеваем БС: ДД.ММ.ГГГГ"
                )

                await m.answer(msg)
                return

        state[chat_id] = "pos"
        await m.answer("Должность:", reply_markup=pos_menu)
        return

    if state.get(chat_id) == "pos":
        if text not in ["Инженер программист", "Программист"]:
            await m.answer("Выбери должность из списка.")
            return

        data[chat_id]["pos"] = text
        state[chat_id] = "project"
        await m.answer("Напиши название проекта")
        return

    if state.get(chat_id) == "project":
        data[chat_id]["project"] = text

        if data[chat_id]["type"] == "👶 Мат помощь (ребенок)":
            path = finish_and_send(chat_id)
            await m.answer_document(FSInputFile(path))
            await notify_application_created(data[chat_id])
            state[chat_id] = "menu"
            await m.answer("Готово", reply_markup=menu)
            return

        state[chat_id] = "date"
        await m.answer("Введи дату начала ДД.ММ.ГГГГ")
        return

    if state.get(chat_id) == "date":
        if not is_valid_date(text):
            await m.answer("Ошибка. Введи дату в формате ДД.ММ.ГГГГ")
            return

        data[chat_id]["start"] = normalize_date(text)
        start_date = datetime.strptime(normalize_date(text), "%d.%m.%Y")
        doc_type = data[chat_id]["type"]

        if doc_type == "🌴 Полный отпуск":
            days = 30
            end_date = start_date + timedelta(days=days - 1)
            data[chat_id]["days"] = str(days)
            data[chat_id]["end"] = end_date.strftime("%d.%m.%Y")
            path = finish_and_send(chat_id)
            await m.answer_document(FSInputFile(path))
            await notify_application_created(data[chat_id])
            state[chat_id] = "menu"
            await m.answer("Готово", reply_markup=menu)
            return

        if doc_type == "🧩 Часть отпуска":
            days = 15
            end_date = start_date + timedelta(days=days - 1)
            data[chat_id]["days"] = str(days)
            data[chat_id]["end"] = end_date.strftime("%d.%m.%Y")
            path = finish_and_send(chat_id)
            await m.answer_document(FSInputFile(path))
            await notify_application_created(data[chat_id])
            state[chat_id] = "menu"
            await m.answer("Готово", reply_markup=menu)
            return

        if doc_type == "📚 Учебный отпуск":
            state[chat_id] = "study_end_date"
            await m.answer("Введи дату конца учебного отпуска ДД.ММ.ГГГГ")
            return

        if doc_type == "📅 БС на один день":
            data[chat_id]["days"] = "1"
            data[chat_id]["end"] = normalize_date(text)
            path = finish_and_send(chat_id)
            await m.answer_document(FSInputFile(path))
            await notify_application_created(data[chat_id])
            state[chat_id] = "menu"
            await m.answer("Готово", reply_markup=menu)
            return

        if doc_type == "💍 Мат помощь (свадьба)":
            days = 3
            end_date = start_date + timedelta(days=days - 1)
            data[chat_id]["days"] = str(days)
            data[chat_id]["end"] = end_date.strftime("%d.%m.%Y")
            path = finish_and_send(chat_id)
            await m.answer_document(FSInputFile(path))
            await notify_application_created(data[chat_id])
            state[chat_id] = "menu"
            await m.answer("Готово", reply_markup=menu)
            return

        state[chat_id] = "days"
        await m.answer("Количество дней")
        return

    if state.get(chat_id) == "study_end_date":
        if not is_valid_date(text):
            await m.answer("Ошибка. Введи дату конца в формате ДД.ММ.ГГГГ")
            return

        start_date = datetime.strptime(data[chat_id]["start"], "%d.%m.%Y")
        end_date = datetime.strptime(normalize_date(text), "%d.%m.%Y")

        if end_date < start_date:
            await m.answer("Дата конца не может быть раньше даты начала. Введи дату конца заново.")
            return

        days = (end_date - start_date).days + 1
        data[chat_id]["end"] = normalize_date(text)
        data[chat_id]["days"] = str(days)

        path = finish_and_send(chat_id)
        await m.answer_document(FSInputFile(path))
        await notify_application_created(data[chat_id])
        state[chat_id] = "menu"
        await m.answer("Готово", reply_markup=menu)
        return

    if state.get(chat_id) == "days":
        if not text.isdigit():
            await m.answer("Введи количество дней цифрой.")
            return

        days = int(text)
        start_date = datetime.strptime(data[chat_id]["start"], "%d.%m.%Y")
        end_date = start_date + timedelta(days=days - 1)

        data[chat_id]["days"] = str(days)
        data[chat_id]["end"] = end_date.strftime("%d.%m.%Y")

        path = finish_and_send(chat_id)
        await m.answer_document(FSInputFile(path))
        await notify_application_created(data[chat_id])
        state[chat_id] = "menu"
        await m.answer("Готово", reply_markup=menu)
        return

    if text == "📜 История":
        state[chat_id] = "history_menu"
        await m.answer("Выбери действие:", reply_markup=history_menu)
        return

    if text == "🔍 Поиск сотрудника":
        state[chat_id] = "history_search"
        await m.answer("Напиши ФИО или часть ФИО сотрудника")
        return

    if text == "📋 Полный список сотрудников":
        employees = load_employees()
        temp_search[chat_id] = employees
        state[chat_id] = "history_choose"
        await m.answer("Выбери сотрудника:", reply_markup=employee_keyboard(employees))
        return

    if state.get(chat_id) == "history_search":
        employees = load_employees()
        found = search_employees_by_text(text, employees)

        if found is None:
            await m.answer("Напиши минимум 2 буквы фамилии или имени.")
            return

        if not found:
            await m.answer("Сотрудник не найден. Попробуй ещё раз.")
            return

        temp_search[chat_id] = found
        state[chat_id] = "history_choose"
        await m.answer("Выбери сотрудника:", reply_markup=employee_keyboard(found))
        return

    if state.get(chat_id) == "history_choose":
        if text not in temp_search.get(chat_id, []):
            await m.answer("Выбери сотрудника только из списка кнопок.")
            return

        history = load_history()
        records = [r for r in history if r.get("fio") == text]

        if not records:
            await m.answer("По этому сотруднику история пустая.", reply_markup=menu)
            state[chat_id] = "menu"
            return

        msg = ""
        for r in records:
            msg += f"ФИО: {r.get('fio')}\n"
            msg += f"Тип: {r.get('type')}\n"
            msg += f"Должность: {r.get('position')}\n"
            msg += f"Проект: {r.get('project', '')}\n"

            if r.get("periods"):
                msg += "Периоды:\n"
                for p in get_periods_from_record(r):
                    msg += f"- {p.get('start')} по {p.get('end')}\n"
            else:
                if r.get("start"):
                    msg += f"Начало: {r.get('start')}\n"

                if r.get("end"):
                    msg += f"Конец: {r.get('end')}\n"

            if r.get("days"):
                msg += f"Дней: {r.get('days')}\n"

            if r.get("return_date"):
                msg += f"Выход на работу: {r.get('return_date')}\n"

            msg += f"Создан: {r.get('created_at')}\n\n"

        await m.answer(msg, reply_markup=menu)
        state[chat_id] = "menu"
        return


async def main():
    normalize_history_file()
    asyncio.create_task(reminder_loop())
    await dp.start_polling(bot)

    
from flask import Flask
import threading
import os

app = Flask(__name__)

@app.route("/")
def home():
    return "OK"

def run_web():
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)


if __name__ == "__main__":
    # запускаем веб в отдельном потоке
    threading.Thread(target=run_web).start()

    # запускаем бота
    asyncio.run(main())
