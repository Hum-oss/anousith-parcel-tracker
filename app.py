# -*- coding: utf-8 -*-
"""
ระบบติดตามพัสดุ อนุสิต — Flask backend
อ่าน/เขียนชีต "เคส" (และ "ประวัติการติดตาม" / "ผู้ใช้งาน") ในสเปรดชีต parcel-tracking-system
โดยตรง — ชีตเดียวกับที่ Apps Script (Code_2.gs) ใช้อยู่ ห้ามลบ/ย้าย/สร้างชีตใหม่

สำคัญ: Apps Script เดิม (sync จาก postinanusit ทุกนาที + อีเมลรายงาน 08:00 น.) ยังรันอยู่คู่ขนาน
Flask นี้ไม่ยุ่งกับ trigger ของ Apps Script เลย แค่อ่าน/เขียนชีตเดียวกัน

รันด้วย (local dev):
    pip install -r requirements.txt
    cp .env.example .env   # แล้วกรอกค่าจริง
    python app.py

รันจริง (production): ดู README.md — deploy ผ่าน Render.com ด้วย gunicorn, single worker
(-w 1) เพราะใช้ threading.Lock ในหน่วยความจำสำหรับ gen เลขแจ้งเรื่อง
"""

import os
import io
import json
import time
import hashlib
import secrets
import threading
from datetime import datetime, timedelta, timezone
from functools import wraps

# โหลดค่าจาก .env สำหรับรันบนเครื่อง (local dev) เท่านั้น — บน Render ตัวแปรมาจาก
# Environment Variables ของ Render โดยตรงอยู่แล้ว ไฟล์ .env จะไม่ถูกอัปโหลดขึ้น GitHub เลย (ดู .gitignore)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from flask import Flask, request, jsonify, render_template, session, send_file

import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook

# ===================== Config =====================

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]  # ไอดีของ parcel-tracking-system

_creds_raw = os.environ["GOOGLE_CREDENTIALS_JSON"]  # เนื้อไฟล์ service_account.json ทั้งไฟล์ (เป็น JSON string)
_creds_info = json.loads(_creds_raw)
_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
_creds = Credentials.from_service_account_info(_creds_info, scopes=_SCOPES)
_gc = gspread.authorize(_creds)

# โฟลเดอร์ Google Drive สำหรับอัปโหลดรูปภาพเคสใหม่/แก้ไขจากหน้าเว็บ
# ไม่บังคับต้องตั้งตอนสตาร์ทแอป (ยังใช้งานฟีเจอร์อื่นได้ตามปกติ) แต่ endpoint อัปโหลดรูปจะแจ้ง error ชัดเจนถ้ายังไม่ตั้ง
DRIVE_UPLOAD_FOLDER_ID = os.environ.get("DRIVE_UPLOAD_FOLDER_ID", "").strip()

# ข้อมูล OAuth ของบัญชี Google จริงที่เป็นเจ้าของโฟลเดอร์อัปโหลดรูปภาพ (เช่น chinawat.new@gmail.com)
# ใช้แทน Service Account ตอนอัปโหลดไฟล์ขึ้น Drive โดยเฉพาะ เพราะ Service Account ไม่มี storage quota
# เป็นของตัวเอง (ข้อจำกัดของ Google) สร้างไฟล์ใหม่ไม่ได้แม้จะมีสิทธิ์ Editor ในโฟลเดอร์ที่แชร์ให้ก็ตาม
# ดูวิธีขอค่าทั้งสามตัวนี้ใน README.md หัวข้อ "อัปโหลดรูปภาพเคส (Drive)" — ได้มาจากการรัน
# get_drive_refresh_token.py ครั้งเดียวบนเครื่อง (ไม่เกี่ยวกับ Service Account เลย)
DRIVE_OAUTH_CLIENT_ID = os.environ.get("GOOGLE_OAUTH_CLIENT_ID", "").strip()
DRIVE_OAUTH_CLIENT_SECRET = os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET", "").strip()
DRIVE_OAUTH_REFRESH_TOKEN = os.environ.get("GOOGLE_OAUTH_REFRESH_TOKEN", "").strip()


def _open_spreadsheet_with_retry(gc, spreadsheet_id, attempts=5, base_delay=2):
    """เชื่อมสเปรดชีตตอนสตาร์ทแอป พร้อม retry แบบ exponential backoff —
    กัน Google Sheets API แผ่วชั่วคราว (เช่น 503) ตอน deploy ทำให้แอปทั้งตัวล่มไปเปล่าๆ
    ทั้งที่รอบ deploy ถัดไปมักจะผ่านปกติอยู่แล้ว"""
    last_err = None
    for attempt in range(attempts):
        try:
            return gc.open_by_key(spreadsheet_id)
        except Exception as e:  # noqa: BLE001 — ตั้งใจดักทุก exception ตอน retry การเชื่อมต่อภายนอก
            last_err = e
            if attempt < attempts - 1:
                time.sleep(base_delay * (2 ** attempt))
    raise last_err


_spreadsheet = _open_spreadsheet_with_retry(_gc, SPREADSHEET_ID)

BKK_TZ = timezone(timedelta(hours=7))

SHEET_CASES = "เคส"
SHEET_LOG = "ประวัติการติดตาม"
SHEET_USERS = "ผู้ใช้งาน"

# ต้องตรงกับลำดับคอลัมน์จริงในชีต "เคส" (ดู setupSheets() ใน Code_2.gs)
CASE_COLS = [
    "ticketNo", "parcelNo", "company", "source", "complainantName", "phone",
    "detail", "photoUrl", "openedAt", "status", "lastChannel", "lastStaff",
    "updatedAt", "closedAt",
]

STATUSES = ["รอติดตาม", "กำลังติดตาม", "ส่งสาขา", "ส่งสาขา/ปิดเรื่อง"]

app = Flask(__name__)
app.secret_key = os.environ["SECRET_KEY"]
app.permanent_session_lifetime = timedelta(hours=12)
# กันคำขออัปโหลดรูปภาพใหญ่เกินไปตั้งแต่ระดับ Flask เอง (ก่อนอ่านเข้าหน่วยความจำทั้งไฟล์) — ดู MAX_PHOTO_BYTES
app.config["MAX_CONTENT_LENGTH"] = 9 * 1024 * 1024

_ws_cache = {}
_ws_lock = threading.Lock()
_ticket_lock = threading.Lock()


def ws(name):
    """แคช worksheet handle ไว้ ลดการเรียก API ซ้ำ"""
    with _ws_lock:
        if name not in _ws_cache:
            _ws_cache[name] = _spreadsheet.worksheet(name)
        return _ws_cache[name]


def now_str():
    return datetime.now(BKK_TZ).strftime("%Y-%m-%d %H:%M:%S")


def parse_dt(s):
    if not s:
        return None
    try:
        return datetime.strptime(str(s)[:19], "%Y-%m-%d %H:%M:%S")
    except ValueError:
        try:
            return datetime.strptime(str(s)[:10], "%Y-%m-%d")
        except ValueError:
            return None


# ===================== Password hashing =====================
# ต้องตรงกับ hashPassword() ใน Code_2.gs เป๊ะๆ (SHA-256 ของ "salt:password")
# เพื่อให้บัญชี/รหัสผ่านที่มีอยู่แล้วในชีต "ผู้ใช้งาน" ใช้ได้กับทั้งสองระบบ โดยไม่ต้องรีเซ็ต

def hash_password(password, salt):
    return hashlib.sha256(f"{salt}:{password}".encode("utf-8")).hexdigest()


def find_user(username):
    rows = ws(SHEET_USERS).get_all_values()
    uname = (username or "").strip().lower()
    for i, row in enumerate(rows[1:], start=2):
        if row and row[0].strip().lower() == uname:
            row = row + [""] * (5 - len(row))
            return {"row": i, "username": row[0].strip(), "nickname": row[1], "role": row[2], "salt": row[3], "hash": row[4]}
    return None


# ===================== Auth decorators (session cookie, ไม่ใช้ sessionToken แบบ Apps Script) =====================

def staff_required(fn):
    @wraps(fn)
    def wrapper(*a, **kw):
        if not session.get("username") or session.get("role") not in ("staff", "admin"):
            return jsonify(ok=False, error="กรุณาเข้าสู่ระบบ"), 401
        return fn(*a, **kw)
    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*a, **kw):
        if session.get("role") != "admin":
            return jsonify(ok=False, error="ต้องมีสิทธิ์ผู้ดูแลระบบ (admin) เท่านั้น"), 403
        return fn(*a, **kw)
    return wrapper


# ===================== Case helpers =====================

def case_row_to_dict(row):
    row = list(row) + [""] * (len(CASE_COLS) - len(row))
    return dict(zip(CASE_COLS, row))


def get_all_cases():
    rows = ws(SHEET_CASES).get_all_values()
    return [case_row_to_dict(r) for r in rows[1:] if r and r[0]]


def find_case(key):
    key = (key or "").strip().lower()
    if not key:
        return None, None
    rows = ws(SHEET_CASES).get_all_values()
    for i, row in enumerate(rows[1:], start=2):
        if not row:
            continue
        ticket = (row[0] or "").strip().lower()
        parcel = (row[1] or "").strip().lower() if len(row) > 1 else ""
        if ticket == key or parcel == key:
            return i, case_row_to_dict(row)
    return None, None


def append_log(ticket_no, channel, status, note, staff_username, staff_nickname):
    ws(SHEET_LOG).append_row(
        [ticket_no, now_str(), channel, status, note, staff_username, staff_nickname],
        value_input_option="USER_ENTERED",
    )


def get_case_history(ticket_no):
    rows = ws(SHEET_LOG).get_all_values()
    out = []
    for row in rows[1:]:
        if row and row[0].strip() == str(ticket_no).strip():
            row = row + [""] * (7 - len(row))
            out.append({
                "datetime": row[1], "channel": row[2], "status": row[3], "note": row[4],
                "staffUsername": row[5], "staffNickname": row[6],
            })
    return out


# เลขแจ้งเรื่องรูปแบบ RQ260819-001 — อ่านค่าปัจจุบันจากชีตสดทุกครั้งก่อนสร้าง (ไม่ใช้ตัวนับแยกในหน่วยความจำ)
# เพราะ Apps Script เดิมก็สร้างเลขจากชีตเดียวกันนี้แบบขนานกันอยู่ (sync จาก postinanusit ทุกนาที)
# threading.Lock ที่นี่กันแค่คำขอซ้อนกันฝั่ง Flask เอง — ไม่ได้กันชนกับ Apps Script 100%
# (ระบบนี้ออกแบบให้รันเป็น single worker process เดียวเท่านั้น ด้วยเหตุผลเดียวกัน)
def generate_ticket_no():
    prefix = "RQ" + datetime.now(BKK_TZ).strftime("%y%m%d")
    with _ticket_lock:
        col = ws(SHEET_CASES).col_values(1)
        nums = []
        for v in col:
            if v.startswith(prefix + "-"):
                tail = v[len(prefix) + 1:]
                if tail.isdigit():
                    nums.append(int(tail))
        nxt = (max(nums) + 1) if nums else 1
        return f"{prefix}-{nxt:03d}"


# ===================== Routes: Frontend =====================

@app.get("/")
def index():
    return render_template("index.html")


# ===================== Routes: Public =====================

@app.get("/api/search")
def api_search():
    q = request.args.get("q", "").strip()
    if not q:
        return jsonify(ok=False, error="กรุณาระบุเลขพัสดุหรือเลขแจ้งเรื่อง"), 400
    _, case = find_case(q)
    if not case:
        return jsonify(ok=False, error="ไม่พบข้อมูลในระบบ")
    # ตามคำขอของเจ้าของกิจการ: แสดงชื่อ/ที่อยู่/เบอร์ผู้แจ้งเรื่องในผลค้นหาสาธารณะด้วย
    # (เดิมตั้งใจไม่คืนค่านี้เพื่อความเป็นส่วนตัว แต่ธุรกิจต้องการให้ลูกค้าที่ค้นหาเห็นข้อมูลนี้)
    return jsonify(ok=True, case={
        "ticketNo": case["ticketNo"], "parcelNo": case["parcelNo"], "company": case["company"],
        "status": case["status"], "lastChannel": case["lastChannel"], "source": case["source"],
        "openedAt": case["openedAt"], "updatedAt": case["updatedAt"], "closedAt": case["closedAt"],
        "detail": case["detail"], "photoUrl": case["photoUrl"],
        "complainantName": case["complainantName"], "phone": case["phone"],
    })


@app.post("/api/complaints")
def api_submit_complaint():
    data = request.get_json(force=True, silent=True) or {}
    parcel_no = (data.get("parcelNo") or "").strip()
    name = (data.get("name") or "").strip()
    phone = (data.get("phone") or "").strip()
    if not parcel_no:
        return jsonify(ok=False, error="ต้องระบุเลขพัสดุ — หากไม่ทราบเลขพัสดุ กรุณาสอบถามเจ้าหน้าที่ก่อน"), 400
    if not name or not phone:
        return jsonify(ok=False, error="กรุณากรอกชื่อและที่อยู่ผู้แจ้ง และเบอร์โทรติดต่อกลับ"), 400

    row_i, existing = find_case(parcel_no)
    now = now_str()
    if existing:
        sheet = ws(SHEET_CASES)
        if not existing["complainantName"]:
            sheet.update_cell(row_i, 5, name)
        if not existing["phone"]:
            sheet.update_cell(row_i, 6, phone)
        append_log(existing["ticketNo"], "อื่นๆ", existing["status"],
                   "ผู้รับแจ้งเรื่องเพิ่มเติม: " + (data.get("detail") or "-"),
                   "ผู้รับ (แจ้งเรื่องเอง)", "-")
        return jsonify(ok=True, ticketNo=existing["ticketNo"], matched=True)

    ticket_no = generate_ticket_no()
    ws(SHEET_CASES).append_row([
        ticket_no, parcel_no, data.get("company") or "", "ผู้รับแจ้งเรื่อง", name, phone,
        data.get("detail") or "", "", now, "รอติดตาม", "", "", now, "",
    ], value_input_option="USER_ENTERED")
    append_log(ticket_no, "อื่นๆ", "รอติดตาม",
               "เปิดเรื่องโดยผู้รับผ่านฟอร์มแจ้งเรื่อง: " + (data.get("detail") or "-"),
               "ผู้รับ (แจ้งเรื่องเอง)", "-")
    return jsonify(ok=True, ticketNo=ticket_no, matched=False)


# ===================== Routes: Auth =====================

@app.post("/api/login")
def api_login():
    data = request.get_json(force=True, silent=True) or {}
    username = (data.get("username") or "").strip()
    password = data.get("password") or ""
    if not username or not password:
        return jsonify(ok=False, error="กรุณากรอกชื่อผู้ใช้และรหัสผ่าน"), 400
    user = find_user(username)
    # ข้อความ error เหมือนกันทั้งกรณีไม่พบผู้ใช้/รหัสผ่านผิด กันการเดาว่าชื่อผู้ใช้ไหนมีอยู่จริง
    if not user or hash_password(password, user["salt"]) != user["hash"]:
        return jsonify(ok=False, error="ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง"), 401
    session.clear()
    session.permanent = True
    session["username"] = user["username"]
    session["nickname"] = user["nickname"]
    session["role"] = user["role"]
    return jsonify(ok=True, username=user["username"], nickname=user["nickname"], role=user["role"])


@app.post("/api/logout")
def api_logout():
    session.clear()
    return jsonify(ok=True)


@app.get("/api/me")
def api_me():
    if not session.get("username"):
        return jsonify(ok=False), 401
    return jsonify(ok=True, username=session["username"], nickname=session["nickname"], role=session["role"])


# ===================== Routes: Staff =====================

@app.get("/api/cases")
@staff_required
def api_list_cases():
    return jsonify(ok=True, cases=get_all_cases())


@app.get("/api/cases/<ticket_no>")
@staff_required
def api_case_detail(ticket_no):
    _, case = find_case(ticket_no)
    if not case:
        return jsonify(ok=False, error="ไม่พบเคสนี้ในระบบ"), 404
    return jsonify(ok=True, case=case, history=get_case_history(case["ticketNo"]))


@app.patch("/api/cases/<ticket_no>")
@staff_required
def api_update_status(ticket_no):
    data = request.get_json(force=True, silent=True) or {}
    status = data.get("status")
    if status not in STATUSES:
        return jsonify(ok=False, error="สถานะไม่ถูกต้อง"), 400
    row_i, case = find_case(ticket_no)
    if not case:
        return jsonify(ok=False, error="ไม่พบเคสนี้ในระบบ"), 404

    sheet = ws(SHEET_CASES)
    now = now_str()
    nickname = session["nickname"]
    # ช่องทางล่าสุด — คงค่าเดิมไว้ถ้าไม่ได้ส่งมาใหม่ (หน้าเว็บปัจจุบันไม่มี UI เลือกช่องทางตอนเปลี่ยนสถานะ)
    channel_val = data.get("channel") or case.get("lastChannel") or ""
    closed_val = now if status == "ส่งสาขา/ปิดเรื่อง" else (case.get("closedAt") or "")

    # อัปเดตคอลัมน์ J:N (สถานะปัจจุบัน, ช่องทางล่าสุด, เจ้าหน้าที่ล่าสุด, วันที่อัปเดตล่าสุด, วันที่ปิดเรื่อง)
    # ในคำขอ API ครั้งเดียว (แต่ก่อนยิงทีละคอลัมน์แยกกันสูงสุด 5 ครั้ง ทำให้อัปเดตสถานะช้า)
    sheet.update(f"J{row_i}:N{row_i}", [[status, channel_val, nickname, now, closed_val]],
                 value_input_option="USER_ENTERED")

    append_log(case["ticketNo"], data.get("channel") or "", status, data.get("note") or "",
               session["username"], nickname)

    updated_case = dict(case)
    updated_case.update({
        "status": status, "lastChannel": channel_val, "lastStaff": nickname,
        "updatedAt": now, "closedAt": closed_val,
    })
    # ส่งเคสที่อัปเดตแล้วกลับไปด้วย เพื่อให้หน้าเว็บอัปเดตข้อมูลในตัวเองได้เลย
    # ไม่ต้องยิง /api/cases โหลดทั้งชีตใหม่ทุกครั้ง (อีกจุดที่ทำให้ก่อนหน้านี้ช้า)
    return jsonify(ok=True, ticketNo=case["ticketNo"], case=updated_case)


@app.post("/api/cases/<ticket_no>/edit")
@staff_required
def api_edit_case_info(ticket_no):
    # แก้ไข "รายละเอียด" และ "เบอร์ติดต่อ" ของเคส — พนักงานทุกคนแก้ไขได้ (@staff_required อนุญาตทั้ง
    # staff และ admin) ฟิลด์อื่นของเคสยังคงแก้ไขไม่ได้จากหน้านี้เหมือนเดิม
    data = request.get_json(force=True, silent=True) or {}
    phone = (data.get("phone") or "").strip()
    detail = (data.get("detail") or "").strip()
    if not phone:
        return jsonify(ok=False, error="กรุณากรอกเบอร์ติดต่อ"), 400
    if not detail:
        return jsonify(ok=False, error="กรุณากรอกรายละเอียด"), 400

    row_i, case = find_case(ticket_no)
    if not case:
        return jsonify(ok=False, error="ไม่พบเคสนี้ในระบบ"), 404

    sheet = ws(SHEET_CASES)
    now = now_str()
    nickname = session["nickname"]
    # คอลัมน์ F = เบอร์ติดต่อ (phone), G = รายละเอียด (detail)
    sheet.update(f"F{row_i}:G{row_i}", [[phone, detail]], value_input_option="USER_ENTERED")
    # อัปเดต "พนักงานดูแล" (lastStaff) เป็นชื่อพนักงานที่แก้ไขข้อมูลครั้งนี้ ให้สอดคล้องกับฟีเจอร์แก้ไขรูปภาพ
    sheet.update(f"L{row_i}:M{row_i}", [[nickname, now]], value_input_option="USER_ENTERED")

    append_log(case["ticketNo"], case.get("lastChannel") or "", case["status"], "แก้ไขเบอร์ติดต่อ/รายละเอียด",
               session["username"], nickname)

    updated_case = dict(case)
    updated_case.update({"phone": phone, "detail": detail, "lastStaff": nickname, "updatedAt": now})
    return jsonify(ok=True, ticketNo=case["ticketNo"], case=updated_case)


# ประเภทไฟล์รูปภาพที่รับอัปโหลด และขนาดสูงสุด (ไบต์) — คู่กับ app.config["MAX_CONTENT_LENGTH"] ด้านบน
ALLOWED_PHOTO_TYPES = {"image/jpeg": "jpg", "image/png": "png", "image/webp": "webp", "image/gif": "gif"}
MAX_PHOTO_BYTES = 8 * 1024 * 1024


@app.post("/api/cases/<ticket_no>/photo")
@staff_required
def api_upload_case_photo(ticket_no):
    # เพิ่ม/แก้ไขรูปภาพประกอบเคส — พนักงานทุกคน (ไม่ใช่แค่แอดมิน) มีสิทธิ์ทำได้ (@staff_required อนุญาตทั้ง
    # staff และ admin) หน้าเว็บจะให้ยืนยันก่อนอัปโหลดจริงอยู่แล้ว ฝั่ง backend นี้จึงไม่มีขั้นตอนยืนยันซ้ำอีก
    if not (DRIVE_UPLOAD_FOLDER_ID and DRIVE_OAUTH_CLIENT_ID and DRIVE_OAUTH_CLIENT_SECRET and DRIVE_OAUTH_REFRESH_TOKEN):
        return jsonify(ok=False, error="ระบบยังไม่ได้ตั้งค่าอัปโหลดรูปภาพ Google Drive ครบถ้วน "
                                        "กรุณาติดต่อผู้ดูแลระบบ"), 500

    row_i, case = find_case(ticket_no)
    if not case:
        return jsonify(ok=False, error="ไม่พบเคสนี้ในระบบ"), 404

    file = request.files.get("photo")
    if not file or not file.filename:
        return jsonify(ok=False, error="กรุณาเลือกไฟล์รูปภาพ"), 400

    mimetype = file.mimetype or ""
    if mimetype not in ALLOWED_PHOTO_TYPES:
        return jsonify(ok=False, error="รองรับเฉพาะไฟล์รูปภาพ (JPG, PNG, WEBP, GIF)"), 400

    file_bytes = file.read()
    if len(file_bytes) > MAX_PHOTO_BYTES:
        return jsonify(ok=False, error="ไฟล์รูปภาพใหญ่เกินไป (สูงสุดไม่เกิน 8MB)"), 400
    if not file_bytes:
        return jsonify(ok=False, error="ไฟล์รูปภาพว่างเปล่า"), 400

    try:
        # import ตรงนี้ (lazy) แทนบนสุดของไฟล์ เพื่อไม่ให้การสร้าง Drive client ไปทำ network call
        # ตอนแอปสตาร์ทขึ้น (จะได้ไม่เพิ่มความเสี่ยงต่อ retry/startup logic ของการเชื่อมต่อ Sheets ด้านบน)
        from google.oauth2.credentials import Credentials as UserCredentials
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseUpload
        # อัปโหลดด้วยสิทธิ์ของบัญชี Google จริงที่เป็นเจ้าของโฟลเดอร์ (ผ่าน OAuth refresh token) แทน
        # Service Account เพราะ Service Account ไม่มี storage quota เป็นของตัวเอง สร้างไฟล์ใหม่ไม่ได้
        # แม้จะมีสิทธิ์ Editor ในโฟลเดอร์ที่แชร์ให้ก็ตาม (ข้อจำกัดจริงของ Google Drive API)
        user_creds = UserCredentials(
            token=None,
            refresh_token=DRIVE_OAUTH_REFRESH_TOKEN,
            client_id=DRIVE_OAUTH_CLIENT_ID,
            client_secret=DRIVE_OAUTH_CLIENT_SECRET,
            token_uri="https://oauth2.googleapis.com/token",
            scopes=["https://www.googleapis.com/auth/drive"],
        )
        drive = build("drive", "v3", credentials=user_creds, cache_discovery=False)
        ext = ALLOWED_PHOTO_TYPES[mimetype]
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mimetype, resumable=False)
        file_meta = {
            "name": f"{case['ticketNo']}_{secrets.token_hex(4)}.{ext}",
            "parents": [DRIVE_UPLOAD_FOLDER_ID],
        }
        created = drive.files().create(body=file_meta, media_body=media, fields="id").execute()
        file_id = created["id"]
        # แชร์เป็นสาธารณะแบบ "ใครมีลิงก์ก็ดูได้" (เหมือนรูปภาพเดิมที่ sync มาจาก AppSheet) เพื่อให้
        # driveThumbUrl()/handlePhotoError() ฝั่งหน้าเว็บ (ที่มี fallback URL หลายแบบ) แสดงผลได้ปกติ
        drive.permissions().create(fileId=file_id, body={"type": "anyone", "role": "reader"}).execute()
    except Exception as e:  # noqa: BLE001 — ครอบคลุม error จาก Drive API ทุกแบบ (สิทธิ์ไม่พอ/API ยังไม่เปิด/ฯลฯ)
        return jsonify(ok=False, error="อัปโหลดรูปภาพไม่สำเร็จ: " + str(e)), 502

    photo_url = f"https://drive.google.com/uc?export=view&id={file_id}"
    now = now_str()
    nickname = session["nickname"]
    had_photo_before = bool((case.get("photoUrl") or "").strip())

    sheet = ws(SHEET_CASES)
    sheet.update(f"H{row_i}", [[photo_url]], value_input_option="USER_ENTERED")
    # อัปเดต "พนักงานดูแล" (lastStaff) เป็นชื่อพนักงานที่ดำเนินการอัปโหลด/แก้ไขรูปภาพครั้งนี้ ตามที่ขอ
    sheet.update(f"L{row_i}:M{row_i}", [[nickname, now]], value_input_option="USER_ENTERED")

    note = "แก้ไขรูปภาพประกอบเคส" if had_photo_before else "เพิ่มรูปภาพประกอบเคส"
    append_log(case["ticketNo"], case.get("lastChannel") or "", case["status"], note,
               session["username"], nickname)

    updated_case = dict(case)
    updated_case.update({"photoUrl": photo_url, "lastStaff": nickname, "updatedAt": now})
    return jsonify(ok=True, ticketNo=case["ticketNo"], case=updated_case)


@app.post("/api/export.xlsx")
@staff_required
def api_export_excel():
    # รับรายการเลขแจ้งเรื่อง (tickets) ที่หน้าเว็บกรอง/แสดงอยู่ตรงๆ มาทาง POST body (ฟอร์ม ไม่ใช่ query string)
    # แทนที่จะให้ backend คำนวณเงื่อนไขกรองซ้ำ — รับประกันว่าไฟล์ที่ออกมาตรงกับที่เห็นบนหน้าจอ 100%
    # (ไม่ว่าจะเป็นหน้าแดชบอร์ดหรือหน้าเคสแจ้งเรื่อง ซึ่งมีเงื่อนไขกรองต่างกัน)
    # เดิมเคยลองใช้ GET + query string (tickets คั่นด้วย comma) แต่พอเคสมีจำนวนมาก URL ยาวเกินไปจน
    # เซิร์ฟเวอร์/พร็อกซีปฏิเสธคำขอ (Request Line too large) จึงเปลี่ยนมาส่งทาง POST body แทน
    # ฝั่งหน้าเว็บใช้การ submit <form method="POST"> จริง (ไม่ใช่ fetch+Blob) เพื่อให้เบราว์เซอร์ดาวน์โหลด
    # ไฟล์ให้เองแบบเดียวกับการเปิดลิงก์ตรงๆ — ไม่มีข้อจำกัดเรื่องความยาว URL และยังดาวน์โหลดได้ชัวร์เหมือนเดิม
    tickets_raw = request.form.get("tickets") or request.args.get("tickets") or ""
    tickets = [t.strip() for t in tickets_raw.split(",") if t.strip()]
    date_from = request.form.get("dateFrom") or request.args.get("dateFrom") or ""
    date_to = request.form.get("dateTo") or request.args.get("dateTo") or ""

    if not tickets:
        return jsonify(ok=False, error="ไม่มีรายการเคสให้ออกรายงาน"), 400

    by_ticket = {c["ticketNo"]: c for c in get_all_cases()}
    rows = [by_ticket[t] for t in tickets if t in by_ticket]

    wb = Workbook()
    wsx = wb.active
    wsx.title = "รายละเอียดเคส"
    header = ["เลขแจ้งเรื่อง", "เลขพัสดุ", "บริษัทขนส่ง", "แหล่งที่มา", "ชื่อและที่อยู่ผู้แจ้ง", "เบอร์โทร",
              "รายละเอียด", "รูปภาพ", "วันที่เปิดเรื่อง", "สถานะปัจจุบัน", "ช่องทางล่าสุด",
              "เจ้าหน้าที่ล่าสุด", "วันที่อัปเดตล่าสุด", "วันที่ปิดเรื่อง"]
    wsx.append(header)
    for c in rows:
        wsx.append([c[k] for k in CASE_COLS])

    summary = wb.create_sheet("สรุปภาพรวม")
    by_status = {s: sum(1 for c in rows if c["status"] == s) for s in STATUSES}
    summary.append(["ภาพรวม", f"{date_from or 'เริ่มต้น'} ถึง {date_to or 'ปัจจุบัน'}"])
    summary.append(["รวมทั้งหมด", len(rows)])
    summary.append([])
    summary.append(["แยกตามสถานะ"])
    for s, n in by_status.items():
        summary.append([s, n])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"รายงานพัสดุ_{date_from or 'เริ่มต้น'}_ถึง_{date_to or 'ปัจจุบัน'}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                      mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ===================== Routes: Admin (users) =====================

@app.get("/api/users")
@admin_required
def api_list_users():
    rows = ws(SHEET_USERS).get_all_values()
    out = []
    for row in rows[1:]:
        if row and row[0]:
            out.append({"username": row[0].strip(), "nickname": row[1], "role": row[2]})
    return jsonify(ok=True, users=out)


@app.post("/api/users")
@admin_required
def api_add_user():
    data = request.get_json(force=True, silent=True) or {}
    username = (data.get("username") or "").strip()
    nickname = data.get("nickname")
    role = data.get("role")
    password = data.get("password") or ""
    if not username or not nickname or not password:
        return jsonify(ok=False, error="กรอกข้อมูลไม่ครบ"), 400
    if len(password) < 6:
        return jsonify(ok=False, error="รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร"), 400
    if role not in ("staff", "admin"):
        return jsonify(ok=False, error="บทบาทไม่ถูกต้อง"), 400
    if find_user(username):
        return jsonify(ok=False, error="มีชื่อผู้ใช้นี้ในระบบอยู่แล้ว"), 400
    salt = secrets.token_hex(16)
    ws(SHEET_USERS).append_row([username, nickname, role, salt, hash_password(password, salt)],
                                value_input_option="USER_ENTERED")
    return jsonify(ok=True)


@app.delete("/api/users/<username>")
@admin_required
def api_remove_user(username):
    user = find_user(username)
    if not user:
        return jsonify(ok=False, error="ไม่พบชื่อผู้ใช้นี้"), 404
    ws(SHEET_USERS).delete_rows(user["row"])
    return jsonify(ok=True)


@app.patch("/api/users/<username>/password")
@admin_required
def api_reset_password(username):
    data = request.get_json(force=True, silent=True) or {}
    new_password = data.get("newPassword") or ""
    user = find_user(username)
    if not user:
        return jsonify(ok=False, error="ไม่พบชื่อผู้ใช้นี้"), 404
    if len(new_password) < 6:
        return jsonify(ok=False, error="รหัสผ่านใหม่ต้องมีอย่างน้อย 6 ตัวอักษร"), 400
    salt = secrets.token_hex(16)
    sheet = ws(SHEET_USERS)
    sheet.update_cell(user["row"], 4, salt)
    sheet.update_cell(user["row"], 5, hash_password(new_password, salt))
    return jsonify(ok=True)


if __name__ == "__main__":
    app.run(debug=os.environ.get("FLASK_DEBUG") == "1", port=int(os.environ.get("PORT", 5000)))
