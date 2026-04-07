
from __future__ import annotations

import os
import io
import re
import time
import secrets
from datetime import datetime, date, timedelta
from types import SimpleNamespace

import pandas as pd
from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, send_file, abort, send_from_directory, jsonify
)
from flask_login import (
    LoginManager, login_user, logout_user,
    login_required, current_user
)
from flask_wtf.csrf import CSRFProtect
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from sqlalchemy import func, text, event
from sqlalchemy.exc import OperationalError

from models import db, User, Issue, FollowUp, Log, Guideline

try:
    import win32com.client as win32
    import pythoncom
except Exception:
    win32 = None
    pythoncom = None


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INSTANCE_DIR = os.path.join(BASE_DIR, "instance")
os.makedirs(INSTANCE_DIR, exist_ok=True)

UPLOAD_FOLDER = os.path.join(BASE_DIR, "static", "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

TEMPLATE_FILES_DIR = os.path.join(BASE_DIR, "static", "templates")
os.makedirs(TEMPLATE_FILES_DIR, exist_ok=True)

TEMPLATE_XLSX = "zovlomj_heregjilt_template.xlsx"

ALLOWED_EXTENSIONS = {"pdf", "doc", "docx", "xls", "xlsx", "png", "jpg", "jpeg"}
MAX_UPLOAD_MB = 20


def create_app() -> Flask:
    app = Flask(__name__, template_folder=os.path.join(BASE_DIR, "templates"))
    app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_urlsafe(32)

    app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///" + os.path.join(INSTANCE_DIR, "database.db")
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
    app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
        "pool_pre_ping": True,
        "connect_args": {
            "check_same_thread": False,
            "timeout": 10,
        },
    }

    app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
    app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
    app.config["SESSION_COOKIE_HTTPONLY"] = True
    app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
    app.config["SESSION_COOKIE_SECURE"] = (os.environ.get("SESSION_COOKIE_SECURE", "0") == "1")
    return app

app = create_app()
db.init_app(app)
csrf = CSRFProtect(app)

app.jinja_env.globals["now"] = datetime.now
app.jinja_env.globals["getattr"] = getattr
app.jinja_env.globals["timedelta"] = timedelta

login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message = "Системд нэвтэрсэн байна."
login_manager.login_message_category = "info"

with app.app_context():
    @event.listens_for(db.engine, "connect")
    def _set_sqlite_pragma(dbapi_connection, connection_record):
        try:
            cursor = dbapi_connection.cursor()
            cursor.execute("PRAGMA foreign_keys=ON;")
            cursor.execute("PRAGMA journal_mode=WAL;")
            cursor.execute("PRAGMA synchronous=NORMAL;")
            cursor.execute("PRAGMA busy_timeout=10000;")
            cursor.close()
        except Exception:
            pass

COMPANIES = [
    "ЭРДЭНЭС МОНГОЛ ХХК", "ЭРДЭНЭТ ҮЙЛДВЭР ТӨҮГ", "ЭРДМИН ХХК", "ЭРДЭНЭТ МЕДИКАЛ ХХК",
    "ШИМ ТЕХНОЛОДЖИ ХХК", "ЭРДЭНЭТ ОЙЛ ХХК", "МОНГОЛИАН СМЕЛТИНГ КОРПОРЕЙШН ХХК",
    "ДАРХАНЫ ТӨМӨРЛӨГИЙН ҮЙЛДВЭР ТӨХХК", "ДАРХАН МАЙНИНГ ХХК", "ЦЕМЕНТ ШОХОЙ ТӨХК",
    "ХӨТӨЛ ЭНЕРЖИ ДУЛААН ХХК", "МОНГОЛЫН ГАЗРЫН ТОС БОЛОВСРУУЛАХ ҮЙЛДВЭР ХХК",
    "МОН-АТОМ ХХК", "БАДРАХ ЭНЕРЖИ ХХК", "МОН ЧЕХ УРАНИУМ ХХК", "ГУРВАН САЙХАН ХХК",
    "ТӨВ АЗИЙН УРАН ХХК", "ЭРДЭНЭС МЕТАН ХХК", "МЕТАН ГАЗ РЕСУРС ХХК",
    "ЭРДЭНЭС БАЯНБОГД ХХК", "ЭРДЭНЭС ОЮУ ТОЛГОЙ ХХК",
    "ЭРДЭНЭС КРИТИКАЛ МИНЕРАЛС ТӨҮГ", "ЭРДЭНЭС АЛТ РЕСУРС ХХК",
    "ЭРДЭНЭС ТАВАН ТОЛГОЙ ХК", "ТАВАН ТОЛГОЙ ТҮЛШ ХХК",
    "ТАВАН ТОЛГОЙ ТӨМӨР ЗАМ ХХК", "ТАВАН ТОЛГОЙ ДЦС ХХК",
    "ГАШУУН СУХАЙТ АВТОЗАМ ХХК", "БАГАНУУР ХК", "БАГАНУУР СУВИЛАЛ ХХК",
    "БАГАНУУР ИЛЧ ХХК", "МОНЦАХИМ ХХК", "ШИВЭЭ ОВОО ХК", "ШИВЭЭ СЕРВИС ХХК",
    "ЭРДЭНЭС ҮТП ХХК", "ЭРДЭНЭС ГАЗ", "ЭРДЭНЭС АШИД ХХК",
    "ЧИНГИС ХААН ҮНДЭСНИЙ БАЯЛАГИЙН САН ХХК", "ЗЭС БОЛОВСРУУЛАХ ҮЙЛДВЭР"
]

AUDITORS = [
    "Ахлах аудитор З.Уянга", "Аудитор Б.Мөнгөнцэцэг", "Аудитор Б.Уранзаяа",
    "Аудитор Н.Мөнхжаргал", "Аудитор А.Бямбадорж", "Аудитор Ж.Хашбат",
    "Аудитор Ж.Баярсайхан", "Аудитор Д.Чулуунбат", "Аудитор Сү.Гэрэлмаа",
    "Аудитор Д.Эрдэнэ-Очир", "Аудитор Ц.Ядамцоо",
    "Ахлах аудитор Б.Батжаргалсайхан", "Аудитор Г.Батхонгор",
    "Аудитор Со.Гэрэлмаа", "Менежер Я.Сарансүх",
    "Ерөнхий аудитор О.Баярмагнай", "Аудитор Б.Чанцалмаа",
    "Аудитор Т.Бурмаа", "Аудитор Б.Ундрал", "Аудитор А.Ариунтунгалаг"
]

COMPANY_AIMAG = {
    "ЭРДЭНЭС МОНГОЛ ХХК": "Ulaanbaatar",
    "ЭРДЭНЭТ ҮЙЛДВЭР ТӨҮГ": "Erdenet",
    "ЭРДМИН ХХК": "Erdenet",
    "ЭРДЭНЭТ МЕДИКАЛ ХХК": "Erdenet",
    "ШИМ ТЕХНОЛОДЖИ ХХК": "Erdenet",
    "ЭРДЭНЭТ ОЙЛ ХХК": "Erdenet",
    "МОНГОЛИАН СМЕЛТИНГ КОРПОРЕЙШН ХХК": "Erdenet",
    "ДАРХАНЫ ТӨМӨРЛӨГИЙН ҮЙЛДВЭР ТӨХХК": "Darkhan-Uul",
    "ДАРХАН МАЙНИНГ ХХК": "Darkhan-Uul",
    "ЦЕМЕНТ ШОХОЙ ТӨХК": "Selenge",
    "ХӨТӨЛ ЭНЕРЖИ ДУЛААН ХХК": "Selenge",
    "МОНГОЛЫН ГАЗРЫН ТОС БОЛОВСРУУЛАХ ҮЙЛДВЭР ХХК": "Dornogovi",
    "МОН-АТОМ ХХК": "Ulaanbaatar",
    "БАДРАХ ЭНЕРЖИ ХХК": "Dornogovi",
    "МОН ЧЕХ УРАНИУМ ХХК": "Ulaanbaatar",
    "ГУРВАН САЙХАН ХХК": "Dundgovi",
    "ТӨВ АЗИЙН УРАН ХХК": "Dornod",
    "ЭРДЭНЭС МЕТАН ХХК": "Omnogovi",
    "МЕТАН ГАЗ РЕСУРС ХХК": "Omnogovi",
    "ЭРДЭНЭС БАЯНБОГД ХХК": "Dornogovi",
    "ЭРДЭНЭС ОЮУ ТОЛГОЙ ХХК": "Ulaanbaatar",
    "ЭРДЭНЭС КРИТИКАЛ МИНЕРАЛС ТӨҮГ": "Khentii",
    "ЭРДЭНЭС АЛТ РЕСУРС ХХК": "Dundgovi",
    "ЭРДЭНЭС ТАВАН ТОЛГОЙ ХК": "Omnogovi",
    "ТАВАН ТОЛГОЙ ТҮЛШ ХХК": "Ulaanbaatar",
    "ТАВАН ТОЛГОЙ ТӨМӨР ЗАМ ХХК": "Omnogovi",
    "ТАВАН ТОЛГОЙ ДЦС ХХК": "Omnogovi",
    "ГАШУУН СУХАЙТ АВТОЗАМ ХХК": "Omnogovi",
    "БАГАНУУР ХК": "Baganuur",
    "БАГАНУУР СУВИЛАЛ ХХК": "Baganuur",
    "БАГАНУУР ИЛЧ ХХК": "Baganuur",
    "МОНЦАХИМ ХХК": "Baganuur",
    "ШИВЭЭ ОВОО ХК": "Govisumber",
    "ШИВЭЭ СЕРВИС ХХК": "Govisumber",
    "ЭРДЭНЭС ҮТП ХХК": "Ulaanbaatar",
    "ЭРДЭНЭС ГАЗ": "Ulaanbaatar",
    "ЭРДЭНЭС АШИД ХХК": "Ulaanbaatar",
    "ЧИНГИС ХААН ҮНДЭСНИЙ БАЯЛАГИЙН САН ХХК": "Ulaanbaatar",
    "ЗЭС БОЛОВСРУУЛАХ ҮЙЛДВЭР": "Ulaanbaatar"
}

def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def parse_date(s: str | None) -> date | None:
    s = (s or "").strip()
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except Exception:
        return None

def parse_money_to_int(val) -> int:
    if val is None:
        return 0
    s = str(val).strip()
    if not s:
        return 0

    nums = re.findall(r"\d+", s)
    if not nums:
        return 0

    joined = "".join(nums)
    if len(nums) >= 2 and nums[-1] == "00":
        joined = joined[:-2]

    try:
        return int(joined) if joined else 0
    except Exception:
        return 0

def is_yes_mn(v: str | None) -> bool:
    s = (v or "").strip().lower()
    return s in {"тийм", "tiim", "yes", "y", "true", "1", "on"}

def apply_issue_filters(q, selected_auditor="", selected_company="", df=None, dt=None):
    if selected_auditor:
        q = q.filter(Issue.auditor_name == selected_auditor)
    if selected_company:
        q = q.filter(Issue.company_name == selected_company)
    if df:
        q = q.filter(Issue.identified_date >= df)
    if dt:
        q = q.filter(Issue.identified_date <= dt)
    return q

def apply_followup_filters(q, selected_auditor="", selected_company="", df=None, dt=None):
    if selected_auditor:
        q = q.filter(FollowUp.auditor_name == selected_auditor)
    if selected_company:
        q = q.filter(Issue.company_name == selected_company)
    if df:
        q = q.filter(Issue.identified_date >= df)
    if dt:
        q = q.filter(Issue.identified_date <= dt)
    return q

def build_stacked_matrix_from_rows(rows, labels_order=None, levels_order=None):
    levels_order = levels_order or ["Өндөр", "Дунд", "Бага"]

    grouped = {}
    for name, level, count in rows:
        key = (name or "—").strip() or "—"
        lvl = (level or "").strip()
        cnt = int(count or 0)

        if key not in grouped:
            grouped[key] = {lvl_name: 0 for lvl_name in levels_order}

        if lvl in grouped[key]:
            grouped[key][lvl] += cnt

    if labels_order is None:
        labels = list(grouped.keys())
    else:
        labels = [x for x in labels_order if x in grouped]

    matrix = []
    for lvl in levels_order:
        matrix.append([grouped.get(label, {}).get(lvl, 0) for label in labels])

    return labels, levels_order, matrix

def admin_required():
    if (not current_user.is_authenticated) or (not getattr(current_user, "is_admin", False)):
        abort(403)

def user_is_blocked(u: User) -> bool:
    return bool(getattr(u, "is_blocked", False))

def can_login(u: User) -> bool:
    return (u is not None) and (not user_is_blocked(u))

def commit_with_retry(max_retries: int = 5, sleep_seconds: float = 0.25):
    for attempt in range(max_retries):
        try:
            db.session.commit()
            return True
        except OperationalError as e:
            db.session.rollback()
            msg = str(e).lower()
            if "database is locked" in msg and attempt < max_retries - 1:
                time.sleep(sleep_seconds * (attempt + 1))
                continue
            raise
        except Exception:
            db.session.rollback()
            raise
    return False

def log_action_mn(action_mn: str, auto_commit: bool = True):
    try:
        ip = request.headers.get("X-Forwarded-For", request.remote_addr) or ""
        path = request.path
        method = request.method
        who = current_user.username if current_user.is_authenticated else "anonymous"
        msg = f"{action_mn} | IP={ip} | {method} {path}"
        db.session.add(Log(user=who, action=msg))
        if auto_commit:
            commit_with_retry()
        else:
            db.session.flush()
    except Exception:
        db.session.rollback()

def is_guideline_member(g) -> bool:
    if current_user.is_authenticated and current_user.is_admin:
        return True
    u = (getattr(current_user, "username", "") or "").strip()
    leader = (getattr(g, "team_leader", "") or "").strip()
    members = [x.strip() for x in (getattr(g, "team_members", "") or "").split(";") if x.strip()]
    return bool(u) and (u == leader or u in members)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

def _sqlite_table_exists(table_name: str) -> bool:
    try:
        r = db.session.execute(
            text("SELECT name FROM sqlite_master WHERE type='table' AND name=:t"),
            {"t": table_name}
        ).fetchone()
        return bool(r)
    except Exception:
        return False

def _table_count(table_name: str) -> int:
    try:
        return int(db.session.execute(text(f"SELECT COUNT(*) FROM {table_name}")).scalar() or 0)
    except Exception:
        return 0

def _to_date(v):
    if v is None:
        return None
    if isinstance(v, date) and not isinstance(v, datetime):
        return v
    if isinstance(v, datetime):
        return v.date()
    try:
        s = str(v).strip()
        if not s:
            return None
        s = s.split(" ")[0]
        return datetime.strptime(s, "%Y-%m-%d").date()
    except Exception:
        return None

def _to_datetime(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v
    try:
        s = str(v).strip()
        if not s:
            return None

        for fmt in (
            "%Y-%m-%d %H:%M:%S.%f",
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d %H:%M",
            "%Y-%m-%d",
        ):
            try:
                return datetime.strptime(s, fmt)
            except Exception:
                pass
        return None
    except Exception:
        return None

def _guideline_existing_tables() -> list[str]:
    tables = []
    for t in ("guideline", "guidelines"):
        if _sqlite_table_exists(t):
            tables.append(t)
    return tables

def _normalize_guideline_row(r: dict, source_table: str):
    obj = SimpleNamespace(**dict(r))
    obj._source_table = source_table
    obj.scope_start = _to_date(getattr(obj, "scope_start", None))
    obj.scope_end = _to_date(getattr(obj, "scope_end", None))
    obj.exec_start = _to_date(getattr(obj, "exec_start", None))
    obj.exec_end = _to_date(getattr(obj, "exec_end", None))
    obj.extended_end = _to_date(getattr(obj, "extended_end", None))
    obj.created_at = _to_datetime(getattr(obj, "created_at", None))
    return obj

def _guideline_table_name() -> str | None:
    candidates = []
    for t in _guideline_existing_tables():
        candidates.append((t, _table_count(t)))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[1], reverse=True)
    return candidates[0][0]

def _guideline_write_table_name() -> str:
    read_table = _guideline_table_name()
    if read_table:
        return read_table

    if _sqlite_table_exists("guideline"):
        return "guideline"
    if _sqlite_table_exists("guidelines"):
        return "guidelines"
    return "guideline"

def _fetch_all_guidelines_for_list():
    tables = _guideline_existing_tables()
    if not tables:
        return []

    out = []
    seen = set()

    for table in tables:
        try:
            rows = db.session.execute(
                text(f"SELECT * FROM {table} ORDER BY id DESC")
            ).mappings().all()

            for r in rows:
                key = (table, r.get("id"))
                if key in seen:
                    continue
                seen.add(key)
                out.append(_normalize_guideline_row(dict(r), table))
        except Exception:
            continue

    out.sort(
        key=lambda x: (
            getattr(x, "created_at", None) or datetime.min,
            getattr(x, "id", 0)
        ),
        reverse=True
    )
    return out

def _fetch_one_guideline_raw(gid: int):
    for table in _guideline_existing_tables():
        try:
            row = db.session.execute(
                text(f"SELECT * FROM {table} WHERE id=:id"),
                {"id": gid}
            ).mappings().first()

            if row:
                return _normalize_guideline_row(dict(row), table)
        except Exception:
            continue
    return None

def _insert_guideline_raw(
    company_name: str,
    audit_type: str,
    audit_subtype: str,
    team_leader: str,
    team_members: str,
    scope_start,
    scope_end,
    exec_start,
    exec_end,
    extended_end,
    extension_note: str,
    created_by: str
) -> int:
    table = _guideline_write_table_name()

    db.session.execute(
        text(f"""
            INSERT INTO {table}
            (
                company_name, audit_type, audit_subtype,
                team_leader, team_members,
                scope_start, scope_end,
                exec_start, exec_end,
                extended_end, extension_note,
                created_by, created_at
            )
            VALUES
            (
                :company_name, :audit_type, :audit_subtype,
                :team_leader, :team_members,
                :scope_start, :scope_end,
                :exec_start, :exec_end,
                :extended_end, :extension_note,
                :created_by, :created_at
            )
        """),
        {
            "company_name": company_name,
            "audit_type": audit_type,
            "audit_subtype": audit_subtype,
            "team_leader": team_leader,
            "team_members": team_members,
            "scope_start": scope_start,
            "scope_end": scope_end,
            "exec_start": exec_start,
            "exec_end": exec_end,
            "extended_end": extended_end,
            "extension_note": extension_note,
            "created_by": created_by,
            "created_at": datetime.utcnow(),
        }
    )
    db.session.flush()

    new_id = db.session.execute(text("SELECT last_insert_rowid()")).scalar()
    return int(new_id)

def _update_guideline_raw(
    gid: int,
    company_name: str,
    audit_type: str,
    audit_subtype: str,
    team_leader: str,
    team_members: str,
    scope_start,
    scope_end,
    exec_start,
    exec_end,
    extended_end,
    extension_note: str
) -> bool:
    g = _fetch_one_guideline_raw(gid)
    if not g:
        return False

    table = getattr(g, "_source_table", None) or _guideline_table_name()
    if not table:
        return False

    db.session.execute(
        text(f"""
            UPDATE {table}
            SET
                company_name=:company_name,
                audit_type=:audit_type,
                audit_subtype=:audit_subtype,
                team_leader=:team_leader,
                team_members=:team_members,
                scope_start=:scope_start,
                scope_end=:scope_end,
                exec_start=:exec_start,
                exec_end=:exec_end,
                extended_end=:extended_end,
                extension_note=:extension_note
            WHERE id=:id
        """),
        {
            "company_name": company_name,
            "audit_type": audit_type,
            "audit_subtype": audit_subtype,
            "team_leader": team_leader,
            "team_members": team_members,
            "scope_start": scope_start,
            "scope_end": scope_end,
            "exec_start": exec_start,
            "exec_end": exec_end,
            "extended_end": extended_end,
            "extension_note": extension_note,
            "id": gid,
        }
    )
    db.session.flush()
    return True

def _delete_guideline_raw(gid: int) -> bool:
    g = _fetch_one_guideline_raw(gid)
    if not g:
        return False

    table = getattr(g, "_source_table", None) or _guideline_table_name()
    if not table:
        return False

    db.session.execute(
        text(f"DELETE FROM {table} WHERE id=:id"),
        {"id": gid}
    )
    db.session.flush()
    return True

def _table_cols(table: str) -> set[str]:
    cols: set[str] = set()
    try:
        rows = db.session.execute(text(f"PRAGMA table_info({table})")).fetchall()
        for r in rows:
            cols.add(str(r[1]))
    except Exception:
        pass
    return cols

def _add_col_if_missing(table: str, col: str, ddl: str):
    try:
        if not _sqlite_table_exists(table):
            return
        cols = _table_cols(table)
        if col not in cols:
            db.session.execute(text(f"ALTER TABLE {table} ADD COLUMN {ddl}"))
            commit_with_retry()
    except Exception:
        db.session.rollback()

def ensure_sqlite_schema():
    _add_col_if_missing("user", "is_blocked", "is_blocked BOOLEAN DEFAULT 0")
    _add_col_if_missing("user", "created_at", "created_at DATETIME")

    _add_col_if_missing("issue", "created_at", "created_at DATETIME")
    _add_col_if_missing("issue", "evidence_file", "evidence_file TEXT")
    _add_col_if_missing("issue", "money_amount", "money_amount TEXT")
    _add_col_if_missing("issue", "has_actual_loss", "has_actual_loss TEXT")

    _add_col_if_missing("follow_up", "evidence_file", "evidence_file TEXT")
    _add_col_if_missing("follow_up", "auditor_evidence_file", "auditor_evidence_file TEXT")
    _add_col_if_missing("follow_up", "reduced_amount", "reduced_amount TEXT")
    _add_col_if_missing("follow_up", "created_at", "created_at DATETIME")

    _add_col_if_missing("guideline", "created_at", "created_at DATETIME")
    _add_col_if_missing("guideline", "approved_pdf", "approved_pdf TEXT")

    _add_col_if_missing("guidelines", "created_at", "created_at DATETIME")
    _add_col_if_missing("guidelines", "approved_pdf", "approved_pdf TEXT")

def ensure_sqlite_indexes():
    stmts = [
        "CREATE INDEX IF NOT EXISTS idx_issue_company_name ON issue(company_name)",
        "CREATE INDEX IF NOT EXISTS idx_issue_auditor_name ON issue(auditor_name)",
        "CREATE INDEX IF NOT EXISTS idx_issue_identified_date ON issue(identified_date)",
        "CREATE INDEX IF NOT EXISTS idx_issue_residual_level ON issue(residual_level)",
        "CREATE INDEX IF NOT EXISTS idx_issue_company_auditor_date ON issue(company_name, auditor_name, identified_date)",
        "CREATE INDEX IF NOT EXISTS idx_followup_issue_id ON follow_up(issue_id)",
        "CREATE INDEX IF NOT EXISTS idx_followup_created_at ON follow_up(created_at)",
        "CREATE INDEX IF NOT EXISTS idx_log_user ON log(user)",
        "CREATE INDEX IF NOT EXISTS idx_log_timestamp ON log(timestamp)",
    ]

    for stmt in stmts:
        try:
            db.session.execute(text(stmt))
        except Exception:
            db.session.rollback()

    try:
        commit_with_retry()
    except Exception:
        db.session.rollback()

def export_excel(df: pd.DataFrame, filename: str, sheet_name: str = "Data"):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    bio.seek(0)
    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        email = (request.form.get("email") or "").strip()
        password = request.form.get("password") or ""

        if not username or not email or not password:
            flash("Бүх талбарыг бөглөнө үү.", "error")
            return redirect(url_for("register"))

        if User.query.filter_by(username=username).first():
            flash("Нэвтрэх нэр бүртгэлтэй байна.", "error")
            return redirect(url_for("register"))

        if User.query.filter_by(email=email).first():
            flash("Имэйл бүртгэлтэй байна.", "error")
            return redirect(url_for("register"))

        admin_username = os.environ.get("ADMIN_USERNAME", "dag_admin")
        is_admin = (username == admin_username)

        try:
            u = User(
                username=username,
                email=email,
                password_hash=generate_password_hash(password),
                is_admin=is_admin,
            )
            u.is_blocked = False

            db.session.add(u)
            db.session.flush()
            log_action_mn("Шинэ хэрэглэгч бүртгүүлэв", auto_commit=False)
            commit_with_retry()

            flash("Амжилттай бүртгэгдлээ!", "success")
            return redirect(url_for("login"))
        except Exception as e:
            db.session.rollback()
            flash(f"Бүртгэх үед алдаа: {e}", "error")
            return redirect(url_for("register"))

    return render_template("register.html")

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""

        user = User.query.filter_by(username=username).first()

        if user and check_password_hash(user.password_hash, password):
            if not can_login(user):
                log_action_mn("Block хэрэглэгч нэвтрэх оролдлого хийв")
                flash("Таны эрх түр хаагдсан байна. Админтай холбогдоно уу.", "error")
                return redirect(url_for("login"))

            login_user(user)
            log_action_mn("Системд нэвтэрлээ")
            return redirect(url_for("index"))

        flash("Нэвтрэх нэр эсвэл нууц үг буруу", "error")
        log_action_mn("Нэвтрэх оролдлого амжилтгүй")
        return redirect(url_for("login"))

    return render_template("login.html")

@app.route("/logout")
@login_required
def logout():
    log_action_mn("Системээс гарлаа")
    logout_user()
    return redirect(url_for("login"))

@app.route("/change_password", methods=["GET", "POST"])
@login_required
def change_password():
    if request.method == "POST":
        current_password = request.form.get("current_password") or ""
        new_password = request.form.get("new_password") or ""
        confirm_password = request.form.get("confirm_password") or ""

        if not current_password or not new_password or not confirm_password:
            flash("Бүх талбарыг бөглөнө үү.", "error")
            return redirect(url_for("change_password"))

        if not check_password_hash(current_user.password_hash, current_password):
            flash("Одоогийн нууц үг буруу байна.", "error")
            return redirect(url_for("change_password"))

        if len(new_password) < 6:
            flash("Шинэ нууц үг хамгийн багадаа 6 тэмдэгттэй байна.", "error")
            return redirect(url_for("change_password"))

        if new_password != confirm_password:
            flash("Шинэ нууц үг давтан оруулсантай таарахгүй байна.", "error")
            return redirect(url_for("change_password"))

        try:
            current_user.password_hash = generate_password_hash(new_password)
            db.session.flush()
            log_action_mn(f"Нууц үгээ сольсон ({current_user.username})", auto_commit=False)
            commit_with_retry()

            flash("Нууц үг амжилттай солигдлоо.", "success")
            return redirect(url_for("index"))
        except Exception as e:
            db.session.rollback()
            flash(f"Нууц үг солих үед алдаа: {e}", "error")
            return redirect(url_for("change_password"))

    return render_template("change_password.html")

@app.route("/")
@login_required
def index():
    selected_auditor = (request.args.get("auditor") or "").strip()
    selected_company = (request.args.get("company") or "").strip()

    selected_period = (request.args.get("period") or "month").strip()
    date_from = (request.args.get("date_from") or "").strip()
    date_to = (request.args.get("date_to") or "").strip()

    df = parse_date(date_from) if date_from else None
    dt = parse_date(date_to) if date_to else None

    levels = ["Өндөр", "Дунд", "Бага"]

    base_q = apply_issue_filters(Issue.query, selected_auditor, selected_company, df, dt)
    issues_filtered = base_q.all()
    issue_ids = [i.id for i in issues_filtered]
    total_issues = len(issues_filtered)

    total_followups = 0
    issues_with_followup = 0
    if issue_ids:
        total_followups = int(
            db.session.query(func.count(FollowUp.id))
            .filter(FollowUp.issue_id.in_(issue_ids))
            .scalar() or 0
        )
        issues_with_followup = int(
            db.session.query(func.count(func.distinct(FollowUp.issue_id)))
            .filter(FollowUp.issue_id.in_(issue_ids))
            .scalar() or 0
        )
    issues_without_followup = max(total_issues - issues_with_followup, 0)

    auditors = [
        r[0] for r in db.session.query(Issue.auditor_name)
        .filter(Issue.auditor_name.isnot(None))
        .filter(Issue.auditor_name != "")
        .distinct()
        .order_by(Issue.auditor_name.asc())
        .all()
    ]
    companies = [
        r[0] for r in db.session.query(Issue.company_name)
        .filter(Issue.company_name.isnot(None))
        .filter(Issue.company_name != "")
        .distinct()
        .order_by(Issue.company_name.asc())
        .all()
    ]

    TOP_AUD = 15
    q_aud_top = db.session.query(
        Issue.auditor_name,
        func.count(Issue.id).label("cnt")
    )
    q_aud_top = apply_issue_filters(q_aud_top, selected_auditor, selected_company, df, dt)
    q_aud_top = (
        q_aud_top
        .group_by(Issue.auditor_name)
        .order_by(func.count(Issue.id).desc(), Issue.auditor_name.asc())
        .limit(TOP_AUD)
        .all()
    )

    auditor_top_labels = [(x[0] or "—") for x in q_aud_top]

    q_aud_lvl = db.session.query(
        Issue.auditor_name,
        Issue.residual_level,
        func.count(Issue.id)
    )
    q_aud_lvl = apply_issue_filters(q_aud_lvl, selected_auditor, selected_company, df, dt)
    q_aud_lvl = q_aud_lvl.group_by(Issue.auditor_name, Issue.residual_level).all()

    auditor_risk_labels, auditor_risk_levels, auditor_risk_matrix = build_stacked_matrix_from_rows(
        q_aud_lvl,
        labels_order=auditor_top_labels,
        levels_order=levels
    )

    TOP_COMP = 12
    q_comp_top = db.session.query(
        Issue.company_name,
        func.count(Issue.id).label("cnt")
    )
    q_comp_top = apply_issue_filters(q_comp_top, selected_auditor, selected_company, df, dt)
    q_comp_top = (
        q_comp_top
        .group_by(Issue.company_name)
        .order_by(func.count(Issue.id).desc(), Issue.company_name.asc())
        .limit(TOP_COMP)
        .all()
    )

    company_top_labels = [(x[0] or "—") for x in q_comp_top]

    q_comp_lvl = db.session.query(
        Issue.company_name,
        Issue.residual_level,
        func.count(Issue.id)
    )
    q_comp_lvl = apply_issue_filters(q_comp_lvl, selected_auditor, selected_company, df, dt)
    q_comp_lvl = q_comp_lvl.group_by(Issue.company_name, Issue.residual_level).all()

    company_risk_companies, company_risk_levels, company_risk_matrix = build_stacked_matrix_from_rows(
        q_comp_lvl,
        labels_order=company_top_labels,
        levels_order=levels
    )

    q_type = db.session.query(
        Issue.risk_classification,
        func.count(Issue.id)
    )
    q_type = apply_issue_filters(q_type, selected_auditor, selected_company, df, dt)
    q_type = (
        q_type
        .group_by(Issue.risk_classification)
        .order_by(func.count(Issue.id).desc())
        .all()
    )

    type_labels = [x[0] or "Бусад" for x in q_type]
    type_values = [int(x[1] or 0) for x in q_type]
    type_total = sum(type_values)

    q_type_lvl = db.session.query(
        Issue.risk_classification,
        Issue.residual_level,
        func.count(Issue.id)
    )
    q_type_lvl = apply_issue_filters(q_type_lvl, selected_auditor, selected_company, df, dt)
    q_type_lvl = q_type_lvl.group_by(Issue.risk_classification, Issue.residual_level).all()

    type_level_map = {lvl: [0] * len(type_labels) for lvl in levels}
    type_index = {t: i for i, t in enumerate(type_labels)}

    for t, lvl, cnt in q_type_lvl:
        t = t or "Бусад"
        lvl = (lvl or "").strip()
        if t in type_index and lvl in type_level_map:
            type_level_map[lvl][type_index[t]] = int(cnt or 0)

    risk_type_levels = levels
    risk_type_matrix = [type_level_map[lvl] for lvl in levels]

    init_labels = levels
    q_init = db.session.query(
        Issue.residual_level,
        func.count(Issue.id)
    )
    q_init = apply_issue_filters(q_init, selected_auditor, selected_company, df, dt)
    q_init = q_init.group_by(Issue.residual_level).all()

    init_map = {k: int(v or 0) for k, v in q_init}
    init_values = [init_map.get(l, 0) for l in init_labels]

    after_labels = levels
    latest_fu_rows = []

    if issue_ids:
        latest_fu_sub = (
            db.session.query(
                FollowUp.issue_id.label("issue_id"),
                func.max(FollowUp.id).label("max_id")
            )
            .filter(FollowUp.issue_id.in_(issue_ids))
            .group_by(FollowUp.issue_id)
            .subquery()
        )

        latest_fu_rows = (
            db.session.query(
                Issue.id.label("issue_id"),
                Issue.auditor_name.label("issue_auditor"),
                FollowUp.id.label("followup_id"),
                FollowUp.residual_level.label("followup_level"),
                FollowUp.auditor_name.label("followup_auditor")
            )
            .join(FollowUp, FollowUp.issue_id == Issue.id)
            .join(
                latest_fu_sub,
                (FollowUp.issue_id == latest_fu_sub.c.issue_id) &
                (FollowUp.id == latest_fu_sub.c.max_id)
            )
            .filter(Issue.id.in_(issue_ids))
            .all()
        )

    after_map = {lvl: 0 for lvl in levels}
    for row in latest_fu_rows:
        lvl = (row.followup_level or "").strip()
        if lvl in after_map:
            after_map[lvl] += 1

    after_values = [after_map.get(l, 0) for l in after_labels]

    def _include_loss(issue: Issue) -> bool:
        amt = parse_money_to_int(issue.money_amount)
        if amt <= 0:
            return False
        flag = (issue.has_actual_loss or "").strip()
        if is_yes_mn(flag):
            return True
        if flag == "":
            return True
        return False

    total_loss_amount = sum(
        parse_money_to_int(i.money_amount)
        for i in issues_filtered
        if _include_loss(i)
    )

    total_reduced_amount = 0
    if issue_ids:
        latest_fu_id_sub = (
            db.session.query(
                FollowUp.issue_id,
                func.max(FollowUp.id).label("mxid")
            )
            .filter(FollowUp.issue_id.in_(issue_ids))
            .group_by(FollowUp.issue_id)
            .subquery()
        )

        latest_fu_money_rows = (
            db.session.query(FollowUp.issue_id, FollowUp.reduced_amount)
            .join(
                latest_fu_id_sub,
                (FollowUp.issue_id == latest_fu_id_sub.c.issue_id) &
                (FollowUp.id == latest_fu_id_sub.c.mxid)
            )
            .all()
        )

        total_reduced_amount = sum(parse_money_to_int(r[1]) for r in latest_fu_money_rows)

    remaining_loss_amount = max(total_loss_amount - total_reduced_amount, 0)

    company_aimag = COMPANY_AIMAG

    level_rank = {"Өндөр": 3, "Дунд": 2, "Бага": 1}
    company_max_level: dict[str, str] = {}
    for comp in (company_aimag or {}).keys():
        q_comp_lvl = db.session.query(Issue.residual_level, func.count(Issue.id)).filter(Issue.company_name == comp)
        if df:
            q_comp_lvl = q_comp_lvl.filter(Issue.identified_date >= df)
        if dt:
            q_comp_lvl = q_comp_lvl.filter(Issue.identified_date <= dt)
        if selected_auditor:
            q_comp_lvl = q_comp_lvl.filter(Issue.auditor_name == selected_auditor)
        if selected_company:
            q_comp_lvl = q_comp_lvl.filter(Issue.company_name == selected_company)
        rows = q_comp_lvl.group_by(Issue.residual_level).all()

        best = ""
        best_r = 0
        for lvl, _cnt in rows:
            lv = (lvl or "").strip()
            r = level_rank.get(lv, 0)
            if r > best_r:
                best_r = r
                best = lv
        if best:
            company_max_level[comp] = best

    aimag_issue_counts = {}
    aimag_risk_split = {}

    for comp, aim in (company_aimag or {}).items():
        q_comp_cnt = Issue.query.filter(Issue.company_name == comp)
        if df:
            q_comp_cnt = q_comp_cnt.filter(Issue.identified_date >= df)
        if dt:
            q_comp_cnt = q_comp_cnt.filter(Issue.identified_date <= dt)
        if selected_auditor:
            q_comp_cnt = q_comp_cnt.filter(Issue.auditor_name == selected_auditor)
        if selected_company:
            q_comp_cnt = q_comp_cnt.filter(Issue.company_name == selected_company)
        cnt = q_comp_cnt.count()
        aimag_issue_counts[aim] = aimag_issue_counts.get(aim, 0) + cnt

    for comp, aim in (company_aimag or {}).items():
        q_comp_risk = db.session.query(Issue.residual_level, func.count(Issue.id)).filter(Issue.company_name == comp)
        if df:
            q_comp_risk = q_comp_risk.filter(Issue.identified_date >= df)
        if dt:
            q_comp_risk = q_comp_risk.filter(Issue.identified_date <= dt)
        if selected_auditor:
            q_comp_risk = q_comp_risk.filter(Issue.auditor_name == selected_auditor)
        if selected_company:
            q_comp_risk = q_comp_risk.filter(Issue.company_name == selected_company)
        rows = q_comp_risk.group_by(Issue.residual_level).all()

        if aim not in aimag_risk_split:
            aimag_risk_split[aim] = {"Өндөр": 0, "Дунд": 0, "Бага": 0}

        for lvl, c in rows:
            lvl = (lvl or "").strip()
            if lvl in aimag_risk_split[aim]:
                aimag_risk_split[aim][lvl] += int(c or 0)

    q_fu_top = db.session.query(
        FollowUp.auditor_name,
        func.count(FollowUp.id).label("cnt")
    ).join(Issue, Issue.id == FollowUp.issue_id)

    q_fu_top = apply_followup_filters(q_fu_top, selected_auditor, selected_company, df, dt)
    q_fu_top = (
        q_fu_top
        .filter(FollowUp.auditor_name.isnot(None))
        .filter(FollowUp.auditor_name != "")
        .group_by(FollowUp.auditor_name)
        .order_by(func.count(FollowUp.id).desc(), FollowUp.auditor_name.asc())
        .all()
    )

    fu_top_labels = [(x[0] or "—") for x in q_fu_top]

    q_fu_lvl = db.session.query(
        FollowUp.auditor_name,
        FollowUp.residual_level,
        func.count(FollowUp.id)
    ).join(Issue, Issue.id == FollowUp.issue_id)

    q_fu_lvl = apply_followup_filters(q_fu_lvl, selected_auditor, selected_company, df, dt)
    q_fu_lvl = (
        q_fu_lvl
        .filter(FollowUp.auditor_name.isnot(None))
        .filter(FollowUp.auditor_name != "")
        .group_by(FollowUp.auditor_name, FollowUp.residual_level)
        .all()
    )

    fu_auditor_labels, fu_auditor_levels, fu_auditor_matrix = build_stacked_matrix_from_rows(
        q_fu_lvl,
        labels_order=fu_top_labels,
        levels_order=levels
    )

    active_guidelines = []
    try:
        today = date.today()
        rows = _fetch_all_guidelines_for_list()
        for g in rows:
            start = getattr(g, "exec_start", None)
            end = getattr(g, "extended_end", None) or getattr(g, "exec_end", None)
            if start and end and start <= today <= end:
                active_guidelines.append(g)
    except Exception:
        active_guidelines = []

    return render_template(
        "index.html",
        total_issues=total_issues,
        total_followups=total_followups,
        issues_with_followup=issues_with_followup,
        issues_without_followup=issues_without_followup,

        auditors=auditors,
        companies=companies,
        selected_auditor=selected_auditor,
        selected_company=selected_company,
        selected_period=selected_period,
        date_from=date_from,
        date_to=date_to,

        total_loss_amount=total_loss_amount,
        total_reduced_amount=total_reduced_amount,
        remaining_loss_amount=remaining_loss_amount,

        auditor_risk_labels=auditor_risk_labels,
        auditor_risk_levels=auditor_risk_levels,
        auditor_risk_matrix=auditor_risk_matrix,

        company_risk_companies=company_risk_companies,
        company_risk_levels=company_risk_levels,
        company_risk_matrix=company_risk_matrix,

        type_labels=type_labels,
        type_values=type_values,
        type_total=type_total,
        risk_type_levels=risk_type_levels,
        risk_type_matrix=risk_type_matrix,

        init_labels=init_labels,
        init_values=init_values,
        after_labels=after_labels,
        after_values=after_values,

        fu_auditor_labels=fu_auditor_labels,
        fu_auditor_levels=fu_auditor_levels,
        fu_auditor_matrix=fu_auditor_matrix,

        company_aimag=company_aimag,
        company_max_level=company_max_level,
        aimag_issue_counts=aimag_issue_counts,
        aimag_risk_split=aimag_risk_split,
        active_guidelines=active_guidelines,
    )


def save_multi_evidence(prefix_name: str, max_files: int = 5) -> list[str]:
    saved: list[str] = []
    for i in range(1, max_files + 1):
        f = request.files.get(f"{prefix_name}{i}")
        if not f or not f.filename:
            continue
        if not allowed_file(f.filename):
            raise ValueError("Файлын төрөл дэмжигдэхгүй байна.")
        filename = secure_filename(f.filename)
        rnd = secrets.token_hex(8)
        final = f"{rnd}_{filename}"
        f.save(os.path.join(app.config["UPLOAD_FOLDER"], final))
        saved.append(final)
    return saved


def append_files(existing: str | None, new_files: list[str]) -> str | None:
    old = [x for x in (existing or "").split(";") if x.strip()]
    combined = old + new_files
    combined = [x for x in combined if x]
    return ";".join(combined) if combined else None


def remove_files(existing: str | None, remove_list: list[str]) -> str | None:
    old = [x for x in (existing or "").split(";") if x.strip()]
    rm = set([x for x in remove_list if x])
    left = [x for x in old if x not in rm]
    return ";".join(left) if left else None


@app.route("/submit", methods=["GET", "POST"])
@login_required
def submit():
    if request.method == "POST":
        detail_issue = (request.form.get("detail_issue") or "").strip()
        detail_context = (request.form.get("detail_context") or "").strip()
        detail_criteria = (request.form.get("detail_criteria") or "").strip()
        detail_cause = (request.form.get("detail_cause") or "").strip()
        detail_impact = (request.form.get("detail_impact") or "").strip()

        recommendation = (request.form.get("issue_text") or "").strip()
        recommendation_number = (request.form.get("recommendation_number") or "").strip()

        implement_due_date = parse_date(request.form.get("implement_due_date"))
        materiality = (request.form.get("materiality") or "").strip()
        report_name = (request.form.get("report_name") or "").strip()

        company_name = (request.form.get("company_name") or "").strip()
        auditor_name = (request.form.get("auditor_name") or "").strip()
        identified_date = parse_date(request.form.get("identified_date"))

        has_actual_loss = (request.form.get("has_actual_loss") or "").strip()
        money_amount = (request.form.get("money_amount") or "").strip()

        risk_event_kind = (request.form.get("risk_event_kind") or "").strip()
        risk_factor = (request.form.get("risk_factor") or "").strip()
        risk_classification = (request.form.get("risk_classification") or "").strip()
        sub_risk_classification = (request.form.get("sub_risk_classification") or "").strip()
        risk_owner = (request.form.get("risk_owner") or "").strip()

        controls_text = (request.form.get("controls_text") or "").strip()
        ctl_design = (request.form.get("ctl_design") or "").strip()
        ctl_operates = (request.form.get("ctl_operates") or "").strip()
        ctl_awareness = (request.form.get("ctl_awareness") or "").strip()
        control_pct = (request.form.get("control_pct") or "").strip()

        residual_score_raw = request.form.get("residual_score")
        residual_level = (request.form.get("residual_level") or "").strip()
        action_rank = (request.form.get("action_rank") or "").strip()
        risk_category = (request.form.get("risk_category") or "").strip()

        if not detail_issue or not detail_criteria or not company_name or not auditor_name or not identified_date or not implement_due_date or not materiality:
            flash("Заавал бөглөх талбаруудыг бүрэн бөглөнө үү.", "error")
            return redirect(url_for("submit"))

        try:
            new_files = save_multi_evidence("evidence_file_", max_files=5)
        except ValueError as e:
            flash(str(e), "error")
            return redirect(url_for("submit"))

        if not new_files:
            flash("Нотлох баримт заавал.", "error")
            return redirect(url_for("submit"))

        rs = None
        try:
            rs = int(residual_score_raw) if (residual_score_raw and str(residual_score_raw).isdigit()) else None
        except Exception:
            rs = None

        try:
            new_issue = Issue(
                company_name=company_name,
                auditor_name=auditor_name,
                issue=detail_issue,
                recommendation=recommendation,
                detail_issue=detail_issue,
                issue_text=recommendation,
                detail_context=detail_context,
                detail_criteria=detail_criteria,
                detail_cause=detail_cause,
                detail_impact=detail_impact,
                recommendation_number=recommendation_number,
                report_name=report_name,
                implement_due_date=implement_due_date,
                identified_date=identified_date,
                materiality=materiality,
                has_actual_loss=has_actual_loss,
                money_amount=money_amount,
                risk_event_kind=risk_event_kind,
                risk_factor=risk_factor,
                risk_classification=risk_classification,
                sub_risk_classification=sub_risk_classification,
                risk_owner=risk_owner,
                controls_text=controls_text,
                ctl_design=ctl_design,
                ctl_operates=ctl_operates,
                ctl_awareness=ctl_awareness,
                control_pct=control_pct,
                residual_score=rs,
                residual_level=residual_level,
                action_rank=action_rank,
                risk_category=risk_category,
                evidence_file=";".join(new_files),
            )
            db.session.add(new_issue)
            db.session.flush()

            log_action_mn(f"Шинээр асуудал бүртгэв (ID={new_issue.id})", auto_commit=False)
            commit_with_retry()

            flash("Асуудал амжилттай бүртгэгдлээ.", "success")
            return redirect(url_for("issues"))

        except Exception as e:
            db.session.rollback()
            flash(f"Хадгалах үед алдаа: {e}", "error")
            return redirect(url_for("submit"))

    return render_template("submit.html", COMPANIES=COMPANIES, AUDITORS=AUDITORS)


@app.route("/uploads/<path:filename>")
@login_required
def uploaded_file(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename, as_attachment=False)


@app.route("/issues")
@login_required
def issues():
    company = (request.args.get("company") or "").strip()
    q = Issue.query
    if company:
        q = q.filter(Issue.company_name == company)
    issues_list = q.order_by(Issue.id.desc()).all()
    return render_template("issues.html", issues=issues_list)


@app.route("/issue/<int:issue_id>")
@login_required
def issue_detail(issue_id: int):
    issue = Issue.query.get_or_404(issue_id)
    return render_template("issue_detail.html", issue=issue)


@app.route("/issue/<int:issue_id>/edit", methods=["GET", "POST"])
@login_required
def edit_issue(issue_id: int):
    admin_required()
    issue = Issue.query.get_or_404(issue_id)

    existing_files = [x for x in (issue.evidence_file or "").split(";") if x.strip()]

    if request.method == "POST":
        issue.detail_issue = (request.form.get("detail_issue") or "").strip()
        issue.issue = issue.detail_issue
        issue.detail_context = (request.form.get("detail_context") or "").strip()
        issue.detail_criteria = (request.form.get("detail_criteria") or "").strip()
        issue.detail_cause = (request.form.get("detail_cause") or "").strip()
        issue.detail_impact = (request.form.get("detail_impact") or "").strip()

        issue.recommendation = (request.form.get("issue_text") or "").strip()
        issue.issue_text = issue.recommendation
        issue.recommendation_number = (request.form.get("recommendation_number") or "").strip()

        issue.implement_due_date = parse_date(request.form.get("implement_due_date"))
        issue.materiality = (request.form.get("materiality") or "").strip()
        issue.report_name = (request.form.get("report_name") or "").strip()

        issue.company_name = (request.form.get("company_name") or "").strip()
        issue.auditor_name = (request.form.get("auditor_name") or "").strip()
        issue.identified_date = parse_date(request.form.get("identified_date"))

        issue.has_actual_loss = (request.form.get("has_actual_loss") or "").strip()
        issue.money_amount = (request.form.get("money_amount") or "").strip()

        issue.risk_event_kind = (request.form.get("risk_event_kind") or "").strip()
        issue.risk_factor = (request.form.get("risk_factor") or "").strip()
        issue.risk_classification = (request.form.get("risk_classification") or "").strip()
        issue.sub_risk_classification = (request.form.get("sub_risk_classification") or "").strip()
        issue.risk_owner = (request.form.get("risk_owner") or "").strip()

        issue.controls_text = (request.form.get("controls_text") or "").strip()
        issue.ctl_design = (request.form.get("ctl_design") or "").strip()
        issue.ctl_operates = (request.form.get("ctl_operates") or "").strip()
        issue.ctl_awareness = (request.form.get("ctl_awareness") or "").strip()
        issue.control_pct = (request.form.get("control_pct") or "").strip()

        rs = request.form.get("residual_score")
        try:
            issue.residual_score = int(rs) if (rs and str(rs).isdigit()) else None
        except Exception:
            issue.residual_score = None
        issue.residual_level = (request.form.get("residual_level") or "").strip()
        issue.action_rank = (request.form.get("action_rank") or "").strip()
        issue.risk_category = (request.form.get("risk_category") or "").strip()

        remove_list = request.form.getlist("remove_files")
        issue.evidence_file = remove_files(issue.evidence_file, remove_list)

        try:
            new_files = save_multi_evidence("evidence_file_", max_files=5)
        except ValueError as e:
            flash(str(e), "error")
            return redirect(url_for("edit_issue", issue_id=issue.id))

        if new_files:
            issue.evidence_file = append_files(issue.evidence_file, new_files)

        try:
            db.session.flush()
            log_action_mn(f"Асуудал засагдлаа (ID={issue.id})", auto_commit=False)
            commit_with_retry()
            flash("Амжилттай хадгаллаа.", "success")
            return redirect(url_for("issue_detail", issue_id=issue.id))
        except Exception as e:
            db.session.rollback()
            flash(f"Хадгалах үед алдаа: {e}", "error")
            return redirect(url_for("edit_issue", issue_id=issue.id))

    return render_template(
        "edit_issue.html",
        issue=issue,
        COMPANIES=COMPANIES,
        AUDITORS=AUDITORS,
        existing_files=existing_files
    )


@app.route("/issue/<int:issue_id>/delete", methods=["POST"])
@login_required
def delete_issue(issue_id: int):
    admin_required()
    issue = Issue.query.get_or_404(issue_id)

    try:
        FollowUp.query.filter(FollowUp.issue_id == issue_id).delete(synchronize_session=False)
        db.session.delete(issue)
        log_action_mn(f"Асуудал устгав (ID={issue_id})", auto_commit=False)
        commit_with_retry()

        flash("Устгалаа.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Устгах үед алдаа: {e}", "error")

    return redirect(url_for("issues"))


@app.route("/export/issues")
@login_required
def export_issues():
    rows = Issue.query.order_by(Issue.id.desc()).all()
    data = []
    for i in rows:
        data.append({
            "ID": i.id,
            "Бүртгэсэн огноо": i.identified_date,
            "Компанийн нэр": i.company_name,
            "Аудитор": i.auditor_name,
            "Эрсдэлийн ангилал": i.risk_classification,
            "Эрсдэлийн түвшин": i.residual_level,
            "Хяналтын үнэлгээ (%)": i.control_pct,
            "Асуудал": i.detail_issue or i.issue,
            "Зөвлөмж": i.recommendation or i.issue_text,
        })
    df = pd.DataFrame(data)
    log_action_mn("Issues Excel татав")
    return export_excel(df, "issues.xlsx", "Issues")


@app.route("/export/followups")
@login_required
def export_followups():
    rows = FollowUp.query.order_by(FollowUp.created_at.desc()).all()
    data = []
    for f in rows:
        data.append({
            "FollowUp ID": f.id,
            "Issue ID": f.issue_id,
            "Арга хэмжээний төрөл": f.action_type,
            "Зорилтот түвшин": f.target_level,
            "Арга хэмжээний төлөвлөгөө": f.action_plan,
            "Board date": f.board_date,
            "Deadline": f.deadline,
            "Хэрэгжилтийн тайлбар": f.note,
            "Хэрэгжилт (%)": f.impl_percent,
            "Зорилтдоо хүрсэн эсэх": f.target_hit,
            "Аудиторын дүгнэлт": f.auditor_conclusion,
            "Бүрэн хэрэгжсэн эсэх": f.fully_implemented,
            "Final %": f.final_percent,
            "Дараагийн анхаарах зүйл": f.follow_up,
            "Дүгнэсэн аудитор": f.auditor_name,
            "Үлдэгдэл оноо": f.residual_score,
            "Үлдэгдэл түвшин": f.residual_level,
            "Арга хэмжээний эрэмбэ": f.action_rank,
            "Нотлох баримт": f.evidence_file,
            "Бууруулсан мөнгөн дүн": getattr(f, "reduced_amount", None),
            "Бүртгэсэн огноо": f.created_at,
        })
    df = pd.DataFrame(data)
    log_action_mn("FollowUps Excel татав")
    return export_excel(df, "followups.xlsx", "FollowUps")


@app.route("/export/guidelines")
@login_required
def export_guidelines():
    rows = _fetch_all_guidelines_for_list()

    data = []
    for g in rows:
        data.append({
            "ID": getattr(g, "id", None),
            "Компанийн нэр": getattr(g, "company_name", None),
            "Аудитын төрөл": getattr(g, "audit_type", None),
            "Дэд төрөл": getattr(g, "audit_subtype", None),
            "Багийн ахлагч": getattr(g, "team_leader", None),
            "Багийн гишүүд": getattr(g, "team_members", None),
            "Хамрах хугацаа (start)": getattr(g, "scope_start", None),
            "Хамрах хугацаа (end)": getattr(g, "scope_end", None),
            "Хэрэгжүүлэх хугацаа (start)": getattr(g, "exec_start", None),
            "Хэрэгжүүлэх хугацаа (end)": getattr(g, "exec_end", None),
            "Сунгасан хугацаа": getattr(g, "extended_end", None),
            "Тайлбар": getattr(g, "extension_note", None),
            "Үүсгэсэн": getattr(g, "created_by", None),
            "Үүсгэсэн огноо": getattr(g, "created_at", None),
            "Батлагдсан PDF": getattr(g, "approved_pdf", None),
        })
    df = pd.DataFrame(data)
    log_action_mn("Guidelines Excel татав")
    return export_excel(df, "guidelines.xlsx", "Guidelines")


@app.route("/export/users")
@login_required
def export_users():
    admin_required()
    rows = User.query.order_by(User.id.desc()).all()
    data = []
    for u in rows:
        data.append({
            "ID": u.id,
            "Нэвтрэх нэр": u.username,
            "Имэйл": u.email,
            "Админ": "Тийм" if bool(u.is_admin) else "Үгүй",
            "Block": "Тийм" if bool(getattr(u, "is_blocked", False)) else "Үгүй",
            "Бүртгэсэн огноо": getattr(u, "created_at", None),
        })
    df = pd.DataFrame(data)
    log_action_mn("Users Excel татав")
    return export_excel(df, "users.xlsx", "Users")


@app.route("/logs")
@login_required
def logs():
    admin_required()

    rows = Log.query.order_by(Log.id.desc()).limit(1000).all()

    out = []
    for r in rows:
        ts = getattr(r, "timestamp", None)
        ts_mn = (ts + timedelta(hours=8)) if ts else None
        out.append({
            "timestamp": ts_mn,
            "user": r.user,
            "action": r.action,
        })

    return render_template("logs.html", logs=out)


@app.route("/followups")
@login_required
def followups():
    selected_auditor = (request.args.get("auditor") or "").strip()
    selected_company = (request.args.get("company") or "").strip()

    q = Issue.query

    if selected_auditor:
        q = q.filter(Issue.auditor_name == selected_auditor)

    if selected_company:
        q = q.filter(Issue.company_name == selected_company)

    issues_list = q.order_by(Issue.id.desc()).all()

    issue_ids = [i.id for i in issues_list]

    latest = {}
    if issue_ids:
        latest_fu_id = (
            db.session.query(
                FollowUp.issue_id,
                func.max(FollowUp.id).label("mxid")
            )
            .filter(FollowUp.issue_id.in_(issue_ids))
            .group_by(FollowUp.issue_id)
            .subquery()
        )

        latest_rows = (
            db.session.query(FollowUp)
            .join(
                latest_fu_id,
                (FollowUp.issue_id == latest_fu_id.c.issue_id) &
                (FollowUp.id == latest_fu_id.c.mxid)
            )
            .all()
        )
        latest = {fu.issue_id: fu for fu in latest_rows}

    auditors = [
        r[0] for r in db.session.query(Issue.auditor_name)
        .filter(Issue.auditor_name.isnot(None))
        .filter(Issue.auditor_name != "")
        .distinct()
        .order_by(Issue.auditor_name.asc())
        .all()
    ]

    companies = [
        r[0] for r in db.session.query(Issue.company_name)
        .filter(Issue.company_name.isnot(None))
        .filter(Issue.company_name != "")
        .distinct()
        .order_by(Issue.company_name.asc())
        .all()
    ]

    return render_template(
        "followups_list.html",
        issues=issues_list,
        latest=latest,
        today=date.today(),
        auditors=auditors,
        companies=companies,
        selected_auditor=selected_auditor,
        selected_company=selected_company,
    )

@app.route("/followup/<int:issue_id>")
@login_required
def followup_form(issue_id: int):
    issue = Issue.query.get_or_404(issue_id)

    followup = (
        FollowUp.query
        .filter_by(issue_id=issue_id)
        .order_by(FollowUp.id.desc())
        .first()
    )

    return render_template(
        "followups.html",
        issue=issue,
        followup=followup,
        AUDITORS=AUDITORS
    )

@app.route("/followup/<int:issue_id>/submit", methods=["POST"])
@login_required
def followup_submit(issue_id: int):
    Issue.query.get_or_404(issue_id)

    saved = []
    files = request.files.getlist("evidence_files")
    for f in files:
        if not f or not f.filename:
            continue
        if not allowed_file(f.filename):
            flash("Файлын төрөл дэмжигдэхгүй байна.", "error")
            return redirect(url_for("followup_form", issue_id=issue_id))
        fn = secure_filename(f.filename)
        rnd = secrets.token_hex(8)
        final = f"{rnd}_{fn}"
        f.save(os.path.join(app.config["UPLOAD_FOLDER"], final))
        saved.append(final)

    fu = FollowUp(issue_id=issue_id)
    db.session.add(fu)

    fu.action_type = (request.form.get("trt_type") or "").strip()
    fu.target_level = (request.form.get("trt_target_level") or "").strip()
    fu.action_plan = (request.form.get("action_plan") or "").strip()
    fu.board_date = parse_date(request.form.get("board_date"))
    fu.deadline = parse_date(request.form.get("deadline"))
    fu.note = (request.form.get("impl_desc") or "").strip()
    fu.impl_percent = (request.form.get("impl_percent") or "").strip()
    fu.target_hit = (request.form.get("trt_target_hit") or "").strip()
    fu.auditor_conclusion = (request.form.get("auditor_conclusion") or "").strip()
    fu.fully_implemented = (request.form.get("fully_implemented") or "").strip()
    fu.final_percent = (request.form.get("final_percent") or "").strip()
    fu.follow_up = (request.form.get("follow_up") or "").strip()
    fu.auditor_name = (request.form.get("auditor_name") or "").strip()

    fu.r5_financial = (request.form.get("r5_financial") or "").strip()
    fu.r5_regulatory = (request.form.get("r5_regulatory") or "").strip()
    fu.r5_stakeholder = (request.form.get("r5_stakeholder") or "").strip()
    fu.r5_operation = (request.form.get("r5_operation") or "").strip()
    fu.r5_health = (request.form.get("r5_health") or "").strip()
    fu.r5_probability = (request.form.get("r5_probability") or "").strip()

    fu.reduced_amount = (request.form.get("reduced_amount") or "").strip()

    rs = request.form.get("residual_score")
    try:
        fu.residual_score = int(rs) if (rs and str(rs).isdigit()) else None
    except Exception:
        fu.residual_score = None

    fu.residual_level = (request.form.get("residual_level") or "").strip()
    fu.action_rank = (request.form.get("action_rank") or "").strip()

    if saved:
        fu.evidence_file = ";".join(saved)
    else:
        fu.evidence_file = None

    try:
        db.session.flush()
        log_action_mn(f"Эргэн хяналт хадгалав (IssueID={issue_id}, FollowUpID={fu.id})", auto_commit=False)
        commit_with_retry()

        flash("Эргэн хяналт амжилттай хадгалагдлаа.", "success")
        return redirect(url_for("followups"))
    except Exception as e:
        db.session.rollback()
        flash(f"Хадгалах үед алдаа: {e}", "error")
        return redirect(url_for("followup_form", issue_id=issue_id))


@app.route("/download_template/<path:filename>")
@login_required
def download_template(filename):
    return send_from_directory(TEMPLATE_FILES_DIR, filename, as_attachment=True)


@app.route("/outlook_draft/<int:issue_id>")
@login_required
def outlook_draft_com(issue_id: int):
    issue = Issue.query.get_or_404(issue_id)
    to = (request.args.get("to") or "").strip()
    if not to:
        return jsonify(ok=False, error="to параметр байхгүй байна.")

    if win32 is None or pythoncom is None:
        return jsonify(ok=False, error="Энэ компьютер дээр Outlook COM (pywin32) боломжгүй байна.")

    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To = to
        mail.Subject = f"Дотоод аудитын зөвлөмжид эргэн хяналт хийх тухай – #{issue.id} ({issue.company_name or ''})"

        html_body = f"""
<html>
<body style="font-family:Calibri,Arial,sans-serif; font-size:14px;">
<p>Сайн байна уу,</p>

<p>Доорх аудитын зөвлөмжийн хэрэгжилтийн талаар хавсралт файлыг бөглөж, нотлох баримттай мэдээлэл ирүүлнэ үү.</p>

<p><b>Компанийн нэр:</b> {issue.company_name or '—'}<br>
<b>Асуудал ID:</b> {issue.id}</p>

<p><b>Асуудал:</b><br>
{(issue.detail_issue or issue.issue or '—')}</p>

<p><b>Зөвлөмж:</b><br>
{(issue.recommendation or issue.issue_text or '—')}</p>

<p>Хүндэтгэсэн,<br>
{current_user.username}</p>
</body>
</html>
"""
        mail.HTMLBody = html_body

        tpl_path = os.path.join(TEMPLATE_FILES_DIR, TEMPLATE_XLSX)
        if os.path.exists(tpl_path):
            mail.Attachments.Add(tpl_path)

        mail.Display(False)
        log_action_mn(f"Outlook draft нээлээ (IssueID={issue_id})")
        return jsonify(ok=True)

    except Exception as e:
        return jsonify(ok=False, error=str(e))

    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


@app.route("/issues/<int:issue_id>/history")
@login_required
def issue_history(issue_id):
    issue = Issue.query.get_or_404(issue_id)

    needle = f"ID={issue_id}"
    logs = (
        Log.query
        .filter(Log.action.like(f"%{needle}%"))
        .order_by(Log.timestamp.desc())
        .all()
    )

    return render_template("issue_history.html", issue=issue, logs=logs)


@app.route("/guideline", methods=["GET", "POST"])
@login_required
def guideline():
    if request.method == "POST":
        company_name = (request.form.get("company_name") or "").strip()
        audit_type = (request.form.get("audit_type") or "").strip()
        audit_subtype = (request.form.get("audit_subtype") or "").strip()
        team_leader = (request.form.get("team_leader") or "").strip()

        members = request.form.getlist("team_members")
        members = [m.strip() for m in members if m and m.strip()]
        team_members = ";".join(members) if members else ""

        scope_start = parse_date(request.form.get("scope_start"))
        scope_end = parse_date(request.form.get("scope_end"))
        exec_start = parse_date(request.form.get("exec_start"))
        exec_end = parse_date(request.form.get("exec_end"))

        extended_end = parse_date(request.form.get("extended_end"))
        extension_note = (request.form.get("extension_note") or "").strip()

        if not company_name or not audit_type or not audit_subtype or not team_leader:
            flash("Компанийн нэр / Аудитын төрөл / Аудитын дэд төрөл / Аудитын багын ахлагч заавал.", "error")
            return redirect(url_for("guideline"))

        if not scope_start or not scope_end or not exec_start or not exec_end:
            flash("Хугацааны огноонуудыг бүрэн оруулна уу.", "error")
            return redirect(url_for("guideline"))

        try:
            new_id = _insert_guideline_raw(
                company_name=company_name,
                audit_type=audit_type,
                audit_subtype=audit_subtype,
                team_leader=team_leader,
                team_members=team_members,
                scope_start=scope_start,
                scope_end=scope_end,
                exec_start=exec_start,
                exec_end=exec_end,
                extended_end=extended_end,
                extension_note=extension_note,
                created_by=current_user.username
            )
            log_action_mn(f"Удирдамж үүсгэв (ID={new_id})", auto_commit=False)
            commit_with_retry()

            flash("Удирдамж амжилттай үүсгэлээ.", "success")
            return redirect(url_for("guideline_list"))
        except Exception as e:
            db.session.rollback()
            flash(f"Удирдамж үүсгэх үед алдаа: {e}", "error")
            return redirect(url_for("guideline"))

    return render_template("guideline.html", COMPANIES=COMPANIES, AUDITORS=AUDITORS)


@app.route("/guideline_list")
@login_required
def guideline_list():
    rows = []
    try:
        rows = _fetch_all_guidelines_for_list()
    except Exception as e:
        flash(f"Удирдамж унших үед алдаа: {e}", "error")
        rows = []

    return render_template("guideline_list.html", guidelines=rows)


@app.route("/guidelines")
@login_required
def guidelines():
    return guideline_list()


@app.route("/_diag/db")
@login_required
def diag_db():
    db_uri = app.config.get("SQLALCHEMY_DATABASE_URI", "")
    db_path = db_uri.replace("sqlite:///", "") if db_uri.startswith("sqlite:///") else ""

    table_rows = []
    for t in ("guideline", "guidelines"):
        if _sqlite_table_exists(t):
            table_rows.append({"table": t, "count": _table_count(t)})

    return jsonify(
        ok=True,
        current_user=current_user.username,
        is_admin=bool(getattr(current_user, "is_admin", False)),
        db_uri=db_uri,
        db_path=db_path,
        chosen_read_table=_guideline_table_name(),
        chosen_write_table=_guideline_write_table_name(),
        tables=table_rows,
    )


@app.route("/users_list")
@login_required
def users_list():
    admin_required()
    users = User.query.order_by(User.id.desc()).all()
    return render_template("users_list.html", users=users)


@app.route("/export/logs")
@login_required
def export_logs():
    admin_required()

    rows = Log.query.order_by(Log.id.desc()).limit(5000).all()

    data = []
    for l in rows:
        ts = getattr(l, "timestamp", None)
        if ts:
            ts = ts + timedelta(hours=8)

        data.append({
            "ID": l.id,
            "Хэрэглэгч": l.user,
            "Үйлдэл": l.action,
            "Цаг": ts,
        })

    df = pd.DataFrame(data)
    log_action_mn("Logs Excel татав")
    return export_excel(df, "logs.xlsx", "Logs")


@app.route("/permissions", methods=["GET"])
@login_required
def permissions():
    admin_required()
    users = User.query.order_by(User.id.desc()).all()
    return render_template("permissions.html", users=users)


@app.route("/permissions/<int:user_id>/grant_admin", methods=["POST"])
@login_required
def grant_admin(user_id: int):
    admin_required()
    u = User.query.get_or_404(user_id)
    try:
        u.is_admin = True
        db.session.flush()
        log_action_mn(f"Админ эрх олгов ({u.username})", auto_commit=False)
        commit_with_retry()
        flash("Админ эрх олголоо.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Алдаа: {e}", "error")
    return redirect(url_for("permissions"))


@app.route("/permissions/<int:user_id>/reset_password", methods=["POST"])
@login_required
def reset_user_password(user_id: int):
    admin_required()
    u = User.query.get_or_404(user_id)

    new_password = (request.form.get("new_password") or "").strip()

    if not new_password:
        flash("Шинэ нууц үг оруулна уу.", "error")
        return redirect(url_for("permissions"))

    if len(new_password) < 6:
        flash("Нууц үг хамгийн багадаа 6 тэмдэгттэй байна.", "error")
        return redirect(url_for("permissions"))

    try:
        u.password_hash = generate_password_hash(new_password)
        db.session.flush()
        log_action_mn(f"Хэрэглэгчийн нууц үгийг reset хийв ({u.username})", auto_commit=False)
        commit_with_retry()
        flash(f"{u.username} хэрэглэгчийн нууц үг шинэчлэгдлээ.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Алдаа: {e}", "error")

    return redirect(url_for("permissions"))


@app.route("/permissions/<int:user_id>/delete", methods=["POST"])
@login_required
def delete_user(user_id: int):
    admin_required()
    u = User.query.get_or_404(user_id)

    if u.id == current_user.id:
        flash("Өөрийгөө устгах боломжгүй.", "error")
        return redirect(url_for("permissions"))

    try:
        username = u.username
        db.session.delete(u)
        log_action_mn(f"Хэрэглэгч устгав ({username})", auto_commit=False)
        commit_with_retry()
        flash("Хэрэглэгч устгагдлаа.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Устгах үед алдаа: {e}", "error")

    return redirect(url_for("permissions"))


@app.route("/guideline/<int:gid>")
@login_required
def guideline_detail(gid: int):
    g = _fetch_one_guideline_raw(gid)
    if not g:
        abort(404)

    issues = Issue.query.filter_by(
        company_name=(g.company_name or "")
    ).order_by(Issue.id.desc()).all()

    latest_fu_id = (
        db.session.query(
            FollowUp.issue_id,
            func.max(FollowUp.id).label("mxid")
        )
        .group_by(FollowUp.issue_id)
        .subquery()
    )

    latest_rows = (
        db.session.query(FollowUp)
        .join(
            latest_fu_id,
            (FollowUp.issue_id == latest_fu_id.c.issue_id) &
            (FollowUp.id == latest_fu_id.c.mxid)
        )
        .all()
    )

    latest = {fu.issue_id: fu for fu in latest_rows}

    return render_template(
        "guideline_detail.html",
        g=g,
        issues=issues,
        latest=latest
    )


@app.route("/guideline/<int:gid>/edit", methods=["GET", "POST"])
@login_required
def guideline_edit(gid: int):
    admin_required()

    g = _fetch_one_guideline_raw(gid)
    if not g:
        abort(404)

    if request.method == "POST":
        company_name = (request.form.get("company_name") or "").strip()
        audit_type = (request.form.get("audit_type") or "").strip()
        audit_subtype = (request.form.get("audit_subtype") or "").strip()
        team_leader = (request.form.get("team_leader") or "").strip()

        members = request.form.getlist("team_members")
        members = [m.strip() for m in members if m.strip()]
        team_members = ";".join(members)

        scope_start = parse_date(request.form.get("scope_start"))
        scope_end = parse_date(request.form.get("scope_end"))
        exec_start = parse_date(request.form.get("exec_start"))
        exec_end = parse_date(request.form.get("exec_end"))
        extended_end = parse_date(request.form.get("extended_end"))
        extension_note = (request.form.get("extension_note") or "").strip()

        try:
            ok = _update_guideline_raw(
                gid=gid,
                company_name=company_name,
                audit_type=audit_type,
                audit_subtype=audit_subtype,
                team_leader=team_leader,
                team_members=team_members,
                scope_start=scope_start,
                scope_end=scope_end,
                exec_start=exec_start,
                exec_end=exec_end,
                extended_end=extended_end,
                extension_note=extension_note
            )
            if not ok:
                db.session.rollback()
                flash("Удирдамж олдсонгүй.", "error")
                return redirect(url_for("guideline_list"))

            log_action_mn(f"Удирдамж засагдлаа (ID={gid})", auto_commit=False)
            commit_with_retry()

            flash("Амжилттай хадгаллаа.", "success")
            return redirect(url_for("guideline_detail", gid=gid))
        except Exception as e:
            db.session.rollback()
            flash(f"Хадгалах үед алдаа: {e}", "error")
            return redirect(url_for("guideline_edit", gid=gid))

    selected_members = [x.strip() for x in (getattr(g, "team_members", "") or "").split(";") if x.strip()]

    return render_template(
        "guideline_edit.html",
        g=g,
        COMPANIES=COMPANIES,
        AUDITORS=AUDITORS,
        selected_members=selected_members
    )


@app.route("/guideline/<int:gid>/delete", methods=["POST"])
@login_required
def guideline_delete(gid: int):
    admin_required()

    try:
        ok = _delete_guideline_raw(gid)
        if not ok:
            db.session.rollback()
            flash("Удирдамжийн хүснэгт олдсонгүй.", "error")
            return redirect(url_for("guideline_list"))

        log_action_mn(f"Удирдамж устгав (ID={gid})", auto_commit=False)
        commit_with_retry()

        flash("Удирдамж амжилттай устгагдлаа.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Удирдамж устгах үед алдаа: {e}", "error")

    return redirect(url_for("guideline_list"))


@app.route("/guideline/<int:gid>/upload_pdf", methods=["POST"])
@login_required
def guideline_upload_pdf(gid: int):
    g = _fetch_one_guideline_raw(gid)
    if not g:
        abort(404)

    f = request.files.get("approved_pdf")
    if not f or not f.filename:
        flash("PDF файл сонгоно уу.", "error")
        return redirect(url_for("guideline_detail", gid=gid))

    ext = f.filename.rsplit(".", 1)[-1].lower() if "." in f.filename else ""
    if ext != "pdf":
        flash("Зөвхөн PDF файл зөвшөөрнө.", "error")
        return redirect(url_for("guideline_detail", gid=gid))

    fn = secure_filename(f.filename)
    rnd = secrets.token_hex(8)
    final = f"{rnd}_{fn}"
    save_path = os.path.join(app.config["UPLOAD_FOLDER"], final)

    try:
        f.save(save_path)

        table = getattr(g, "_source_table", None) or _guideline_table_name()
        if not table:
            flash("Удирдамжийн хүснэгт олдсонгүй.", "error")
            return redirect(url_for("guideline_detail", gid=gid))

        db.session.execute(
            text(f"UPDATE {table} SET approved_pdf=:p WHERE id=:id"),
            {"p": final, "id": gid}
        )
        db.session.flush()

        log_action_mn(f"Батлагдсан PDF хавсаргав (GuidelineID={gid})", auto_commit=False)
        commit_with_retry()

        flash("PDF амжилттай хавсаргалаа.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"PDF хавсаргах үед алдаа: {e}", "error")

    return redirect(url_for("guideline_detail", gid=gid))


@app.route("/permissions/<int:user_id>/revoke_admin", methods=["POST"])
@login_required
def revoke_admin(user_id: int):
    admin_required()
    u = User.query.get_or_404(user_id)

    if u.id == current_user.id:
        flash("Өөрийн админ эрхийг цуцлах боломжгүй.", "error")
        return redirect(url_for("permissions"))

    try:
        u.is_admin = False
        db.session.flush()
        log_action_mn(f"Админ эрх цуцлав ({u.username})", auto_commit=False)
        commit_with_retry()
        flash("Админ эрх цуцлагдлаа.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Алдаа: {e}", "error")

    return redirect(url_for("permissions"))


@app.route("/permissions/<int:user_id>/block", methods=["POST"])
@login_required
def block_user(user_id: int):
    admin_required()
    u = User.query.get_or_404(user_id)
    if u.id == current_user.id:
        flash("Өөрийгөө block хийх боломжгүй.", "error")
        return redirect(url_for("permissions"))

    try:
        u.is_blocked = True
        db.session.flush()
        log_action_mn(f"Хэрэглэгчийг block хийв ({u.username})", auto_commit=False)
        commit_with_retry()
        flash("Block хийлээ.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Алдаа: {e}", "error")

    return redirect(url_for("permissions"))


@app.route("/permissions/<int:user_id>/unblock", methods=["POST"])
@login_required
def unblock_user(user_id: int):
    admin_required()
    u = User.query.get_or_404(user_id)
    try:
        u.is_blocked = False
        db.session.flush()
        log_action_mn(f"Хэрэглэгч unblock хийв ({u.username})", auto_commit=False)
        commit_with_retry()
        flash("Unblock хийлээ.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Алдаа: {e}", "error")
    return redirect(url_for("permissions"))


def ensure_default_admin():
    if not User.query.filter_by(username="dag_admin").first():
        u = User(
            username="dag_admin",
            email="dag_admin@em.mn",
            password_hash=generate_password_hash("123"),
            is_admin=True
        )
        u.is_blocked = False
        db.session.add(u)
        commit_with_retry()


with app.app_context():
    db.create_all()
    ensure_sqlite_schema()
    ensure_sqlite_indexes()
    ensure_default_admin()

print("RUNNING:", __file__)
print("DB URI =", app.config["SQLALCHEMY_DATABASE_URI"])


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True, use_reloader=False, threaded=True)