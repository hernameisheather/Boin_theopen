"""
학생 성적/피드백 공유 포털
- 선생님: 관리자 페이지에서 학생/기록을 직접 추가·수정·삭제 (Excel 자동 저장)
- 학부모: 학생코드 + PIN으로 본인 자녀의 기록만 열람
"""
import os
import io
from datetime import datetime
from collections import defaultdict
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, send_file
)
from werkzeug.utils import secure_filename
from openpyxl import load_workbook, Workbook

# ─── 환경 변수 ────────────────────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

APP_SECRET = os.environ.get("APP_SECRET", "change-me-in-production")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "teacher1234")
DATA_DIR = os.environ.get("DATA_DIR", os.path.join(os.path.dirname(__file__), "data"))
EXCEL_PATH = os.path.join(DATA_DIR, "students.xlsx")

os.makedirs(DATA_DIR, exist_ok=True)

# ─── Flask 앱 ────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = APP_SECRET
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10MB

# ─── 데이터 ──────────────────────────────────────────────────
_data_cache = {"mtime": 0, "students": {}, "records": []}


def _parse_date(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if v is None:
        return ""
    return str(v).strip()


def load_data():
    """Excel을 읽어서 메모리에 캐싱. 파일 수정시각이 바뀌면 재로딩."""
    if not os.path.exists(EXCEL_PATH):
        _data_cache.update({"mtime": 0, "students": {}, "records": []})
        return _data_cache

    mtime = os.path.getmtime(EXCEL_PATH)
    if mtime == _data_cache["mtime"]:
        return _data_cache

    wb = load_workbook(EXCEL_PATH, data_only=True)

    students = {}
    if "학생명단" in wb.sheetnames:
        ws = wb["학생명단"]
        for row in list(ws.iter_rows(values_only=True))[1:]:
            if not row or row[0] is None:
                continue
            code = str(row[0]).strip()
            name = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            pin = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
            parent = str(row[3]).strip() if len(row) > 3 and row[3] else ""
            if code:
                students[code] = {"name": name, "pin": pin, "parent": parent}

    records = []
    if "기록" in wb.sheetnames:
        ws = wb["기록"]
        for row in list(ws.iter_rows(values_only=True))[1:]:
            if not row or row[0] is None or (len(row) > 1 and row[1] is None):
                continue
            records.append({
                "date": _parse_date(row[0]),
                "student_code": str(row[1]).strip(),
                "category": str(row[2]).strip() if len(row) > 2 and row[2] else "",
                "score": str(row[3]).strip() if len(row) > 3 and row[3] is not None else "",
                "feedback": str(row[4]).strip() if len(row) > 4 and row[4] else "",
            })

    _data_cache.update({"mtime": mtime, "students": students, "records": records})
    return _data_cache


def save_data(students, records):
    """현재 메모리 상태를 Excel 파일로 저장."""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "학생명단"
    ws1.append(["학생코드", "학생이름", "PIN", "학부모이름(선택)"])
    for code, s in students.items():
        ws1.append([code, s.get("name", ""), s.get("pin", ""), s.get("parent", "")])

    ws2 = wb.create_sheet("기록")
    ws2.append(["날짜", "학생코드", "항목", "점수", "피드백"])
    for r in records:
        ws2.append([
            r.get("date", ""),
            r.get("student_code", ""),
            r.get("category", ""),
            r.get("score", ""),
            r.get("feedback", ""),
        ])

    for ws, widths in [(ws1, [10, 12, 8, 18]), (ws2, [12, 10, 14, 8, 50])]:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + i)].width = w

    wb.save(EXCEL_PATH)
    _data_cache["mtime"] = 0  # 캐시 무효화
    load_data()


def records_for(student_code):
    """특정 학생의 기록을 날짜 내림차순으로 반환, 카테고리별 그룹핑."""
    data = load_data()
    items = [r for r in data["records"] if r["student_code"] == student_code]
    items.sort(key=lambda r: r["date"], reverse=True)

    grouped = defaultdict(list)
    for r in items:
        grouped[r["category"] or "기타"].append(r)
    return items, dict(grouped)


# ─── 인증 ────────────────────────────────────────────────────
def admin_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("is_admin"):
            return redirect(url_for("admin_login"))
        return f(*args, **kwargs)
    return wrapper


def parent_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("student_code"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper


# ─── 라우트: 학부모 ──────────────────────────────────────────
@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        code = request.form.get("student_code", "").strip()
        pin = request.form.get("pin", "").strip()

        data = load_data()
        student = data["students"].get(code)
        if not student:
            flash("학생코드가 올바르지 않습니다.", "error")
        elif student["pin"] and student["pin"] != pin:
            flash("PIN 번호가 일치하지 않습니다.", "error")
        else:
            session["student_code"] = code
            session["student_name"] = student["name"]
            return redirect(url_for("my_page"))

    return render_template("login.html")


@app.route("/me")
@parent_required
def my_page():
    code = session["student_code"]
    data = load_data()
    student = data["students"].get(code)
    if not student:
        session.clear()
        return redirect(url_for("login"))

    items, grouped = records_for(code)

    stats = {}
    for cat, recs in grouped.items():
        nums = []
        for r in recs:
            s = r["score"].replace("점", "").replace(" ", "")
            try:
                nums.append(float(s))
            except (ValueError, AttributeError):
                pass
        if nums:
            stats[cat] = {
                "count": len(nums),
                "avg": round(sum(nums) / len(nums), 1),
                "max": max(nums),
                "min": min(nums),
            }

    return render_template(
        "student.html",
        student=student,
        records=items,
        grouped=grouped,
        stats=stats,
        last_update=datetime.fromtimestamp(data["mtime"]).strftime("%Y-%m-%d %H:%M") if data["mtime"] else "—",
    )


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ─── 라우트: 관리자 ──────────────────────────────────────────
@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        if request.form.get("password") == ADMIN_PASSWORD:
            session["is_admin"] = True
            return redirect(url_for("admin"))
        flash("관리자 비밀번호가 올바르지 않습니다.", "error")
    return render_template("admin_login.html")


@app.route("/admin", methods=["GET", "POST"])
@admin_required
def admin():
    if request.method == "POST" and "excel_file" in request.files:
        f = request.files["excel_file"]
        if f.filename:
            filename = secure_filename(f.filename)
            if not filename.lower().endswith((".xlsx", ".xlsm")):
                flash("Excel(.xlsx) 파일만 업로드 가능합니다.", "error")
            else:
                f.save(EXCEL_PATH)
                _data_cache["mtime"] = 0
                load_data()
                flash("업로드 완료. 기존 데이터가 새 파일로 덮어씌워졌습니다.", "success")
        return redirect(url_for("admin"))

    data = load_data()
    return render_template(
        "admin.html",
        students=data["students"],
        record_count=len(data["records"]),
        last_update=datetime.fromtimestamp(data["mtime"]).strftime("%Y-%m-%d %H:%M") if data["mtime"] else "—",
        excel_exists=os.path.exists(EXCEL_PATH),
    )


@app.route("/admin/logout")
def admin_logout():
    session.pop("is_admin", None)
    return redirect(url_for("admin_login"))


# ─── 라우트: 관리자 — 기록 관리 ──────────────────────────────
@app.route("/admin/records", methods=["GET", "POST"])
@admin_required
def admin_records():
    data = load_data()

    if request.method == "POST":
        date = request.form.get("date", "").strip()
        student_code = request.form.get("student_code", "").strip()
        category = request.form.get("category", "").strip()
        score = request.form.get("score", "").strip()
        feedback = request.form.get("feedback", "").strip()

        if not date:
            flash("날짜를 입력해주세요.", "error")
        elif not student_code:
            flash("학생을 선택해주세요.", "error")
        elif student_code not in data["students"]:
            flash("등록되지 않은 학생코드입니다.", "error")
        else:
            data["records"].append({
                "date": date,
                "student_code": student_code,
                "category": category or "기타",
                "score": score,
                "feedback": feedback,
            })
            save_data(data["students"], data["records"])
            flash(f"{data['students'][student_code]['name']} 학생의 기록이 추가되었습니다.", "success")
        return redirect(url_for("admin_records",
                                student=request.args.get("student", ""),
                                category=request.args.get("category", "")))

    filter_student = request.args.get("student", "").strip()
    filter_category = request.args.get("category", "").strip()

    indexed = list(enumerate(data["records"]))
    if filter_student:
        indexed = [(i, r) for i, r in indexed if r["student_code"] == filter_student]
    if filter_category:
        indexed = [(i, r) for i, r in indexed if r["category"] == filter_category]

    indexed.sort(key=lambda x: x[1]["date"], reverse=True)

    categories = sorted({r["category"] for r in data["records"] if r["category"]})
    today = datetime.now().strftime("%Y-%m-%d")

    return render_template(
        "admin_records.html",
        records=indexed,
        students=data["students"],
        categories=categories,
        filter_student=filter_student,
        filter_category=filter_category,
        today=today,
    )


@app.route("/admin/records/<int:idx>/edit", methods=["GET", "POST"])
@admin_required
def admin_record_edit(idx):
    data = load_data()

    if idx < 0 or idx >= len(data["records"]):
        flash("기록을 찾을 수 없습니다.", "error")
        return redirect(url_for("admin_records"))

    if request.method == "POST":
        if request.form.get("action") == "delete":
            del data["records"][idx]
            save_data(data["students"], data["records"])
            flash("기록이 삭제되었습니다.", "success")
        else:
            data["records"][idx] = {
                "date": request.form.get("date", "").strip(),
                "student_code": request.form.get("student_code", "").strip(),
                "category": request.form.get("category", "").strip() or "기타",
                "score": request.form.get("score", "").strip(),
                "feedback": request.form.get("feedback", "").strip(),
            }
            save_data(data["students"], data["records"])
            flash("기록이 수정되었습니다.", "success")
        return redirect(url_for("admin_records"))

    return render_template(
        "admin_record_edit.html",
        record=data["records"][idx],
        idx=idx,
        students=data["students"],
        categories=sorted({r["category"] for r in data["records"] if r["category"]}),
    )


# ─── 라우트: 관리자 — 학생 관리 ──────────────────────────────
@app.route("/admin/students", methods=["GET", "POST"])
@admin_required
def admin_students():
    data = load_data()

    if request.method == "POST":
        code = request.form.get("student_code", "").strip()
        name = request.form.get("student_name", "").strip()
        pin = request.form.get("pin", "").strip()
        parent = request.form.get("parent", "").strip()

        if not code or not name:
            flash("학생코드와 이름은 필수입니다.", "error")
        elif code in data["students"]:
            flash(f"학생코드 '{code}'는 이미 사용 중입니다.", "error")
        else:
            data["students"][code] = {"name": name, "pin": pin, "parent": parent}
            save_data(data["students"], data["records"])
            flash(f"학생 '{name}'이(가) 추가되었습니다.", "success")
        return redirect(url_for("admin_students"))

    return render_template("admin_students.html", students=data["students"])


@app.route("/admin/students/<code>/edit", methods=["GET", "POST"])
@admin_required
def admin_student_edit(code):
    data = load_data()

    if code not in data["students"]:
        flash("학생을 찾을 수 없습니다.", "error")
        return redirect(url_for("admin_students"))

    if request.method == "POST":
        if request.form.get("action") == "delete":
            del data["students"][code]
            data["records"] = [r for r in data["records"] if r["student_code"] != code]
            save_data(data["students"], data["records"])
            flash(f"학생 '{code}'와(과) 관련된 모든 기록이 삭제되었습니다.", "success")
            return redirect(url_for("admin_students"))

        new_code = request.form.get("student_code", "").strip()
        new_name = request.form.get("student_name", "").strip()
        new_pin = request.form.get("pin", "").strip()
        new_parent = request.form.get("parent", "").strip()

        if not new_code or not new_name:
            flash("학생코드와 이름은 필수입니다.", "error")
            return redirect(url_for("admin_student_edit", code=code))

        if new_code != code and new_code in data["students"]:
            flash(f"학생코드 '{new_code}'는 이미 사용 중입니다.", "error")
            return redirect(url_for("admin_student_edit", code=code))

        if new_code != code:
            data["students"].pop(code)
            for r in data["records"]:
                if r["student_code"] == code:
                    r["student_code"] = new_code

        data["students"][new_code] = {"name": new_name, "pin": new_pin, "parent": new_parent}
        save_data(data["students"], data["records"])
        flash("학생 정보가 수정되었습니다.", "success")
        return redirect(url_for("admin_students"))

    return render_template(
        "admin_student_edit.html",
        student=data["students"][code],
        code=code,
    )


# ─── 라우트: Excel 다운로드/템플릿 ───────────────────────────
@app.route("/admin/download")
@admin_required
def admin_download():
    """현재 데이터를 Excel 파일로 다운로드 (백업)."""
    if not os.path.exists(EXCEL_PATH):
        flash("아직 데이터가 없습니다.", "error")
        return redirect(url_for("admin"))
    return send_file(
        EXCEL_PATH,
        as_attachment=True,
        download_name=f"students_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/admin/template")
def download_template():
    """빈 Excel 템플릿 다운로드."""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "학생명단"
    ws1.append(["학생코드", "학생이름", "PIN", "학부모이름(선택)"])
    ws1.append(["S001", "김민지", "1234", "김민지 어머니"])
    ws1.append(["S002", "이도윤", "5678", "이도윤 어머니"])

    ws2 = wb.create_sheet("기록")
    ws2.append(["날짜", "학생코드", "항목", "점수", "피드백"])
    ws2.append(["2026-05-18", "S001", "Daily Test", "92",
                "어휘 문제에서 실수가 있었지만 독해는 완벽했습니다."])
    ws2.append(["2026-05-18", "S001", "숙제", "완료", "꼼꼼하게 잘 했습니다."])

    for ws, widths in [(ws1, [10, 12, 8, 18]), (ws2, [12, 10, 14, 8, 50])]:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="students_template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ─── 헬스체크 ────────────────────────────────────────────────
@app.route("/healthz")
def healthz():
    return "ok", 200


# ─── 메인 ────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
