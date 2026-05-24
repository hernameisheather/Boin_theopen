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
_data_cache = {
    "mtime": 0,
    "students": {},
    "records": [],
    "homework": "",       # 이번주 숙제 (전체 공통, 단일 값)
    "messages": [],       # 신쌤의 한마디: [{"student_code": "", "content": "..."}]
}


def _parse_date(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if v is None:
        return ""
    return str(v).strip()


def _parse_workbook(wb):
    """openpyxl Workbook을 (students, records, homework, messages)로 파싱.
    `기록` 시트는 두 가지 형식 모두 자동 감지:
      - 신형식: 날짜 | 학생코드 | 학생이름 | 항목 | 점수 | 피드백
      - 구형식: 날짜 | 학생코드 | 항목 | 점수 | 피드백
    `공지사항` 시트(선택): 종류 | 대상학생코드 | 내용
    """
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
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            header = [str(c).strip() if c is not None else "" for c in rows[0]]
            has_name_col = len(header) >= 3 and header[2] in ("학생이름", "이름", "name", "Name")

            for row in rows[1:]:
                if not row or row[0] is None or (len(row) > 1 and row[1] is None):
                    continue
                if has_name_col:
                    records.append({
                        "date": _parse_date(row[0]),
                        "student_code": str(row[1]).strip(),
                        "category": str(row[3]).strip() if len(row) > 3 and row[3] else "",
                        "score": str(row[4]).strip() if len(row) > 4 and row[4] is not None else "",
                        "feedback": str(row[5]).strip() if len(row) > 5 and row[5] else "",
                    })
                else:
                    records.append({
                        "date": _parse_date(row[0]),
                        "student_code": str(row[1]).strip(),
                        "category": str(row[2]).strip() if len(row) > 2 and row[2] else "",
                        "score": str(row[3]).strip() if len(row) > 3 and row[3] is not None else "",
                        "feedback": str(row[4]).strip() if len(row) > 4 and row[4] else "",
                    })

    homework = ""
    messages = []
    if "공지사항" in wb.sheetnames:
        ws = wb["공지사항"]
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            header = [str(c).strip() if c is not None else "" for c in rows[0]]
            # 신형식: 종류 | 대상학생코드 | 학생이름 | 내용
            # 구형식: 종류 | 대상학생코드 | 내용
            has_name_col = len(header) >= 3 and header[2] in ("학생이름", "이름", "name", "Name")
            content_idx = 3 if has_name_col else 2

            for row in rows[1:]:
                if not row or row[0] is None:
                    continue
                kind = str(row[0]).strip()
                code = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                content = str(row[content_idx]).strip() if len(row) > content_idx and row[content_idx] else ""
                if not content:
                    continue
                if kind in ("이번주숙제", "이번주 숙제", "숙제", "homework"):
                    homework = content
                elif kind in ("한마디", "신쌤의한마디", "신쌤의 한마디", "메시지", "message"):
                    messages.append({"student_code": code, "content": content})

    return students, records, homework, messages


def load_data():
    """Excel을 읽어서 메모리에 캐싱. 파일 수정시각이 바뀌면 재로딩."""
    if not os.path.exists(EXCEL_PATH):
        _data_cache.update({"mtime": 0, "students": {}, "records": [], "homework": "", "messages": []})
        return _data_cache

    mtime = os.path.getmtime(EXCEL_PATH)
    if mtime == _data_cache["mtime"]:
        return _data_cache

    wb = load_workbook(EXCEL_PATH, data_only=True)
    students, records, homework, messages = _parse_workbook(wb)

    _data_cache.update({
        "mtime": mtime,
        "students": students,
        "records": records,
        "homework": homework,
        "messages": messages,
    })
    return _data_cache


def save_data(students=None, records=None, homework=None, messages=None):
    """현재 메모리 상태를 Excel 파일로 저장.
    None을 전달하면 현재 캐시 값을 그대로 유지 (덮어쓰기 안 함).
    """
    if students is None or records is None or homework is None or messages is None:
        current = load_data()
        if students is None:
            students = current["students"]
        if records is None:
            records = current["records"]
        if homework is None:
            homework = current["homework"]
        if messages is None:
            messages = current["messages"]

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "학생명단"
    ws1.append(["학생코드", "학생이름", "PIN", "학부모이름(선택)"])
    for code, s in students.items():
        ws1.append([code, s.get("name", ""), s.get("pin", ""), s.get("parent", "")])

    ws2 = wb.create_sheet("기록")
    ws2.append(["날짜", "학생코드", "학생이름", "항목", "점수", "피드백"])
    for r in records:
        code = r.get("student_code", "")
        name = students.get(code, {}).get("name", "")
        ws2.append([
            r.get("date", ""),
            code,
            name,
            r.get("category", ""),
            r.get("score", ""),
            r.get("feedback", ""),
        ])

    ws3 = wb.create_sheet("공지사항")
    ws3.append(["종류", "대상학생코드", "학생이름", "내용"])
    if homework:
        ws3.append(["이번주숙제", "", "", homework])
    for m in messages:
        code = m.get("student_code", "")
        name = students.get(code, {}).get("name", "") if code else ""
        ws3.append(["한마디", code, name, m.get("content", "")])

    for ws, widths in [
        (ws1, [10, 12, 8, 18]),
        (ws2, [12, 10, 12, 14, 8, 50]),
        (ws3, [14, 14, 12, 60]),
    ]:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + i)].width = w

    wb.save(EXCEL_PATH)
    _data_cache["mtime"] = 0  # 캐시 무효화
    load_data()


def _parse_uploaded_excel(file_storage):
    """업로드된 Excel 파일을 임시 저장 후 파싱.
    반환: (students, records, homework, messages)
    """
    import tempfile
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        file_storage.save(tmp.name)
        tmp_path = tmp.name
    try:
        wb = load_workbook(tmp_path, data_only=True)
        return _parse_workbook(wb)
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass


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


def _parse_score(raw):
    """문자열 점수에서 숫자만 추출 (예: '92점' → 92.0). 실패 시 None."""
    if raw is None:
        return None
    s = str(raw).replace("점", "").replace(" ", "")
    try:
        return float(s)
    except (ValueError, AttributeError):
        return None


def _build_chart(points, width=560, height=160, padding=32):
    """날짜순으로 정렬된 [{date, score}] 리스트를 SVG 좌표 데이터로 변환."""
    if not points:
        return None
    inner_w = width - 2 * padding
    inner_h = height - 2 * padding
    n = len(points)
    scores = [p["score"] for p in points]
    min_s, max_s = min(scores), max(scores)
    rng = max(max_s - min_s, 10)
    y_min = max(0, min_s - rng * 0.15)
    y_max = min(100, max_s + rng * 0.15) if max_s <= 100 else max_s + rng * 0.15
    if y_max - y_min < 10:
        y_max = y_min + 10

    def x_at(i):
        return width / 2 if n == 1 else padding + (i / (n - 1)) * inner_w

    def y_at(v):
        return padding + (1 - (v - y_min) / (y_max - y_min)) * inner_h

    plot = [{
        "x": round(x_at(i), 1),
        "y": round(y_at(p["score"]), 1),
        "date": p["date"],
        "score": p["score"],
    } for i, p in enumerate(points)]

    path = "M " + " L ".join(f"{p['x']},{p['y']}" for p in plot)

    return {
        "width": width,
        "height": height,
        "padding": padding,
        "y_min": int(round(y_min)),
        "y_max": int(round(y_max)),
        "points": plot,
        "path": path,
    }


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

    # 학생 본인 카테고리별 통계
    stats = {}
    for cat, recs in grouped.items():
        nums = [v for v in (_parse_score(r["score"]) for r in recs) if v is not None]
        if nums:
            stats[cat] = {
                "count": len(nums),
                "avg": round(sum(nums) / len(nums), 1),
                "max": max(nums),
                "min": min(nums),
            }

    # 반 전체 카테고리별 평균 (모든 학생 합산)
    class_buckets = defaultdict(list)
    for r in data["records"]:
        v = _parse_score(r["score"])
        if v is not None:
            class_buckets[r["category"] or "기타"].append(v)
    class_avg = {
        cat: round(sum(vs) / len(vs), 1)
        for cat, vs in class_buckets.items() if vs
    }

    # 카테고리별 추세 그래프 데이터 (이 학생만, 날짜 오름차순)
    charts = {}
    for cat, recs in grouped.items():
        pts = []
        for r in recs:
            v = _parse_score(r["score"])
            if v is not None and r["date"]:
                pts.append({"date": r["date"], "score": v})
        pts.sort(key=lambda p: p["date"])
        if len(pts) >= 2:
            charts[cat] = _build_chart(pts)

    # 이 학생에게 보여줄 한마디: 공통(빈 코드) + 본인 개별
    my_messages = [
        m for m in data["messages"]
        if not m.get("student_code") or m.get("student_code") == code
    ]

    # 반 등수 계산: 항목명 기준 (날짜 무관, 모든 반 통합)
    # 같은 학생이 같은 항목에 여러 점수가 있으면 최고점 기준으로 순위 산정
    cat_best = defaultdict(dict)  # cat -> {code -> best_score}
    for r in data["records"]:
        sc = _parse_score(r["score"])
        if sc is None:
            continue
        cat = r["category"] or "기타"
        sc_code = r["student_code"]
        if sc_code not in cat_best[cat] or sc > cat_best[cat][sc_code]:
            cat_best[cat][sc_code] = sc

    # 경쟁식 순위 (동점은 같은 등수, 다음 등수는 인원만큼 건너뜀): 1, 2, 2, 4
    ranks = {}  # "code|category" -> (rank, total)
    for cat, code_scores in cat_best.items():
        if len(code_scores) < 2:
            continue  # 응시자 1명이면 등수 의미 없음
        sorted_entries = sorted(code_scores.items(), key=lambda x: -x[1])
        prev_score = None
        current_rank = 0
        for i, (c, s) in enumerate(sorted_entries):
            if s != prev_score:
                current_rank = i + 1
                prev_score = s
            ranks[f"{c}|{cat}"] = (current_rank, len(code_scores))

    # 날짜별 그룹핑 (이 학생만, 날짜 내림차순)
    by_date_dict = defaultdict(list)
    for r in items:
        by_date_dict[r["date"]].append(r)
    by_date = sorted(by_date_dict.items(), key=lambda x: x[0], reverse=True)

    return render_template(
        "student.html",
        student=student,
        records=items,
        grouped=grouped,
        stats=stats,
        class_avg=class_avg,
        charts=charts,
        homework=data["homework"],
        messages=my_messages,
        by_date=by_date,
        ranks=ranks,
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
        mode = request.form.get("mode", "overwrite")

        if f.filename:
            filename = secure_filename(f.filename)
            if not filename.lower().endswith((".xlsx", ".xlsm")):
                flash("Excel(.xlsx) 파일만 업로드 가능합니다.", "error")
            else:
                try:
                    new_students, new_records, new_homework, new_messages = _parse_uploaded_excel(f)
                except Exception as e:
                    flash(f"엑셀 파일을 읽을 수 없습니다: {e}", "error")
                    return redirect(url_for("admin"))

                if mode == "append":
                    current = load_data()

                    # 학생: 신규 코드만 추가, 기존 코드는 건너뜀
                    merged_students = dict(current["students"])
                    added_students = 0
                    skipped_students = 0
                    for code, s in new_students.items():
                        if code not in merged_students:
                            merged_students[code] = s
                            added_students += 1
                        else:
                            skipped_students += 1

                    # 기록 중복 판별: (날짜, 학생코드, 항목, 점수, 피드백) 5개 모두 일치
                    def _rec_key(r):
                        return (r.get("date", ""), r.get("student_code", ""),
                                r.get("category", ""), r.get("score", ""),
                                r.get("feedback", ""))
                    seen_recs = {_rec_key(r) for r in current["records"]}
                    unique_records = []
                    skipped_records = 0
                    for r in new_records:
                        k = _rec_key(r)
                        if k in seen_recs:
                            skipped_records += 1
                        else:
                            seen_recs.add(k)
                            unique_records.append(r)
                    merged_records = current["records"] + unique_records

                    # 이번주 숙제: 업로드 파일에 있으면 교체
                    homework_changed = bool(new_homework) and new_homework != current["homework"]
                    final_homework = new_homework if new_homework else current["homework"]

                    # 한마디 중복 판별: (학생코드, 내용) 모두 일치
                    def _msg_key(m):
                        return (m.get("student_code", ""), m.get("content", ""))
                    seen_msgs = {_msg_key(m) for m in current["messages"]}
                    unique_msgs = []
                    skipped_msgs = 0
                    for m in new_messages:
                        k = _msg_key(m)
                        if k in seen_msgs:
                            skipped_msgs += 1
                        else:
                            seen_msgs.add(k)
                            unique_msgs.append(m)
                    merged_messages = current["messages"] + unique_msgs

                    save_data(merged_students, merged_records, final_homework, merged_messages)

                    # 결과 메시지 조립
                    parts = []
                    parts.append(f"신규 학생 {added_students}명" +
                                 (f" (중복 {skipped_students}명 건너뜀)" if skipped_students else ""))
                    parts.append(f"새 기록 {len(unique_records)}건" +
                                 (f" (중복 {skipped_records}건 건너뜀)" if skipped_records else ""))
                    if homework_changed:
                        parts.append("이번주 숙제 갱신")
                    parts.append(f"새 한마디 {len(unique_msgs)}건" +
                                 (f" (중복 {skipped_msgs}건 건너뜀)" if skipped_msgs else ""))
                    flash("추가 완료: " + " · ".join(parts), "success")
                else:
                    # 덮어쓰기: 전체 교체
                    save_data(new_students, new_records, new_homework, new_messages)
                    flash(
                        f"덮어쓰기 완료: 학생 {len(new_students)}명, 기록 {len(new_records)}건, "
                        f"한마디 {len(new_messages)}건으로 전체 교체됨.",
                        "success"
                    )
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
            data["messages"] = [m for m in data["messages"] if m.get("student_code") != code]
            save_data(data["students"], data["records"], data["homework"], data["messages"])
            flash(f"학생 '{code}'와(과) 관련된 모든 기록·한마디가 삭제되었습니다.", "success")
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


# ─── 라우트: 관리자 — 공지사항(이번주 숙제 + 한마디) ────────
@app.route("/admin/announcements", methods=["GET", "POST"])
@admin_required
def admin_announcements():
    data = load_data()

    if request.method == "POST":
        action = request.form.get("action", "")

        if action == "update_homework":
            new_homework = request.form.get("homework", "").strip()
            save_data(homework=new_homework)
            if new_homework:
                flash("이번주 숙제가 업데이트되었습니다.", "success")
            else:
                flash("이번주 숙제가 삭제되었습니다.", "success")

        elif action == "add_message":
            content = request.form.get("content", "").strip()
            student_code = request.form.get("student_code", "").strip()
            if not content:
                flash("한마디 내용을 입력해주세요.", "error")
            elif student_code and student_code not in data["students"]:
                flash("선택한 학생코드가 존재하지 않습니다.", "error")
            else:
                new_messages = data["messages"] + [{"student_code": student_code, "content": content}]
                save_data(messages=new_messages)
                target = data["students"][student_code]["name"] if student_code else "전체"
                flash(f"한마디가 추가되었습니다. (대상: {target})", "success")

        elif action == "delete_message":
            try:
                idx = int(request.form.get("idx", -1))
            except ValueError:
                idx = -1
            if 0 <= idx < len(data["messages"]):
                new_messages = list(data["messages"])
                del new_messages[idx]
                save_data(messages=new_messages)
                flash("한마디가 삭제되었습니다.", "success")
            else:
                flash("삭제할 한마디를 찾을 수 없습니다.", "error")

        return redirect(url_for("admin_announcements"))

    return render_template(
        "admin_announcements.html",
        homework=data["homework"],
        messages=data["messages"],
        students=data["students"],
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
    """빈 Excel 템플릿 다운로드 (3시트: 학생명단, 기록, 공지사항)."""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "학생명단"
    ws1.append(["학생코드", "학생이름", "PIN", "학부모이름(선택)"])
    ws1.append(["S001", "김민지", "1234", "김민지 어머니"])
    ws1.append(["S002", "이도윤", "5678", "이도윤 어머니"])

    ws2 = wb.create_sheet("기록")
    ws2.append(["날짜", "학생코드", "학생이름", "항목", "점수", "피드백"])
    ws2.append(["2026-05-18", "S001", "김민지", "Daily Test", "92",
                "어휘 문제에서 실수가 있었지만 독해는 완벽했습니다."])
    ws2.append(["2026-05-18", "S001", "김민지", "숙제", "완료", "꼼꼼하게 잘 했습니다."])
    ws2.append(["2026-05-18", "S002", "이도윤", "Daily Test", "85",
                "문법 파트에서 2개 틀렸습니다. 복습 권장."])

    ws3 = wb.create_sheet("공지사항")
    ws3.append(["종류", "대상학생코드", "학생이름", "내용"])
    ws3.append(["이번주숙제", "", "", "워크북 32-45쪽 풀고, 단어 50개 외워오기"])
    ws3.append(["한마디", "", "", "시험 기간 화이팅! 모두 잘 할 수 있을거예요."])
    ws3.append(["한마디", "S001", "김민지", "단어시험 1등 축하해요. 다음주도 기대할게요!"])

    for ws, widths in [
        (ws1, [10, 12, 8, 18]),
        (ws2, [12, 10, 12, 14, 8, 50]),
        (ws3, [14, 14, 12, 60]),
    ]:
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
