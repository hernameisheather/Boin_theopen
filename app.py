"""
학생 성적/피드백 공유 포털
- 선생님: 관리자 페이지에서 학생/기록을 직접 추가·수정·삭제 (Excel 자동 저장)
- 학부모: 학생코드 + PIN으로 본인 자녀의 기록만 열람
"""
import os
import io
import random
from datetime import datetime, timedelta
from collections import defaultdict
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, send_file
)
from werkzeug.utils import secure_filename
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.datavalidation import DataValidation

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
    "homeworks": [],      # 이번주 숙제 목록: [{"content": "...", "published": True/False}]
    "messages": [],       # 신쌤의 한마디: [{"student_code": "...", "content": "...", "published": True/False}]
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

            # 비고(플래그)와 완료(처리됨) 컬럼 위치 자동 탐지 — 어디에 있어도 OK
            flag_idx = None
            resolved_idx = None
            for i, h in enumerate(header):
                if h in ("비고", "특이사항", "플래그", "flag", "메모"):
                    flag_idx = i
                if h in ("완료", "해결", "처리", "처리됨", "resolved", "done"):
                    resolved_idx = i

            def _parse_resolved(v):
                if v is None:
                    return False
                s = str(v).strip().lower()
                return s in ("o", "완료", "true", "1", "y", "예", "✓", "v", "처리")

            for row in rows[1:]:
                if not row or row[0] is None or (len(row) > 1 and row[1] is None):
                    continue
                if has_name_col:
                    rec = {
                        "date": _parse_date(row[0]),
                        "student_code": str(row[1]).strip(),
                        "category": str(row[3]).strip() if len(row) > 3 and row[3] else "",
                        "score": str(row[4]).strip() if len(row) > 4 and row[4] is not None else "",
                        "feedback": str(row[5]).strip() if len(row) > 5 and row[5] else "",
                    }
                else:
                    rec = {
                        "date": _parse_date(row[0]),
                        "student_code": str(row[1]).strip(),
                        "category": str(row[2]).strip() if len(row) > 2 and row[2] else "",
                        "score": str(row[3]).strip() if len(row) > 3 and row[3] is not None else "",
                        "feedback": str(row[4]).strip() if len(row) > 4 and row[4] else "",
                    }
                rec["flag"] = (str(row[flag_idx]).strip()
                               if flag_idx is not None and len(row) > flag_idx and row[flag_idx]
                               else "")
                rec["resolved"] = (_parse_resolved(row[resolved_idx])
                                   if resolved_idx is not None and len(row) > resolved_idx
                                   else False)
                records.append(rec)

    homeworks = []
    messages = []
    if "공지사항" in wb.sheetnames:
        ws = wb["공지사항"]
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            header = [str(c).strip() if c is not None else "" for c in rows[0]]
            has_name_col = len(header) >= 3 and header[2] in ("학생이름", "이름", "name", "Name")
            content_idx = 3 if has_name_col else 2

            # 상태(게시/내림) 컬럼 자동 탐지
            status_idx = None
            for i, h in enumerate(header):
                if h in ("상태", "게시", "공개", "status", "published"):
                    status_idx = i

            def _is_published(v):
                if v is None:
                    return True  # 기본값은 게시
                s = str(v).strip().lower()
                if not s:
                    return True
                # 내림/비공개로 처리할 값들
                return s not in ("내림", "비공개", "off", "no", "n", "0", "false", "unpublished", "비활성", "x")

            for row in rows[1:]:
                if not row or row[0] is None:
                    continue
                kind = str(row[0]).strip()
                code = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                content = str(row[content_idx]).strip() if len(row) > content_idx and row[content_idx] else ""
                if not content:
                    continue
                published = True
                if status_idx is not None and len(row) > status_idx:
                    published = _is_published(row[status_idx])

                if kind in ("이번주숙제", "이번주 숙제", "숙제", "homework"):
                    homeworks.append({"content": content, "published": published})
                elif kind in ("한마디", "신쌤의한마디", "신쌤의 한마디", "메시지", "message"):
                    messages.append({"student_code": code, "content": content, "published": published})

    return students, records, homeworks, messages


def load_data():
    """Excel을 읽어서 메모리에 캐싱. 파일 수정시각이 바뀌면 재로딩."""
    if not os.path.exists(EXCEL_PATH):
        _data_cache.update({"mtime": 0, "students": {}, "records": [], "homeworks": [], "messages": []})
        return _data_cache

    mtime = os.path.getmtime(EXCEL_PATH)
    if mtime == _data_cache["mtime"]:
        return _data_cache

    wb = load_workbook(EXCEL_PATH, data_only=True)
    students, records, homeworks, messages = _parse_workbook(wb)

    _data_cache.update({
        "mtime": mtime,
        "students": students,
        "records": records,
        "homeworks": homeworks,
        "messages": messages,
    })
    return _data_cache


def _add_records_dropdowns(ws):
    """`기록` 시트의 비고(G열) / 완료(H열)에 드롭다운 추가."""
    # 비고 컬럼 — 재시/숙제미비/추가과제/결석/우수
    dv_flag = DataValidation(
        type="list",
        formula1='"재시,숙제미비,추가과제,결석,우수"',
        allow_blank=True,
        showErrorMessage=True,
    )
    dv_flag.error = "재시 / 숙제미비 / 추가과제 / 결석 / 우수 중에서 선택하세요. (자유 입력이 필요하면 데이터 검증 해제 후 입력)"
    dv_flag.errorTitle = "잘못된 비고"
    dv_flag.prompt = "드롭다운에서 선택"
    dv_flag.promptTitle = "비고 선택"
    ws.add_data_validation(dv_flag)
    dv_flag.add("G2:G2000")

    # 완료 컬럼 — O / 완료
    dv_done = DataValidation(
        type="list",
        formula1='"O,완료"',
        allow_blank=True,
        showErrorMessage=True,
    )
    dv_done.error = "처리됐으면 O 또는 완료를, 미처리는 비워두세요."
    dv_done.errorTitle = "잘못된 완료 표시"
    dv_done.prompt = "처리됐으면 O 선택"
    dv_done.promptTitle = "완료 표시"
    ws.add_data_validation(dv_done)
    dv_done.add("H2:H2000")


def _add_announcements_dropdown(ws):
    """`공지사항` 시트의 상태(E열)에 게시/내림 드롭다운."""
    dv = DataValidation(
        type="list",
        formula1='"게시,내림"',
        allow_blank=True,
        showErrorMessage=True,
    )
    dv.error = "게시 또는 내림 중에서 선택하세요."
    dv.errorTitle = "잘못된 상태"
    ws.add_data_validation(dv)
    dv.add("E2:E1000")


def save_data(students=None, records=None, homeworks=None, messages=None):
    """현재 메모리 상태를 Excel 파일로 저장.
    None을 전달하면 현재 캐시 값을 그대로 유지 (덮어쓰기 안 함).
    """
    if students is None or records is None or homeworks is None or messages is None:
        current = load_data()
        if students is None:
            students = current["students"]
        if records is None:
            records = current["records"]
        if homeworks is None:
            homeworks = current["homeworks"]
        if messages is None:
            messages = current["messages"]

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "학생명단"
    ws1.append(["학생코드", "학생이름", "PIN", "학부모이름(선택)"])
    for code, s in students.items():
        ws1.append([code, s.get("name", ""), s.get("pin", ""), s.get("parent", "")])

    ws2 = wb.create_sheet("기록")
    ws2.append(["날짜", "학생코드", "학생이름", "항목", "점수", "피드백", "비고", "완료"])
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
            r.get("flag", ""),
            "O" if r.get("resolved") else "",
        ])

    ws3 = wb.create_sheet("공지사항")
    ws3.append(["종류", "대상학생코드", "학생이름", "내용", "상태"])
    for h in homeworks:
        ws3.append([
            "이번주숙제",
            "",
            "",
            h.get("content", ""),
            "게시" if h.get("published") else "내림",
        ])
    for m in messages:
        code = m.get("student_code", "")
        name = students.get(code, {}).get("name", "") if code else ""
        ws3.append([
            "한마디",
            code,
            name,
            m.get("content", ""),
            "게시" if m.get("published") else "내림",
        ])

    for ws, widths in [
        (ws1, [10, 12, 8, 18]),
        (ws2, [12, 10, 12, 14, 8, 40, 12, 8]),
        (ws3, [14, 14, 12, 50, 10]),
    ]:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + i)].width = w

    _add_records_dropdowns(ws2)
    _add_announcements_dropdown(ws3)

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


# ─── 클리닉 자동 기록 (매주 화/목) ────────────────────────────
CLINIC_CATEGORY = "주간 클리닉"
CLINIC_BASE_FEEDBACK = "주간 혼공학습지 풀이 및 서술형 피드백 완료."
CLINIC_TIPS = [
    "[학습 포인트] 지문 속 복잡한 구문 — 주절과 종속절을 먼저 분리해서 분석하는 연습 필요.",
    "[학습 포인트] 직역이 안되는 영어 표현 — 문맥 속 의역으로 접근하는 훈련 강조.",
    "[학습 포인트] 3줄 도식화를 통한 전체 주제 파악 — 도입/전개/결론 구조 의식.",
    "[학습 포인트] 주제 배열 영작 서술형 — 한국어 해석이 없는 주제 영작 시 핵심 키워드 먼저 정리.",
    "[학습 포인트] 요약문 영작 — 원문의 군더더기를 빼고 핵심 메시지만 간결하게 표현.",
    "[학습 포인트] 지문 속 추상명사 — 구체적 예시로 의미를 풀어 이해하는 습관.",
    "[학습 포인트] 관계대명사절 — 선행사부터 다시 짚어가며 정확히 해석.",
    "[학습 포인트] 도치·강조 구문 — 평서문으로 재배열 후 의미 확인.",
]
CLINIC_STUDENTS_BY_WEEKDAY = {
    1: ["정주원", "유한선"],                          # 화요일 클리닉
    3: ["박서원", "장우영", "오우진", "이성우"],       # 목요일 클리닉 (이성우 추가)
    5: ["박도환"],                                    # 토요일 클리닉
}
_BOOTSTRAP_MARKER = os.path.join(DATA_DIR, ".clinic_bootstrap_v1.json")
_BOOTSTRAP_V2_MARKER = os.path.join(DATA_DIR, ".clinic_bootstrap_v2_sat.json")
_BOOTSTRAP_V3_THU_MARKER = os.path.join(DATA_DIR, ".clinic_bootstrap_v3_thu.json")
_CLEANUP_FUTURE_MARKER = os.path.join(DATA_DIR, ".clinic_cleanup_future_v1.json")


def _find_or_create_clinic_student(name, students):
    """이름으로 학생을 찾고, 없으면 CLI### 코드로 신규 생성."""
    for code, s in students.items():
        if s.get("name") == name:
            return code
    nums = []
    for c in students:
        if c.startswith("CLI"):
            try:
                nums.append(int(c[3:]))
            except ValueError:
                pass
    new_code = f"CLI{(max(nums) + 1 if nums else 1):03d}"
    students[new_code] = {"name": name, "pin": "1234", "parent": ""}
    return new_code


def add_clinic_records(clinic_date_str, student_names, extras_by_name=None):
    """클리닉 기록을 추가/업데이트.
    - 기존 기록 없으면 새로 추가
    - 기존 기록 있으면: extra가 있으면 [추가] 행으로 누적, 없으면 건너뜀
    반환: (added, updated, skipped) 학생명 리스트
    """
    extras_by_name = extras_by_name or {}
    data = load_data()
    students = dict(data["students"])
    records = list(data["records"])

    # 기존 클리닉 기록을 (날짜, 학생코드) → records 인덱스로 매핑
    existing_idx = {}
    for i, r in enumerate(records):
        if (r.get("date") == clinic_date_str
                and r.get("category") == CLINIC_CATEGORY):
            existing_idx[r.get("student_code")] = i

    added_names = []
    updated_names = []
    skipped_names = []

    for name in student_names:
        code = _find_or_create_clinic_student(name, students)
        extra = (extras_by_name.get(name, "") or "").strip()

        if code in existing_idx:
            # 이미 존재 — extra 있으면 추가, 없으면 건너뜀
            if extra:
                idx = existing_idx[code]
                cur_fb = records[idx].get("feedback", "")
                line = f"[추가] {extra}"
                if line not in cur_fb:
                    new_fb = cur_fb.rstrip() + ("\n" if cur_fb else "") + line
                    records[idx]["feedback"] = new_fb
                    updated_names.append(name)
                else:
                    skipped_names.append(name)
            else:
                skipped_names.append(name)
        else:
            # 신규 기록
            tip = random.choice(CLINIC_TIPS)
            feedback = f"{CLINIC_BASE_FEEDBACK} {tip}"
            if extra:
                feedback += f"\n[추가] {extra}"
            records.append({
                "date": clinic_date_str,
                "student_code": code,
                "category": CLINIC_CATEGORY,
                "score": "완료",
                "feedback": feedback,
                "flag": "",
                "resolved": False,
            })
            added_names.append(name)

    if added_names or updated_names:
        save_data(students, records)
    return added_names, updated_names, skipped_names


def _this_week_weekday_date(weekday_num):
    """이번주 (월~일 기준) weekday(0=Mon, 1=Tue, 3=Thu)의 날짜 문자열 반환."""
    today = datetime.now()
    diff = weekday_num - today.weekday()
    target = today + timedelta(days=diff)
    return target.strftime("%Y-%m-%d")


def _last_weekday_date(weekday_num):
    """오늘 또는 가장 가까운 과거 weekday의 날짜 문자열 반환."""
    today = datetime.now()
    diff = (today.weekday() - weekday_num) % 7
    target = today - timedelta(days=diff)
    return target.strftime("%Y-%m-%d")


def scheduled_tuesday_clinic():
    date_str = _last_weekday_date(1)
    add_clinic_records(date_str, CLINIC_STUDENTS_BY_WEEKDAY[1])


def scheduled_thursday_clinic():
    date_str = _last_weekday_date(3)
    add_clinic_records(date_str, CLINIC_STUDENTS_BY_WEEKDAY[3])


def scheduled_saturday_clinic():
    date_str = _last_weekday_date(5)
    add_clinic_records(date_str, CLINIC_STUDENTS_BY_WEEKDAY[5])


_scheduler_started = False

def start_clinic_scheduler():
    """APScheduler를 한 번만 시작."""
    global _scheduler_started
    if _scheduler_started:
        return
    try:
        from apscheduler.schedulers.background import BackgroundScheduler
        from apscheduler.triggers.cron import CronTrigger
        sched = BackgroundScheduler(timezone="Asia/Seoul")
        sched.add_job(scheduled_tuesday_clinic,
                      CronTrigger(day_of_week="tue", hour=23, minute=30),
                      id="tue_clinic", replace_existing=True)
        sched.add_job(scheduled_thursday_clinic,
                      CronTrigger(day_of_week="thu", hour=23, minute=30),
                      id="thu_clinic", replace_existing=True)
        sched.add_job(scheduled_saturday_clinic,
                      CronTrigger(day_of_week="sat", hour=13, minute=30),
                      id="sat_clinic", replace_existing=True)
        sched.start()
        _scheduler_started = True
    except Exception as e:
        # 스케줄러 실패해도 앱은 계속 동작
        print(f"[WARN] Scheduler start failed: {e}")


def run_one_time_bootstrap():
    """첫 배포 시 이번주 화요일 클리닉 기록 자동 추가 (유한선: 어법성판단 추가).
    한 번만 실행되도록 마커 파일로 관리.
    """
    import json
    if os.path.exists(_BOOTSTRAP_MARKER):
        return
    try:
        tue_date = _this_week_weekday_date(1)  # 이번주 화요일
        extras = {"유한선": "어법성판단 문제 오답 풀이 추가 진행."}
        added, _, _ = add_clinic_records(tue_date, CLINIC_STUDENTS_BY_WEEKDAY[1], extras)
        os.makedirs(os.path.dirname(_BOOTSTRAP_MARKER), exist_ok=True)
        with open(_BOOTSTRAP_MARKER, "w", encoding="utf-8") as f:
            json.dump({
                "ran_at": datetime.now().isoformat(),
                "tue_date": tue_date,
                "added_students": added,
            }, f, ensure_ascii=False)
    except Exception as e:
        print(f"[WARN] Bootstrap failed: {e}")


def run_one_time_bootstrap_v3_thursday():
    """이번주 목요일 클리닉 기록 자동 추가 (배포 1회만).
    오늘이 목요일이면 오늘 날짜로, 이미 지난주이면 가장 가까운 과거 목요일로.
    """
    import json
    if os.path.exists(_BOOTSTRAP_V3_THU_MARKER):
        return
    try:
        # 미래 날짜 방지: 이번주 목요일이 오늘 이후면 가장 가까운 과거 목요일 사용
        today = datetime.now()
        diff = 3 - today.weekday()
        if diff > 0:
            # 이번주 목요일이 미래 → 지난주 목요일 (또는 패스)
            thu_date = (today + timedelta(days=diff - 7)).strftime("%Y-%m-%d")
        else:
            thu_date = (today + timedelta(days=diff)).strftime("%Y-%m-%d")
        names = list(CLINIC_STUDENTS_BY_WEEKDAY[3])
        added, _, _ = add_clinic_records(thu_date, names)
        os.makedirs(os.path.dirname(_BOOTSTRAP_V3_THU_MARKER), exist_ok=True)
        with open(_BOOTSTRAP_V3_THU_MARKER, "w", encoding="utf-8") as f:
            json.dump({
                "ran_at": datetime.now().isoformat(),
                "thu_date": thu_date,
                "added_students": added,
            }, f, ensure_ascii=False)
    except Exception as e:
        print(f"[WARN] Bootstrap v3 (Thursday) failed: {e}")


def run_cleanup_future_clinic_records():
    """미래 날짜로 잘못 추가된 클리닉 기록을 일괄 삭제.
    이전 토요일 부트스트랩이 만든 미래 기록을 정리하기 위한 1회성 작업.
    """
    import json
    if os.path.exists(_CLEANUP_FUTURE_MARKER):
        return
    try:
        data = load_data()
        today_str = datetime.now().strftime("%Y-%m-%d")
        records = list(data["records"])
        before = len(records)
        records = [
            r for r in records
            if not (r.get("category") == CLINIC_CATEGORY
                    and r.get("date", "") > today_str)
        ]
        removed = before - len(records)
        if removed > 0:
            save_data(data["students"], records)
        os.makedirs(os.path.dirname(_CLEANUP_FUTURE_MARKER), exist_ok=True)
        with open(_CLEANUP_FUTURE_MARKER, "w", encoding="utf-8") as f:
            json.dump({
                "ran_at": datetime.now().isoformat(),
                "removed_count": removed,
            }, f, ensure_ascii=False)
    except Exception as e:
        print(f"[WARN] cleanup failed: {e}")


# 모듈 import 시점에 부트스트랩 + 스케줄러 가동
# 순서: (1) 화요일 기록 — 이전 배포에서 이미 실행됨
#       (2) 미래 클리닉 기록 정리 — 토요일 부트스트랩이 잘못 만든 기록 제거
#       (3) 이번주 목요일 기록 — 이성우 포함
#       (4) 스케줄러 가동 — 화 23:30 / 목 23:30 / 토 13:30
run_one_time_bootstrap()
run_cleanup_future_clinic_records()
run_one_time_bootstrap_v3_thursday()
start_clinic_scheduler()


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

    # 이 학생에게 보여줄 한마디: 게시 중 + (공통 OR 본인 개별)
    my_messages = [
        m for m in data["messages"]
        if m.get("published", True)
        and (not m.get("student_code") or m.get("student_code") == code)
    ]

    # 게시 중인 숙제만 (가장 최근 게시본 사용)
    published_homeworks = [h["content"] for h in data["homeworks"] if h.get("published", True)]
    homework = "\n\n".join(published_homeworks) if published_homeworks else ""

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

    # 미처리 특이사항 (재시/숙제미비/추가과제 등)
    my_flags = [r for r in items if r.get("flag") and not r.get("resolved")]
    my_flags.sort(key=lambda r: r["date"], reverse=True)

    return render_template(
        "student.html",
        student=student,
        records=items,
        grouped=grouped,
        stats=stats,
        class_avg=class_avg,
        charts=charts,
        homework=homework,
        messages=my_messages,
        by_date=by_date,
        ranks=ranks,
        my_flags=my_flags,
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
                    new_students, new_records, new_homeworks, new_messages = _parse_uploaded_excel(f)
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

                    # 기록 중복 판별
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

                    # 숙제 중복 판별: content 동일
                    seen_hws = {h.get("content", "") for h in current["homeworks"]}
                    unique_hws = []
                    skipped_hws = 0
                    for h in new_homeworks:
                        if h.get("content", "") in seen_hws:
                            skipped_hws += 1
                        else:
                            seen_hws.add(h.get("content", ""))
                            unique_hws.append(h)
                    merged_homeworks = current["homeworks"] + unique_hws

                    # 한마디 중복 판별: (학생코드, 내용) 일치
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

                    save_data(merged_students, merged_records, merged_homeworks, merged_messages)

                    parts = []
                    parts.append(f"신규 학생 {added_students}명" +
                                 (f" (중복 {skipped_students}명 건너뜀)" if skipped_students else ""))
                    parts.append(f"새 기록 {len(unique_records)}건" +
                                 (f" (중복 {skipped_records}건 건너뜀)" if skipped_records else ""))
                    parts.append(f"새 숙제 {len(unique_hws)}건" +
                                 (f" (중복 {skipped_hws}건 건너뜀)" if skipped_hws else ""))
                    parts.append(f"새 한마디 {len(unique_msgs)}건" +
                                 (f" (중복 {skipped_msgs}건 건너뜀)" if skipped_msgs else ""))
                    flash("추가 완료: " + " · ".join(parts), "success")
                else:
                    save_data(new_students, new_records, new_homeworks, new_messages)
                    flash(
                        f"덮어쓰기 완료: 학생 {len(new_students)}명, 기록 {len(new_records)}건, "
                        f"숙제 {len(new_homeworks)}건, 한마디 {len(new_messages)}건으로 전체 교체됨.",
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
        action = request.form.get("action", "add")

        if action == "bulk_delete":
            # 선택된 인덱스들을 받아서 일괄 삭제
            raw_indices = request.form.getlist("selected")
            try:
                indices = sorted({int(i) for i in raw_indices}, reverse=True)
            except (ValueError, TypeError):
                indices = []

            if not indices:
                flash("삭제할 기록을 선택해주세요.", "error")
            else:
                removed = 0
                for i in indices:
                    if 0 <= i < len(data["records"]):
                        del data["records"][i]
                        removed += 1
                save_data(data["students"], data["records"])
                flash(f"{removed}건의 기록이 일괄 삭제되었습니다.", "success")
            return redirect(url_for("admin_records",
                                    student=request.args.get("student", ""),
                                    category=request.args.get("category", "")))

        if action == "bulk_save":
            # 인라인 편집된 점수/피드백/비고/완료를 일괄 저장
            changed = 0
            for i in range(len(data["records"])):
                # 폼에 present_{i}=1 마커가 없으면 (필터로 가려진 행) 건너뜀
                if request.form.get(f"present_{i}") != "1":
                    continue
                r = data["records"][i]
                new_score = request.form.get(f"score_{i}", "").strip()
                new_feedback = request.form.get(f"feedback_{i}", "").strip()
                new_flag = request.form.get(f"flag_{i}", "").strip()
                new_resolved = request.form.get(f"resolved_{i}") == "on"
                if (new_score != r.get("score", "") or
                    new_feedback != r.get("feedback", "") or
                    new_flag != r.get("flag", "") or
                    new_resolved != bool(r.get("resolved"))):
                    r["score"] = new_score
                    r["feedback"] = new_feedback
                    r["flag"] = new_flag
                    r["resolved"] = new_resolved
                    changed += 1
            if changed > 0:
                save_data(data["students"], data["records"])
                flash(f"{changed}건의 기록이 일괄 수정되었습니다.", "success")
            else:
                flash("변경된 내용이 없습니다.", "success")
            return redirect(url_for("admin_records",
                                    student=request.args.get("student", ""),
                                    category=request.args.get("category", "")))

        # 기본: 새 기록 추가
        date = request.form.get("date", "").strip()
        student_code = request.form.get("student_code", "").strip()
        category = request.form.get("category", "").strip()
        score = request.form.get("score", "").strip()
        feedback = request.form.get("feedback", "").strip()
        flag = request.form.get("flag", "").strip()
        resolved = request.form.get("resolved") == "on"

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
                "flag": flag,
                "resolved": resolved,
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
                "flag": request.form.get("flag", "").strip(),
                "resolved": request.form.get("resolved") == "on",
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


# ─── 라우트: 관리자 — 클리닉 관리 ────────────────────────
@app.route("/admin/clinic", methods=["GET", "POST"])
@admin_required
def admin_clinic():
    if request.method == "POST":
        action = request.form.get("action", "")
        weekday_map = {"run_tuesday": 1, "run_thursday": 3, "run_saturday": 5}
        weekday = weekday_map.get(action)
        if weekday is None:
            flash("올바르지 않은 요청입니다.", "error")
            return redirect(url_for("admin_clinic"))

        date_str = _this_week_weekday_date(weekday)
        names = list(CLINIC_STUDENTS_BY_WEEKDAY[weekday])

        # 이번 회차 한정 추가 학생 (쉼표/공백으로 구분)
        extra_names_raw = request.form.get("extra_students", "").strip()
        if extra_names_raw:
            for n in extra_names_raw.replace(",", " ").split():
                n = n.strip()
                if n and n not in names:
                    names.append(n)

        # 학생별 추가 메모
        extras = {}
        for n in names:
            note = request.form.get(f"extra_{n}", "").strip()
            if note:
                extras[n] = note

        added, updated, skipped = add_clinic_records(date_str, names, extras)
        msg_parts = [f"{date_str} 클리닉 기록"]
        if added:
            msg_parts.append(f"신규 추가 {len(added)}명 ({', '.join(added)})")
        if updated:
            msg_parts.append(f"메모 추가됨 {len(updated)}명 ({', '.join(updated)})")
        if skipped:
            msg_parts.append(f"변경 없음 {len(skipped)}명")
        if not (added or updated or skipped):
            msg_parts.append("처리할 학생이 없습니다")
        flash(" · ".join(msg_parts), "success")
        return redirect(url_for("admin_clinic"))

    # 최근 클리닉 기록 확인
    data = load_data()
    recent_clinic = sorted(
        [r for r in data["records"] if r.get("category") == CLINIC_CATEGORY],
        key=lambda r: r.get("date", ""),
        reverse=True,
    )[:20]

    return render_template(
        "admin_clinic.html",
        tue_students=CLINIC_STUDENTS_BY_WEEKDAY[1],
        thu_students=CLINIC_STUDENTS_BY_WEEKDAY[3],
        sat_students=CLINIC_STUDENTS_BY_WEEKDAY[5],
        this_tue=_this_week_weekday_date(1),
        this_thu=_this_week_weekday_date(3),
        this_sat=_this_week_weekday_date(5),
        recent=recent_clinic,
        students=data["students"],
        scheduler_active=_scheduler_started,
    )


# ─── 라우트: 관리자 — 시험 등수 모아보기 ──────────────────
@app.route("/admin/rankings")
@admin_required
def admin_rankings():
    data = load_data()

    # 항목별로 학생별 점수 모으기 (숫자 점수만)
    cat_scores = defaultdict(lambda: defaultdict(list))  # cat -> {code -> [scores]}
    for r in data["records"]:
        sc = _parse_score(r["score"])
        if sc is None:
            continue
        cat = r["category"] or "기타"
        cat_scores[cat][r["student_code"]].append(sc)

    # 항목별 등수 계산 (경쟁식: 1, 2, 2, 4)
    rankings = {}
    for cat, code_scores in cat_scores.items():
        entries = []
        for code, scores in code_scores.items():
            entries.append({
                "code": code,
                "name": data["students"].get(code, {}).get("name", "?"),
                "best": max(scores),
                "avg": round(sum(scores) / len(scores), 1),
                "count": len(scores),
            })
        # 최고점 기준 내림차순
        entries.sort(key=lambda e: (-e["best"], -e["avg"], e["name"]))
        prev_best = None
        current_rank = 0
        for i, e in enumerate(entries):
            if e["best"] != prev_best:
                current_rank = i + 1
                prev_best = e["best"]
            e["rank"] = current_rank
        rankings[cat] = entries

    # 항목명 알파벳 정렬 (한국어 가나다)
    sorted_cats = sorted(rankings.keys())

    # 종합 순위: 메달 개수 기준 (시험마다 만점 다르므로 평균 의미 없음)
    # 금=1등, 은=2등, 동=3등
    medal_counts = defaultdict(lambda: {"gold": 0, "silver": 0, "bronze": 0, "categories": 0})
    for cat, entries in rankings.items():
        for e in entries:
            code = e["code"]
            medal_counts[code]["categories"] += 1
            if e["rank"] == 1:
                medal_counts[code]["gold"] += 1
            elif e["rank"] == 2:
                medal_counts[code]["silver"] += 1
            elif e["rank"] == 3:
                medal_counts[code]["bronze"] += 1

    overall = []
    for code, counts in medal_counts.items():
        overall.append({
            "code": code,
            "name": data["students"].get(code, {}).get("name", "?"),
            "gold": counts["gold"],
            "silver": counts["silver"],
            "bronze": counts["bronze"],
            "total_medals": counts["gold"] + counts["silver"] + counts["bronze"],
            "category_count": counts["categories"],
        })

    # 정렬: 금 많은 순 → 은 많은 순 → 동 많은 순 → 이름
    overall.sort(key=lambda x: (-x["gold"], -x["silver"], -x["bronze"], x["name"]))

    # 경쟁식 등수 (금·은·동 개수 모두 같으면 같은 등수)
    prev_key = None
    current_rank = 0
    for i, e in enumerate(overall):
        key = (e["gold"], e["silver"], e["bronze"])
        if key != prev_key:
            current_rank = i + 1
            prev_key = key
        e["rank"] = current_rank

    return render_template(
        "admin_rankings.html",
        rankings=rankings,
        categories=sorted_cats,
        overall=overall,
        total_students=len(data["students"]),
    )


# ─── 라우트: 관리자 — 특이사항 (재시/숙제미비/추가과제) ────
@app.route("/admin/flags", methods=["GET", "POST"])
@admin_required
def admin_flags():
    data = load_data()

    if request.method == "POST":
        action = request.form.get("action", "")
        try:
            idx = int(request.form.get("idx", -1))
        except (ValueError, TypeError):
            idx = -1

        if 0 <= idx < len(data["records"]):
            if action == "resolve":
                data["records"][idx]["resolved"] = True
                save_data(data["students"], data["records"])
                flash("완료 처리되었습니다.", "success")
            elif action == "unresolve":
                data["records"][idx]["resolved"] = False
                save_data(data["students"], data["records"])
                flash("미처리로 되돌렸습니다.", "success")
            elif action == "clear_flag":
                data["records"][idx]["flag"] = ""
                data["records"][idx]["resolved"] = False
                save_data(data["students"], data["records"])
                flash("비고가 제거되었습니다.", "success")
        return redirect(url_for("admin_flags"))

    # 미처리 특이사항: flag 있고 resolved 아닌 것
    active = [(i, r) for i, r in enumerate(data["records"])
              if r.get("flag") and not r.get("resolved")]
    active_by_date = defaultdict(list)
    for i, r in active:
        active_by_date[r["date"]].append((i, r))
    active_list = sorted(active_by_date.items(), key=lambda x: x[0], reverse=True)

    # 처리 완료 이력 (최근 30건)
    resolved_recs = [(i, r) for i, r in enumerate(data["records"])
                     if r.get("flag") and r.get("resolved")]
    resolved_recs.sort(key=lambda x: x[1]["date"], reverse=True)
    resolved_recs = resolved_recs[:30]

    # 플래그 종류별 카운트
    flag_counts = defaultdict(int)
    for _, r in active:
        flag_counts[r["flag"]] += 1

    return render_template(
        "admin_flags.html",
        active_by_date=active_list,
        resolved_recs=resolved_recs,
        students=data["students"],
        total_active=len(active),
        flag_counts=dict(flag_counts),
    )


# ─── 라우트: 관리자 — 공지사항(이번주 숙제 + 한마디) ────────
@app.route("/admin/announcements", methods=["GET", "POST"])
@admin_required
def admin_announcements():
    data = load_data()

    if request.method == "POST":
        action = request.form.get("action", "")

        # ─── 이번주 숙제 ────────────────────────────────
        if action == "add_homework":
            content = request.form.get("homework", "").strip()
            if not content:
                flash("숙제 내용을 입력해주세요.", "error")
            else:
                new_hws = list(data["homeworks"]) + [{"content": content, "published": True}]
                save_data(homeworks=new_hws)
                flash("새 숙제가 게시되었습니다.", "success")

        elif action in ("publish_homework", "unpublish_homework", "delete_homework"):
            try:
                idx = int(request.form.get("idx", -1))
            except ValueError:
                idx = -1
            if 0 <= idx < len(data["homeworks"]):
                new_hws = list(data["homeworks"])
                if action == "delete_homework":
                    del new_hws[idx]
                    flash("숙제가 삭제되었습니다.", "success")
                elif action == "publish_homework":
                    new_hws[idx] = dict(new_hws[idx], published=True)
                    flash("숙제가 다시 게시되었습니다.", "success")
                else:
                    new_hws[idx] = dict(new_hws[idx], published=False)
                    flash("숙제를 내렸습니다 (보관됨).", "success")
                save_data(homeworks=new_hws)
            else:
                flash("대상 숙제를 찾을 수 없습니다.", "error")

        # ─── 한마디 ─────────────────────────────────────
        elif action == "add_message":
            content = request.form.get("content", "").strip()
            student_code = request.form.get("student_code", "").strip()
            if not content:
                flash("한마디 내용을 입력해주세요.", "error")
            elif student_code and student_code not in data["students"]:
                flash("선택한 학생코드가 존재하지 않습니다.", "error")
            else:
                new_messages = list(data["messages"]) + [{
                    "student_code": student_code,
                    "content": content,
                    "published": True,
                }]
                save_data(messages=new_messages)
                target = data["students"][student_code]["name"] if student_code else "전체"
                flash(f"한마디가 게시되었습니다. (대상: {target})", "success")

        elif action in ("publish_message", "unpublish_message", "delete_message"):
            try:
                idx = int(request.form.get("idx", -1))
            except ValueError:
                idx = -1
            if 0 <= idx < len(data["messages"]):
                new_messages = list(data["messages"])
                if action == "delete_message":
                    del new_messages[idx]
                    flash("한마디가 삭제되었습니다.", "success")
                elif action == "publish_message":
                    new_messages[idx] = dict(new_messages[idx], published=True)
                    flash("한마디가 다시 게시되었습니다.", "success")
                else:
                    new_messages[idx] = dict(new_messages[idx], published=False)
                    flash("한마디를 내렸습니다 (보관됨).", "success")
                save_data(messages=new_messages)
            else:
                flash("대상 한마디를 찾을 수 없습니다.", "error")

        return redirect(url_for("admin_announcements"))

    # 분리 (인덱스 유지 위해 enumerate)
    hws_indexed = list(enumerate(data["homeworks"]))
    hws_published = [(i, h) for i, h in hws_indexed if h.get("published")]
    hws_archived = [(i, h) for i, h in hws_indexed if not h.get("published")]

    msgs_indexed = list(enumerate(data["messages"]))
    msgs_published = [(i, m) for i, m in msgs_indexed if m.get("published")]
    msgs_archived = [(i, m) for i, m in msgs_indexed if not m.get("published")]

    return render_template(
        "admin_announcements.html",
        hws_published=hws_published,
        hws_archived=hws_archived,
        msgs_published=msgs_published,
        msgs_archived=msgs_archived,
        students=data["students"],
    )


# ─── 라우트: Excel 다운로드/템플릿 ───────────────────────────
@app.route("/admin/download")
@admin_required
def admin_download():
    """현재 데이터를 Excel 파일로 다운로드.
    다운로드 직전에 최신 형식(비고/완료 컬럼 포함)으로 재저장하여 항상 신형식 보장.
    """
    if not os.path.exists(EXCEL_PATH):
        flash("아직 데이터가 없습니다. 빈 템플릿을 받으시려면 'Excel 템플릿 다운로드'를 이용하세요.", "error")
        return redirect(url_for("admin"))
    # 신형식 보장을 위해 재저장
    data = load_data()
    save_data(data["students"], data["records"], data["homeworks"], data["messages"])
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
    ws2.append(["날짜", "학생코드", "학생이름", "항목", "점수", "피드백", "비고", "완료"])
    ws2.append(["2026-05-18", "S001", "김민지", "Daily Test", "92",
                "어휘 문제에서 실수가 있었지만 독해는 완벽했습니다.", "", ""])
    ws2.append(["2026-05-18", "S001", "김민지", "숙제", "완료", "꼼꼼하게 잘 했습니다.", "", ""])
    ws2.append(["2026-05-18", "S002", "이도윤", "단어시험", "60",
                "기준 미달, 재시 필요.", "재시", ""])
    ws2.append(["2026-05-17", "S002", "이도윤", "숙제", "미제출",
                "다음 시간 꼭 챙겨오세요.", "숙제미비", "O"])

    ws3 = wb.create_sheet("공지사항")
    ws3.append(["종류", "대상학생코드", "학생이름", "내용", "상태"])
    ws3.append(["이번주숙제", "", "", "워크북 32-45쪽 풀고, 단어 50개 외워오기", "게시"])
    ws3.append(["이번주숙제", "", "", "지난주 숙제: 모의고사 1회 풀어오기", "내림"])
    ws3.append(["한마디", "", "", "시험 기간 화이팅! 모두 잘 할 수 있을거예요.", "게시"])
    ws3.append(["한마디", "S001", "김민지", "단어시험 1등 축하해요. 다음주도 기대할게요!", "게시"])

    for ws, widths in [
        (ws1, [10, 12, 8, 18]),
        (ws2, [12, 10, 12, 14, 8, 40, 12, 8]),
        (ws3, [14, 14, 12, 50, 10]),
    ]:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + i)].width = w

    _add_records_dropdowns(ws2)
    _add_announcements_dropdown(ws3)

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
