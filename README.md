# 학생 성적·피드백 공유 포털

선생님이 로컬 Excel 파일로 입력한 학생별 Daily Test, 숙제, 피드백을
학부모가 **자녀 본인의 기록만** 웹에서 열람할 수 있는 프로그램입니다.

- 👨‍🏫 **선생님**: 로컬 Excel 작성 → 관리자 페이지에서 업로드 (1초)
- 👨‍👩‍👧 **학부모**: 고정된 사이트 주소 + 학생코드 + PIN으로 본인 자녀 기록 열람
- 🔒 다른 학생 기록은 절대 볼 수 없음 (서버 단에서 차단)

---

## 1. 로컬에서 먼저 테스트하기

### 1-1) Python 설치
[python.org](https://www.python.org/downloads/) 에서 Python 3.10 이상 설치.
설치할 때 **"Add Python to PATH"** 체크박스 꼭 켜세요.

### 1-2) 폴더 이동 + 패키지 설치 (PowerShell)
```powershell
cd C:\Users\herna\student-portal
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### 1-3) 환경변수 파일 만들기
```powershell
copy .env.example .env
notepad .env
```
`ADMIN_PASSWORD` 값을 본인이 기억할 비밀번호로 변경해서 저장.

### 1-4) 실행
```powershell
python app.py
```
브라우저에서 `http://localhost:5000` 접속.

---

## 2. Excel 파일 작성

관리자 페이지(`/admin`) 접속 → "빈 Excel 템플릿 다운로드" 클릭하면
바로 사용 가능한 샘플 파일을 받을 수 있습니다.

### 시트 1: `학생명단`
| 학생코드 | 학생이름 | PIN  | 학부모이름(선택) |
|---------|---------|------|----------------|
| S001    | 김민지   | 1234 | 김민지 어머니   |
| S002    | 이도윤   | 5678 | 이도윤 어머니   |

- **학생코드**: 학부모에게 알려줄 ID. `S001`, `2026A01` 등 자유롭게.
- **PIN**: 4~6자리 숫자 권장. 비워두면 PIN 없이 학생코드만으로 접속.

### 시트 2: `기록`
| 날짜       | 학생코드 | 항목       | 점수 | 피드백                          |
|-----------|---------|-----------|------|-------------------------------|
| 2026-05-18| S001    | Daily Test| 92   | 어휘 문제에서 실수가 있었지만...  |
| 2026-05-18| S001    | 숙제      | 완료 | 꼼꼼하게 잘 했습니다.            |
| 2026-05-17| S001    | Daily Test| 88   | 시간 관리 연습 필요.            |

- **항목**: `Daily Test`, `숙제`, `단어시험` 등 자유롭게. 같은 항목끼리 자동 그룹핑.
- **점수**: 숫자(`92`) 또는 텍스트(`완료`, `A+`) 둘 다 OK. 숫자만 자동 평균 계산.
- 매일 새 행을 아래에 추가하면 됩니다.

작성한 Excel을 관리자 페이지에서 **업로드** 클릭 → 학부모 페이지에 즉시 반영.

---

## 3. 인터넷에 배포 (학부모에게 주소 안내용)

### Render 무료 배포 (추천, 가장 쉬움)

1. [github.com](https://github.com) 가입 → 새 repository 만들기 → 이 폴더 전체 업로드
   (또는 GitHub Desktop 으로 push)
2. [render.com](https://render.com) 가입 (GitHub 계정으로 가능)
3. Dashboard → **New +** → **Blueprint** 선택
4. GitHub repository 연결 → `render.yaml` 자동 감지 → Apply
5. **Environment** 탭에서 `ADMIN_PASSWORD` 값 입력
6. 배포 완료 후 받는 URL (`https://student-portal-xxxx.onrender.com`)을
   학부모에게 공유

> Render 무료 플랜은 15분 미접속시 절전 모드로 들어가서, 첫 접속이 30초쯤 걸릴 수 있습니다.
> 유료 $7/월 플랜은 항상 깨어있어요.

### 대안: Railway, Fly.io, PythonAnywhere
같은 코드로 다 배포 가능. `gunicorn app:app` 으로 시작하면 됩니다.

---

## 4. 학부모에게 안내할 내용 (예시)

> 안녕하세요. 자녀의 학습 기록을 아래 사이트에서 확인하실 수 있습니다.
>
> **사이트**: https://student-portal-xxxx.onrender.com
> **학생코드**: S001
> **PIN**: 1234
>
> 매일 저녁 업데이트됩니다.

---

## 5. 운영 팁

- **PIN 분실 학부모**: 관리자 페이지에 모든 학생의 PIN이 보입니다.
- **학생 추가/제거**: Excel `학생명단` 시트만 수정 후 재업로드.
- **데이터 백업**: 본인이 로컬에 보관하는 Excel 파일이 곧 백업입니다.
- **PIN 보안**: PIN은 단순한 1차 차단용입니다. 강한 보안이 필요하면
  복잡한 PIN(예: `K7m2`)을 사용하세요. 사이트 URL 자체를 비공개로 유지하는 게 가장 효과적.

---

## 6. 파일 구조

```
student-portal/
├── app.py              # Flask 서버
├── requirements.txt    # Python 패키지 목록
├── render.yaml         # Render 배포 설정
├── .env.example        # 환경변수 예시
├── .gitignore
├── data/               # Excel 파일 저장 위치
│   └── students.xlsx   # (업로드시 생성)
├── templates/          # HTML
│   ├── base.html
│   ├── login.html      # 학부모 로그인
│   ├── student.html    # 자녀 기록 페이지
│   ├── admin_login.html
│   └── admin.html      # 선생님 관리자 페이지
└── static/
    └── style.css
```

---

## 7. 자주 묻는 문제

**Q. 업로드한 Excel이 안 보여요.**
- 시트 이름이 정확히 `학생명단`, `기록` 인지 확인.
- 첫 줄은 헤더이고, 두 번째 줄부터 데이터.

**Q. 학부모가 다른 자녀 코드를 입력하면 어떻게 되나요?**
- 그 학생의 PIN을 모르면 로그인 자체가 안 됩니다.
- PIN을 안다고 해도 그 학생 기록만 보입니다. 다른 학생은 절대 보이지 않음.

**Q. Render 무료 플랜이 15분 후 자는 게 싫어요.**
- 유료 $7/월 플랜 사용, 또는 UptimeRobot 등으로 5분마다 핑.
