# POS Voucher Generator

`카드매출.xlsx`와 `매장별 당일 매출내역.xlsx`를 업로드해 자동전표 업로드용 엑셀을 생성하는 FastAPI 앱입니다.

## Local Run

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```

브라우저에서 `http://127.0.0.1:8000` 접속

## Current Rule Scope

- 사업장: `서울역`, `청량리역` (청량리역은 기본 골격만 추가되어 있으며 `stores` 매핑 입력 필요)
- 자동 생성 제외: `프레퍼스`, `밀본`
- 매핑/예외 규칙: `config/voucher_config.json`
- 회계일자: 파일 자동 추출 + 화면 수동 입력(`YYYYMMDD` 또는 `YYMMDD`) 지원
- 템플릿: 기본 파일(`자동전표 양식.xlsx`)이 프로젝트 루트에 없으면 화면에서 템플릿을 업로드해야 함
- 전표 계산 규칙(신규): `401511=-할인`, `103300=현금`, `102700=카드 승인금액(승인현황)`, `401510=잔액(차대합 일치)`
- 정산주문목록: 선택 업로드 시 `C열(가맹점)` + `F열(결제금액)`을 매장 매핑해 `102700` 추가 라인 생성
- 정산 검증: 생성된 전표 파일과 정산주문목록 파일을 별도 업로드하여 차이 검증 가능
- 가맹점 매핑표: `config/store_name_mapping.json` (파일 간 상이한 매장명 매핑, 직접 수정 가능)

## Files

- 엔트리포인트: `app/main.py`
- 생성 로직: `app/generator.py`
- 입력 파싱: `app/parsers.py`
- 설정 로더: `app/config.py`
- 웹 템플릿: `templates/index.html`
- Railway 설정: `railway.json`, `Procfile`

## Railway Deploy

1. 이 저장소를 GitHub에 push
2. Railway에서 New Project -> Deploy from GitHub Repo 선택
3. 서비스의 Start Command는 `uvicorn app.main:app --host 0.0.0.0 --port $PORT` (이미 `railway.json`/`Procfile`에 반영됨)
4. 배포 후 `/health` 경로가 200 응답인지 확인

## Config Editing

`config/voucher_config.json`에서 아래를 변경하면 코드 수정 없이 반영됩니다.

- 매장 매핑 (`stores`)
- 자동 포함/제외 (`enabled`)
- 수기 확인 메모 (`manual_review_reason`)
- 결제수단 매핑 (`payment_methods`)
