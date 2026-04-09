# EasyPOS 자동전표/매출양식 자동입력 프로그램 소개서 (Codex 재구현용)

## 1) 프로그램 목적
이 프로그램은 서울역(및 확장 가능한 사업장) POS 데이터를 받아 다음 2개 산출물을 생성한다.

1. 자동전표 업로드용 엑셀
2. 당사 매출양식 자동입력 엑셀 (맨 앞에 `(회계일자) 검증` 시트 포함)

추가로, 생성된 전표와 정산주문목록의 차이를 검증하는 기능을 제공한다.

## 2) 기술 스택/구조
- 언어: Python 3.11+
- 웹: FastAPI + Jinja2
- 엑셀 처리: openpyxl
- 주요 파일:
`app/main.py`: 웹 엔드포인트/업로드/다운로드
`app/generator.py`: 전표/매출양식 생성 로직
`app/parsers.py`: 카드/당일/정산/전표 파서
`app/config.py`: 설정 모델 로딩
`config/voucher_config.json`: 사업장/계정/결제수단/매장 매핑
`config/store_name_mapping.json`: 정산 가맹점 별칭 매핑
`templates/index.html`: UI

## 3) 입력/출력 사양
### 입력
- 필수: 카드매출 파일(`.xlsx`), 매장별 당일 매출내역 파일(`.xlsx`)
- 선택: 정산주문목록 파일(`.xlsx`), 전표 템플릿(`자동전표 양식.xlsx`), 매출양식 템플릿(`*임대을매출*POS매출*.xlsx`), 회계일자

### 출력
- 자동전표 파일: `{사업장명}POS매장임대을매출_{YYYYMMDD}.xlsx`
- 매출양식 파일: `{매출양식파일명}_{YYYYMMDD}_자동입력.xlsx`

## 4) 핵심 엔드포인트
### `POST /generate` (통합 생성)
한 번의 업로드로 아래 2개를 동시에 생성한다.
1. 자동전표
2. 매출양식 자동입력

폼 필드명:
- `business_key`
- `account_date_input`
- `card_sales_file`
- `daily_sales_file`
- `settlement_order_file` (선택)
- `template_file` (전표 템플릿, 선택)
- `sales_template_file` (매출양식 템플릿, 선택)

### `POST /generate-sales-input` (매출양식 단독 생성)
폼 필드명:
- `business_key`
- `sales_account_date_input`
- `card_sales_file_sales`
- `daily_sales_file_sales`
- `settlement_order_file_sales` (선택)
- `sales_template_file` (선택)

### `POST /verify-settlement`
생성된 전표와 정산주문목록 비교 검증.

### 다운로드
- 전표: `GET /download/{result_id}`
- 매출양식: `GET /download-sales-input/{result_id}`

## 5) 파싱 규칙 (정확히 재현할 것)
### 카드매출 파싱
필수 헤더: `가맹점명`, `구분`, `승인일자`, `카드사명`, `승인금액`

규칙:
1. `구분 == 승인`만 집계
2. `승인번호` 컬럼이 있으면, `구분 == 취소`의 승인번호를 먼저 수집
3. 집계 시 취소 승인번호와 같은 승인건은 제외
4. 카드사명은 `payment_methods[].source_card_name`으로 결제키 매핑

### 당일매출 파싱
신양식 헤더: `매장명`, `영업일자`, `총매출`, `할인`, `현금매출`

규칙:
1. 파일명/행 값에서 회계일 추정
2. 같은 매장/회계일은 합산
3. 신양식은 전자화폐 값이 없으므로 0으로 처리

구양식(레거시)은 고정 컬럼 기반으로 파싱:
- C: 매장명, E: 총매출, H: 할인, K: 현금매출, T: 전자화폐

### 정산주문목록 파싱
필수 헤더: `가맹점`, `결제금액`

규칙:
1. `결제 승인일`이 있으면 회계일 필터에 활용
2. `store_name_mapping`과 정규화 비교로 가맹점-매장 매핑
3. 부분문자열 보조 매칭 지원
4. 매장별 정산금액 합산

## 6) 자동전표 생성 규칙
전표 라인 생성 핵심:
- `gross_sales(401510)`
- `discount(401511)`는 음수
- `cash(103300)`
- `receivable(102700)` 카드사별 금액
- 정산주문목록 금액은 추가 `receivable(102700)` 라인으로 반영

계산식:
- `receivable_total = 카드사합 + 정산금액`
- `debit_total = 현금 + receivable_total`
- `gross_sales_amount = debit_total + 할인(abs)`

## 7) 매출양식 자동입력 규칙
### 시트/매장 매핑
- `business.active_stores` 기준 반복
- 시트명 매칭: `output_name`, `source_name`, 서울역점 제거명, 정규화 비교, 부분문자열
- 매칭 실패 매장은 스킵하고 notes에 기록

### 입력 범위
- 일자 행: A열의 날짜가 회계일과 같은 행(탐색 범위 4~34)
- 수정 컬럼:
`B` 현금
`C` 전자화폐
`E` 할인(abs)
`M~U` 카드사별 금액
- 36행 이후 절대 수정 금지

### 정산 보정(매출양식)
- 정산 결제수단 키: `settlement_order.partner_payment_key` (기본 `common`)
- 해당 결제수단의 `daily_field` 기준으로 보정 대상 결정
- 현재 핵심 요구사항: **가산이 아닌 대체**

대체 규칙:
- `daily_field`가 전자화폐 계열이면 `C열 = 정산금액`으로 대체
- `daily_field`가 현금 계열이면 `B열 = 정산금액`으로 대체
- 검증시트 총매출값도 동일한 대체 기준으로 재계산
  - 전자화폐 대체: `총매출 = 기존총매출 - 기존전자화폐 + 정산금액`
  - 현금 대체: `총매출 = 기존총매출 - 기존현금 + 정산금액`

## 8) 검증 시트 규칙
- 시트명: `({YYYYMMDD}) 검증`
- 결과 파일의 맨 앞(index 0)에 생성
- 가능하면 `검증시트.xlsx` 템플릿 첫 시트를 스타일/서식/수식 포함 복제
- 기본 데이터 열:
`A 매장명`
`B 현금`
`C 전자화폐`
`D 카드매출합`
`E 할인`
`F 총매출`
`M~U 카드사`
`V = SUM(M:U)`
`W = 카드매출파일 합계`
`X = V=W`

## 9) 설정 파일 핵심 (재구현 필수)
`config/voucher_config.json`:
- `businesses`: 사업장별 전표 코드, 계정코드, 매장목록
- `payment_methods`: 카드사/정산 결제수단 정의
- `settlement_order.partner_payment_key`: 정산에 쓸 결제수단 키
- `settlement_order.management_data`: 정산 전표 관리항목

`config/store_name_mapping.json`:
- 정산 가맹점 별칭 매핑(가맹점명 흔들림 대응)

## 10) 예외/오류 처리 정책
- 회계일 자동 추출 실패 시 사용자 입력값 요구
- 카드/당일 회계일 불일치면 오류
- 템플릿 미존재 시 업로드 요청 오류
- 매출양식에서 채울 매장이 0개면 오류

## 11) 새 Codex 채팅에 전달할 재구현 지시 템플릿
아래를 새 채팅에 그대로 전달하면 된다.

"""
FastAPI 기반으로 POS 자동화 툴을 재구현해줘.
필수 기능은 다음과 같아:
1) 카드매출/당일매출/정산/회계일자를 받아 자동전표 파일 생성
2) 같은 입력으로 당사 매출양식 템플릿에도 자동입력 파일 생성
3) 매출양식 결과물 맨 앞에 `(회계일자) 검증` 시트 생성
4) 카드 취소건은 승인번호 기준으로 제외
5) 정산 보정은 매출양식에서 가산이 아니라 대체(B 또는 C)로 처리
6) UI는 /generate(통합), /generate-sales-input(단독), /verify-settlement 제공
7) 설정은 config/voucher_config.json + config/store_name_mapping.json 사용
8) 36행 이후 미수정, 시트 매핑 실패 매장은 notes에 기록

구현 파일 구조:
- app/main.py, app/generator.py, app/parsers.py, app/config.py, templates/index.html
- 출력 다운로드 엔드포인트 /download, /download-sales-input
"""
