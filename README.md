# CRQ Risk Calculator

사이버 리스크 정량화(Cyber Risk Quantification) 계산기. 업종·연매출·보안등급(SecurityScorecard) 3가지 입력만으로 연간 사이버 손실 분포를 Monte Carlo 시뮬레이션으로 계산합니다.

**Live**: https://crq-risk-calculator.onrender.com

## 구조

- `static/index.html` — 프론트엔드 (순수 HTML/CSS/JS, 빌드 과정 없음)
- `server.py` — 백엔드 HTTP 서버 (Python 표준 라이브러리 `http.server`만 사용)
- `crq_api.py` — REST API 레이어 (입력 검증, JSON 응답)
- `crq_monte_carlo.py` — Monte Carlo 계산 엔진
- `RiskSync_CRQ_DataWarehouse_LATEST.xlsx` — 업종별 빈도/심각도 등 원본 데이터 (IRIS, Verizon DBIR 기반)

## 로컬 실행

```bash
pip install -r requirements.txt
python3 server.py
# http://localhost:8765
```

## API

| Method | Path | 설명 |
|---|---|---|
| GET | `/api/options` | 업종/매출구간/보안등급 드롭다운 데이터 |
| POST | `/api/calculate` | `{industry, revenue, ssc_grade?}` → 계산 결과 |
| GET | `/api/methodology` | 근거 자료 (Assumption/Evidence Register 등) |
| POST | `/api/report/pdf` | 결과 리포트 PDF 다운로드 (headless Chrome 필요) |

## 배포

- **백엔드**: Render (`server.py`가 정적 파일 서빙 + API 모두 처리)
- 코드 push 시 Render가 자동으로 재배포함

### 알려진 제약

Render 환경에는 headless Chrome이 없어서 `/api/report/pdf`가 500을 반환합니다. 이 경우 프론트엔드가 자동으로 브라우저 인쇄 대화상자(Print → Save as PDF)로 대체합니다. 원클릭 PDF 다운로드는 Chrome이 설치된 로컬 환경에서만 동작합니다.
