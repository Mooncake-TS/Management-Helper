# 테마상품 매입·매출 대시보드

ERP에서 내려받은 전년도 매출, 올해 매출, 매입 엑셀을 읽어 매입·매출 현황과 예측 발주량을 보여주는 Streamlit 대시보드입니다.

## 사용 방법

1. 저장소 최상위 폴더(`app.py` 옆) 또는 `data/`에 다음 엑셀 3개를 올립니다.
   - `전년도 판매.xlsx`
   - `금년도 판매_08.31.xlsx`
   - `매입_08.31.xlsx`
2. 배포된 대시보드를 열면 세 파일을 자동으로 읽고 분석합니다.
3. 왼쪽에서 연결된 파일명을 확인합니다. 파일명 뒤의 숫자는 데이터 기준일이며, 날짜가 바뀌어도 앞부분 이름으로 파일을 찾습니다.
4. 임시로 다른 파일을 사용하려면 왼쪽 `다른 파일로 분석하기`를 펼쳐 업로드합니다. 해당 파일만 저장된 파일보다 우선 사용합니다.

확장자는 `.xlsx`, `.xlsm`을 지원합니다. 기준일은 `08.31`처럼 같은 형식으로 유지하고, 갱신할 때 이전 파일을 교체하는 방식을 권장합니다. 같은 종류의 파일이 여러 개면 날짜 숫자가 큰 파일을 선택합니다. 파일 이름의 날짜는 표시와 파일 선택에만 사용하며, 실제 분석 날짜는 엑셀 내용을 따릅니다.

업로드한 원본 파일은 수정하지 않습니다. 앱이 실제 헤더를 자동으로 찾고 월별 합계·소계 행을 제외한 뒤, `tag_master.csv`를 기준으로 상품을 분류합니다.

## GitHub에 올릴 파일

```text
app.py
data_loader.py
forecasting.py
tag_master.csv
requirements.txt
README.md
.gitignore
.streamlit/config.toml
```

`run_dashboard.bat`은 Windows 로컬 실행용이므로 GitHub에 함께 올려도 되지만 클라우드 구동에는 필요하지 않습니다.

## 배포

1. 위 파일을 GitHub 저장소에 올립니다.
2. [Streamlit Community Cloud](https://share.streamlit.io/)에서 저장소를 연결합니다.
3. 앱 진입 파일로 `app.py`를 선택합니다.
4. 배포가 끝나면 생성된 `streamlit.app` 링크를 공유합니다.

원본 엑셀은 `.gitignore` 규칙으로 기본 제외됩니다. 웹에서 업로드하거나, Git CLI에서는 자동 연결할 파일만 명시적으로 `git add -f`로 추가합니다.

## 로컬 실행

Windows에서는 `run_dashboard.bat`을 실행하거나 아래 명령을 사용합니다.

```powershell
streamlit run app.py
```

앱 폴더와 `data/`에서 위 이름의 파일을 자동으로 찾습니다. `PREVIOUS_SALES_PATH`, `CURRENT_SALES_PATH`, `PURCHASE_PATH` 환경변수가 지정되면 해당 파일을 우선 사용하며, 기존 PC의 OneDrive 경로도 대체 경로로 유지합니다.

## 예측 발주 계산

```text
예상수요 = 작년 동일 기간 판매량 × (1 + 적용 성장률)
목표재고 = 예상수요 + 안전재고
추천 발주량 = 목표재고 - 현재고 - 입고예정수량
```

최종 추천량은 입력한 발주단위에 맞춰 올림 처리합니다.
