# 테마상품 매입·매출 대시보드

매출 `Rawdata` 시트와 매입 `Purchase_Rawdata` 시트를 읽어 매입·매출 현황과 예측 발주량을 보여주는 Streamlit 대시보드입니다.

## 사용 방법

1. 배포된 대시보드 링크를 엽니다.
2. 왼쪽 `매출 Summary`에 매출 엑셀 파일을 올립니다.
3. 왼쪽 `매입 Summary`에 매입 엑셀 파일을 올립니다.
4. 두 파일의 처리가 끝나면 전체 현황, 월 상세, 예측 발주, 데이터 상태 탭이 자동으로 표시됩니다.

업로드한 원본 파일은 수정하지 않으며, 합계·소계가 제거된 Rawdata 시트만 읽습니다.

## GitHub에 올릴 파일

```text
app.py
data_loader.py
forecasting.py
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

원본 엑셀 파일은 `.gitignore` 규칙으로 GitHub 업로드 대상에서 제외됩니다.

## 로컬 실행

Windows에서는 `run_dashboard.bat`을 실행하거나 아래 명령을 사용합니다.

```powershell
streamlit run app.py
```

현재 PC에서는 기존 OneDrive 파일을 자동으로 찾을 수 있으며, 다른 환경에서는 왼쪽 업로드 칸을 사용합니다.

## 예측 발주 계산

```text
예상수요 = 작년 동일 기간 판매량 × (1 + 적용 성장률)
목표재고 = 예상수요 + 안전재고
추천 발주량 = 목표재고 - 현재고 - 입고예정수량
```

최종 추천량은 입력한 발주단위에 맞춰 올림 처리합니다.
