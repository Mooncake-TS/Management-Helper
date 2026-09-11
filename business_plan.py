"""Read the monthly business-plan source and render independent plan analytics."""
from __future__ import annotations

from io import BytesIO
from pathlib import Path
import re
import unicodedata

import pandas as pd
from openpyxl import load_workbook


def find_plan_file(root: Path) -> Path | None:
    files = [p for folder in (root, root / "data") for p in folder.glob("*")
             if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm"}
             and unicodedata.normalize("NFC", p.stem).startswith("사업계획")
             and not p.name.startswith("~$")]
    def order(p):
        dates = tuple(map(int, re.findall(r"\d+", p.stem)))
        return dates, p.stat().st_mtime_ns, p.name
    return max(files, key=order) if files else None


def read_plan(content: bytes) -> tuple[pd.DataFrame, dict]:
    """Use monthly detail only. Ignore source ratio formulas and annual subtotals."""
    book = load_workbook(BytesIO(content), read_only=True, data_only=True)
    candidates = [s for s in book if "사업계획" in s.title]
    sheet = candidates[0] if candidates else book.worksheets[0]
    rows = list(sheet.values)
    header = next((i for i, r in enumerate(rows[:20])
                   if len(r) > 3 and "아쿠아샵" in str(r[2]) and r[3]), None)
    if header is None:
        book.close()
        raise ValueError("매장별 목표 표의 헤더를 찾지 못했습니다.")
    # C is the report total; Y is notes. Store columns are D through X in this template.
    stores = {c: re.sub(r"\s+", " ", str(rows[header][c])).strip()
              for c in range(3, min(24, len(rows[header]))) if rows[header][c]}
    records, month, mismatches, missing_cache = [], None, [], []
    for row_number, row in enumerate(rows[header + 1:], header + 2):
        marker = str(row[0] or "").strip()
        match = re.fullmatch(r"(\d{1,2})월", marker)
        if match:
            month = int(match[1])
        elif marker:
            month = None  # annual and secondary summaries must never be counted twice
        if month is None or not 1 <= month <= 12:
            continue
        label = re.fullmatch(r"(20\d{2})년\s*매출(\s*목표)?", str(row[1] or "").strip())
        if not label:
            continue
        year, kind = int(label[1]), "목표" if label[2] else "실적"
        values = []
        for col, store in stores.items():
            value = row[col] if col < len(row) else None
            if value is not None and not isinstance(value, (int, float)):
                raise ValueError(f"{sheet.title} {row_number}행 {store}: 숫자가 아닌 금액이 있습니다.")
            values.append(value)
            records.append({"연도": year, "월": month, "매장": store,
                            "구분": kind, "금액": float(value) if value is not None else 0.0})
        if not any(v is not None for v in values) and (kind == "목표" or row[2] is None):
            missing_cache.append(f"{year}년 {month}월 {kind}")
        total = row[2]
        if isinstance(total, (int, float)) and abs(sum(v or 0 for v in values) - total) > 1:
            mismatches.append(f"{year}년 {month}월 {kind}")
    name = sheet.title
    book.close()
    frame = pd.DataFrame(records)
    if frame.empty:
        raise ValueError("월별 매출·목표 데이터를 찾지 못했습니다.")
    if frame.duplicated(["연도", "월", "매장", "구분"]).any():
        raise ValueError("같은 연도·월·매장에 중복 행이 있습니다.")
    if missing_cache:
        raise ValueError("금액 전체가 비어 있습니다. Excel에서 계산 후 저장해 주세요: " + ", ".join(missing_cache[:4]))
    actual = frame[frame["구분"].eq("실적")]
    latest = int(actual["연도"].max())
    nonzero = actual[actual["연도"].eq(latest) & actual["금액"].ne(0)]
    inferred = int(nonzero["월"].max()) if not nonzero.empty else 1
    return frame, {"sheet": name, "latest_year": latest, "inferred_close": inferred,
                   "mismatches": mismatches, "stores": list(stores.values())}


def monthly(frame: pd.DataFrame, year: int, kind: str, stores: list[str]) -> pd.Series:
    chosen = frame[frame["연도"].eq(year) & frame["구분"].eq(kind) & frame["매장"].isin(stores)]
    return chosen.groupby("월")["금액"].sum().reindex(range(1, 13))


def complete_sum(series: pd.Series) -> float | None:
    return float(series.sum()) if len(series) and series.notna().all() else None


def ratio(numerator, denominator):
    if numerator is None or denominator is None or pd.isna(numerator) or pd.isna(denominator) or denominator == 0:
        return None
    return float(numerator / denominator)


def summarize(frame, year, close, stores):
    actual = monthly(frame, year, "실적", stores)
    target = monthly(frame, year, "목표", stores)
    previous = monthly(frame, year - 1, "실적", stores)
    value, plan, annual, prior = (complete_sum(actual.loc[:close]), complete_sum(target.loc[:close]),
                                 complete_sum(target), complete_sum(previous.loc[:close]))
    growth = ratio(value, prior)
    return {"actual": value, "target": plan, "annual": annual, "previous": prior,
            "attainment": ratio(value, plan), "progress": ratio(value, annual),
            "growth": None if growth is None else growth - 1,
            "gap": None if plan is None or value is None else value - plan}


def render_business_plan(root: Path):
    import streamlit as st
    import plotly.graph_objects as go

    @st.cache_data(show_spinner=False)
    def cached_read(data):
        return read_plan(data)

    st.subheader("사업계획 분석")
    st.caption("매장별 매출 목표와 마감 실적을 비교합니다. 이 탭의 조회 조건은 아래에서 별도로 선택하며, 금액은 VAT 포함입니다.")
    path = find_plan_file(root)
    with st.expander("사업계획 파일 · 매월 업데이트 안내", expanded=path is None):
        upload = st.file_uploader("사업계획 파일", type=["xlsx", "xlsm"], key="bp_upload")
        st.write(f"자동 연결: {path.name}" if path else "앱 폴더 또는 data 폴더에 사업계획으로 시작하는 엑셀 파일을 올려주세요.")
        st.caption("매월 마감 후 기존 파일을 교체하세요. 날짜를 붙일 경우 사업계획_09.30.xlsx처럼 동일한 형식을 사용하세요. 여러 파일이 있으면 파일명 날짜가 큰 파일을 우선 읽습니다.")
    if upload is None and path is None:
        return
    try:
        frame, info = cached_read(upload.getvalue() if upload else path.read_bytes())
    except Exception as exc:
        st.error(f"사업계획 자료를 읽지 못했습니다: {exc}")
        return
    source_name = upload.name if upload else path.name
    years = sorted(frame.loc[frame["구분"].eq("실적"), "연도"].unique().tolist())
    controls = st.columns([1, 1, 2])
    year = controls[0].selectbox("분석 연도", years, index=len(years) - 1, key="bp_year")
    default_close = info["inferred_close"] if year == info["latest_year"] else 12
    close = controls[1].selectbox("마감 월", list(range(1, 13)), index=default_close - 1,
                                  format_func=lambda x: f"{x}월", key=f"bp_close_{year}_{default_close}")
    selected = controls[2].multiselect("분석 매장", info["stores"], key="bp_stores", placeholder="전체 매장")
    stores = selected or info["stores"]
    st.caption(f"{source_name} · {info['sheet']} · {year}년 1~{close}월 마감 · {'전체 매장' if not selected else f'{len(stores)}개 매장'}")
    if year == info["latest_year"]:
        st.caption(f"금액이 있는 마지막 월은 {info['inferred_close']}월입니다. 이를 기본 마감 월로 사용합니다. 매출 0으로 마감한 달은 마감 월을 직접 선택하세요.")
        if close > info["inferred_close"]:
            st.warning("금액이 있는 마지막 월 이후를 선택했습니다. 선택한 월까지 실제 마감되었는지 확인하세요. 미입력 월의 0도 계산에 포함됩니다.")
    if info["mismatches"]:
        st.caption(f"원본 합계와 매장 합산이 다른 항목 {len(info['mismatches'])}건은 매장별 금액으로 계산했습니다. 자세한 항목은 아래 계산 기준에서 확인할 수 있습니다.")
    stats = summarize(frame, year, close, stores)
    money = lambda x: "자료 없음" if x is None or pd.isna(x) else f"{x / 1e8:,.2f}억 원"
    percent = lambda x: "비교 기준 없음" if x is None or pd.isna(x) else f"{x:.1%}"
    cards = st.columns(3) + st.columns(3)
    for col, label, value in zip(cards,
        ["연간 목표", "마감 누적 매출", "마감 누적 목표", "연간 목표 진척률", "누적 목표 달성률", "전년 동기 성장률"],
        [money(stats["annual"]), money(stats["actual"]), money(stats["target"]), percent(stats["progress"]), percent(stats["attainment"]), percent(stats["growth"])]):
        col.metric(label, value)
    st.caption(f"1~{close}월 누적 목표 {money(stats['target'])} · 목표 대비 차액 {money(stats['gap'])} · 전년 동기 매출 {money(stats['previous'])}")
    if stats["annual"] is None:
        st.info(f"{year}년의 연간 목표 자료가 없어 목표 달성률은 표시하지 않습니다. 매출 추이와 전년 동기 성장률은 확인할 수 있습니다.")

    def line(fig, name, x, y, color, dash=None):
        fig.add_trace(go.Scatter(name=name, x=x, y=[None if pd.isna(v) else float(v) for v in y],
                                mode="lines+markers", connectgaps=False,
                                line=dict(color=color, width=2.5, dash=dash), marker=dict(size=7),
                                hovertemplate="%{x}<br>%{y:,.0f}원<extra>%{fullData.name}</extra>"))

    def show(fig, key, title, percent_axis=False, height=390):
        fig.update_layout(title=dict(text=title, font=dict(size=17)), height=height,
                          margin=dict(l=10, r=15, t=65, b=85),
                          legend=dict(orientation="h", y=-0.2, x=0), hovermode="x unified",
                          yaxis_title="성장률" if percent_axis else "매출 (원)",
                          yaxis_tickformat=".0%" if percent_axis else ",.0f",
                          xaxis=dict(type="category"), template="plotly_white")
        st.plotly_chart(fig, width="stretch", key=key)

    history = []
    for y in years:
        same = summarize(frame, y, close, stores)
        full = complete_sum(monthly(frame, y, "실적", stores))
        if y == info["latest_year"] and (info["inferred_close"] < 12 or (y == year and close < 12)):
            full = None
        # Incomplete current-year rows cannot masquerade as future-period actuals.
        available = y != info["latest_year"] or close <= (close if y == year else info["inferred_close"])
        history.append({"연도": y, "연간 목표": same["annual"], "연간 실적": full,
                        f"1~{close}월 실적": same["actual"] if available else None,
                        "전년 동기 성장률": same["growth"] if available else None,
                        "누적 목표 달성률": same["attainment"] if available else None})
    hist = pd.DataFrame(history)
    st.divider()
    fig = go.Figure()
    line(fig, "연간 실적 (연 마감)", hist["연도"].astype(str), hist["연간 실적"], "#94a3b8")
    line(fig, f"1~{close}월 누적 실적", hist["연도"].astype(str), hist[f"1~{close}월 실적"], "#2563eb")
    line(fig, "연간 목표 (자료 있는 연도)", hist["연도"].astype(str), hist["연간 목표"], "#f59e0b", "dash")
    show(fig, "bp_year_chart", "연도별 목표 · 매출 추이")
    st.caption("성장 비교는 모든 연도에 동일한 마감 월을 적용합니다. 연 마감 전인 최신 연도의 연간 실적과 자료가 없는 연도의 목표는 표시하지 않습니다.")
    fig = go.Figure(go.Bar(x=hist["연도"].astype(str), y=hist["전년 동기 성장률"],
                          marker_color=["#2563eb" if pd.notna(v) and v >= 0 else "#ef4444" for v in hist["전년 동기 성장률"]],
                          text=["" if pd.isna(v) else f"{v:+.1%}" for v in hist["전년 동기 성장률"]], textposition="auto",
                          hovertemplate="%{x}년<br>%{y:+.1%}<extra></extra>"))
    show(fig, "bp_growth_chart", f"연도별 전년 동기 성장률 · 1~{close}월", True, 320)

    actual = monthly(frame, year, "실적", stores)
    target = monthly(frame, year, "목표", stores)
    previous = monthly(frame, year - 1, "실적", stores)
    actual.loc[close + 1:] = float("nan")
    labels = [f"{m}월" for m in range(1, 13)]
    left, right = st.columns(2)
    with left:
        fig = go.Figure()
        line(fig, "월 목표", labels, target, "#f59e0b", "dash")
        line(fig, f"{year}년 실적", labels, actual, "#2563eb")
        line(fig, f"{year-1}년 실적", labels, previous, "#94a3b8", "dot")
        show(fig, "bp_month_chart", "월별 목표 · 실적 추이")
    with right:
        fig = go.Figure()
        line(fig, "누적 목표", labels, target.cumsum(skipna=False), "#f59e0b", "dash")
        line(fig, "누적 실적", labels, actual.cumsum(skipna=False), "#2563eb")
        line(fig, "전년 누적 실적", labels, previous.cumsum(skipna=False), "#94a3b8", "dot")
        show(fig, "bp_cumulative_chart", "누적 목표 대비 진행 현황")

    st.markdown("#### 매장별 달성 현황")
    store_rows = []
    for store in stores:
        s = summarize(frame, year, close, [store])
        store_rows.append({"매장": store, "연간 목표": s["annual"], "누적 목표": s["target"],
                           "누적 실적": s["actual"], "목표 대비 차액": s["gap"],
                           "누적 목표 달성률": s["attainment"], "연간 목표 진척률": s["progress"],
                           "전년 동기 성장률": s["growth"]})
    table = pd.DataFrame(store_rows).sort_values("목표 대비 차액", na_position="last")
    eligible = table[table["누적 목표"].gt(0) & table["목표 대비 차액"].lt(0)].head(8)
    if not eligible.empty:
        fig = go.Figure(go.Bar(y=eligible["매장"][::-1], x=-eligible["목표 대비 차액"][::-1], orientation="h",
                              marker_color="#f59e0b", hovertemplate="%{y}<br>부족액 %{x:,.0f}원<extra></extra>"))
        fig.update_layout(title="누적 목표 미달 금액 · 상위 8개 매장", height=max(300, len(eligible) * 38 + 90),
                          xaxis=dict(title="목표까지 부족한 금액 (원)", tickformat=",.0f"),
                          margin=dict(l=10, r=20, t=60, b=30), template="plotly_white")
        st.plotly_chart(fig, width="stretch", key="bp_store_gap_chart")

    def show_table(data, key):
        view = data.copy()
        config = {}
        for col in view:
            if "률" in col:
                view[col] = pd.to_numeric(view[col], errors="coerce") * 100
                config[col] = st.column_config.NumberColumn(col, format="%.1f%%")
            elif col not in {"매장", "월", "연도"}:
                config[col] = st.column_config.NumberColumn(col, format="localized", help="원 · VAT 포함")
        st.dataframe(view, hide_index=True, width="stretch", column_config=config)
        st.download_button("표 내려받기 (CSV)", data.to_csv(index=False).encode("utf-8-sig"),
                           f"사업계획_{year}_{close:02d}월_{key}.csv", "text/csv", key=f"bp_download_{key}",
                           help="CSV의 비율은 0~1 기준 숫자입니다. 예: 0.885 = 88.5%")
    show_table(table, "매장별")
    st.caption("목표나 전년 실적이 0이면 해당 비율은 빈칸으로 표시합니다. 목표가 없는 매출도 전체 실적에는 포함합니다.")
    with st.expander("월별 수치 · 연도별 비교표"):
        details = pd.DataFrame({"월": labels, "월 목표": target.values, "월 실적": actual.values,
                                "전년 월 실적": previous.values,
                                "월 목표 달성률": [ratio(a, t) for a, t in zip(actual, target)],
                                "전년 동월 성장률": [None if ratio(a, p) is None else ratio(a, p)-1 for a, p in zip(actual, previous)]})
        show_table(details, "월별")
        show_table(hist, "연도별")
    with st.expander("계산 기준"):
        if info["mismatches"]:
            st.write("원본 합계와 매장 합산 차이: " + ", ".join(info["mismatches"]))
        st.markdown("- 연간 목표 진척률 = 마감 누적 실적 ÷ 1~12월 목표\n"
                    "- 누적 목표 달성률 = 마감 누적 실적 ÷ 같은 기간 목표\n"
                    "- 전년 동기 성장률 = 마감 누적 실적 ÷ 전년 같은 기간 실적 − 1\n"
                    "- 사업계획의 월별 매장 금액만 읽으며 원본 비율 수식과 합계행은 사용하지 않습니다.\n"
                    "- 이 탭의 실적은 사업계획 파일에 입력된 마감 매출입니다. ERP 상품별 매출과 별도로 분석합니다.\n"
                    "- 월별 행 안의 빈 매장 칸은 0으로 읽습니다. 미마감 월은 그래프와 달성률에서 제외합니다.")
