from __future__ import annotations

import os
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from data_loader import load_dashboard_data
from forecasting import calculate_order_recommendation, calculate_ytd_growth


st.set_page_config(
    page_title="테마상품 매입·매출 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.4rem; padding-bottom: 2rem;}
    [data-testid="stMetric"] {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        padding: 0.8rem 1rem;
        border-radius: 0.75rem;
    }
    .small-note {color:#64748b; font-size:0.88rem;}
    .recommend-card {
        background: linear-gradient(135deg, #eff6ff 0%, #ecfdf5 100%);
        border: 1px solid #93c5fd;
        border-radius: 1rem;
        padding: 1.2rem 1.4rem;
        margin: 0.5rem 0 1rem 0;
    }
    .recommend-label {color:#475569; font-size:0.95rem; font-weight:600;}
    .recommend-number {color:#0f172a; font-size:2.35rem; font-weight:750; line-height:1.2;}
    .recommend-sub {color:#64748b; font-size:0.88rem; margin-top:0.35rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


def default_file(filename: str, env_name: str) -> Path | None:
    candidates = [
        Path(os.environ[env_name]) if os.environ.get(env_name) else None,
        Path.cwd() / "data" / filename,
        Path.home() / "OneDrive" / "바탕 화면" / "개인 프로젝트" / filename,
        Path.home() / "Desktop" / "개인 프로젝트" / filename,
    ]
    return next((path for path in candidates if path and path.exists()), None)


DEFAULT_SALES = default_file("매출 Summary_Rev.1.xlsm", "SALES_SUMMARY_PATH")
DEFAULT_PURCHASE = default_file("매입 Summary_Rev.1.xlsm", "PURCHASE_SUMMARY_PATH")


@st.cache_data(show_spinner="엑셀 Rawdata를 읽고 있습니다...")
def load_from_inputs(
    sales_bytes: bytes | None,
    purchase_bytes: bytes | None,
    sales_path: str,
    purchase_path: str,
    sales_mtime: float,
    purchase_mtime: float,
):
    del sales_mtime, purchase_mtime
    sales_source = sales_bytes if sales_bytes is not None else sales_path
    purchase_source = purchase_bytes if purchase_bytes is not None else purchase_path
    return load_dashboard_data(sales_source, purchase_source)


def format_number(value: float) -> str:
    return f"{value:,.0f}"


def format_money(value: float) -> str:
    if abs(value) >= 100_000_000:
        return f"{value / 100_000_000:,.1f}억 원"
    if abs(value) >= 10_000:
        return f"{value / 10_000:,.0f}만 원"
    return f"{value:,.0f}원"


def growth(current: float, previous: float) -> float | None:
    if previous == 0:
        return None
    return current / previous - 1


def delta_text(rate: float | None) -> str:
    return "비교 기준 없음" if rate is None else f"{rate:+.1%} 전년 대비"


def filtered_for_years(
    frame: pd.DataFrame,
    years: list[int],
    months: tuple[int, int],
    categories: list[str],
    sku_filter: set[str] | None,
) -> pd.DataFrame:
    mask = (
        frame["연도"].isin(years)
        & frame["월"].between(months[0], months[1])
        & frame["유형"].isin(categories)
    )
    if sku_filter is not None:
        mask &= frame["표준SKU"].isin(sku_filter)
    return frame[mask].copy()


def monthly_values(
    frame: pd.DataFrame,
    year: int,
    value_col: str,
    months: tuple[int, int] = (1, 12),
) -> pd.Series:
    return (
        frame[frame["연도"].eq(year)]
        .groupby("월", observed=True)[value_col]
        .sum()
        .reindex(range(months[0], months[1] + 1), fill_value=0)
    )


def comparison_chart(
    frame: pd.DataFrame,
    year: int,
    value_col: str,
    title: str,
    money: bool,
    months: tuple[int, int],
) -> go.Figure:
    current = monthly_values(frame, year, value_col, months)
    previous = monthly_values(frame, year - 1, value_col, months)
    month_numbers = list(range(months[0], months[1] + 1))
    labels = [f"{month}월" for month in month_numbers]

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=labels,
            y=current.values,
            mode="lines+markers",
            name=str(year),
            line=dict(color="#2563EB", width=3),
            marker=dict(size=9),
            customdata=month_numbers,
            hovertemplate=(
                "%{x}<br>" + ("금액 %{y:,.0f}원" if money else "수량 %{y:,.0f}개") + "<extra></extra>"
            ),
        )
    )
    fig.add_trace(
        go.Scatter(
            x=labels,
            y=previous.values,
            mode="lines+markers",
            name=str(year - 1),
            line=dict(color="#94A3B8", width=2, dash="dot"),
            marker=dict(size=7),
            customdata=month_numbers,
            hovertemplate=(
                "%{x}<br>" + ("금액 %{y:,.0f}원" if money else "수량 %{y:,.0f}개") + "<extra></extra>"
            ),
        )
    )
    fig.update_layout(
        title=title,
        height=390,
        margin=dict(l=15, r=15, t=55, b=20),
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        xaxis_title=None,
        yaxis_title=None,
        yaxis_tickformat=",.0f",
    )
    return fig


def purchase_sales_flow(
    sales: pd.DataFrame,
    purchase: pd.DataFrame,
    year: int,
    months: tuple[int, int],
) -> go.Figure:
    sales_qty = monthly_values(sales, year, "수량", months)
    purchase_qty = monthly_values(purchase, year, "수량", months)
    labels = [f"{month}월" for month in range(months[0], months[1] + 1)]

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=labels,
            y=purchase_qty.values,
            name="매입수량",
            marker_color="#A7F3D0",
            hovertemplate="%{x}<br>매입 %{y:,.0f}개<extra></extra>",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=labels,
            y=sales_qty.values,
            name="판매수량",
            mode="lines+markers",
            line=dict(color="#0F766E", width=3),
            marker=dict(size=8),
            hovertemplate="%{x}<br>판매 %{y:,.0f}개<extra></extra>",
        )
    )
    fig.update_layout(
        title=f"{year}년 매입 시점과 판매 흐름",
        height=390,
        margin=dict(l=15, r=15, t=55, b=20),
        barmode="group",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis_tickformat=",.0f",
    )
    return fig


def selected_month_from_event(event, fallback: int) -> int:
    try:
        points = event.selection.points
    except (AttributeError, TypeError):
        try:
            points = event.get("selection", {}).get("points", [])
        except AttributeError:
            return fallback
    if not points:
        return fallback
    custom = points[0].get("customdata")
    if isinstance(custom, (list, tuple)):
        custom = custom[0] if custom else None
    try:
        return int(custom)
    except (TypeError, ValueError):
        return fallback


def item_comparison(
    frame: pd.DataFrame,
    year: int,
    month: int,
    amount_col: str,
) -> pd.DataFrame:
    keys = ["표준SKU", "품목명"]
    current = (
        frame[frame["연도"].eq(year) & frame["월"].eq(month)]
        .groupby(keys, observed=True)[["수량", amount_col]]
        .sum()
        .rename(columns={"수량": "금년 수량", amount_col: "금년 금액"})
    )
    previous = (
        frame[frame["연도"].eq(year - 1) & frame["월"].eq(month)]
        .groupby(keys, observed=True)[["수량", amount_col]]
        .sum()
        .rename(columns={"수량": "전년 수량", amount_col: "전년 금액"})
    )
    merged = current.join(previous, how="outer").fillna(0).reset_index()
    merged["수량 성장률"] = (merged["금년 수량"] / merged["전년 수량"] - 1).where(
        merged["전년 수량"].ne(0)
    )
    merged["금액 성장률"] = (merged["금년 금액"] / merged["전년 금액"] - 1).where(
        merged["전년 금액"].ne(0)
    )
    merged["금년 평균단가"] = (merged["금년 금액"] / merged["금년 수량"]).where(
        merged["금년 수량"].ne(0)
    )
    merged["전년 평균단가"] = (merged["전년 금액"] / merged["전년 수량"]).where(
        merged["전년 수량"].ne(0)
    )
    return merged.sort_values("금년 금액", ascending=False)


def build_product_catalog(frame: pd.DataFrame) -> pd.DataFrame:
    catalog = (
        frame[["유형", "표준SKU", "품목명", "모델번호", "COLOR", "SIZE"]]
        .sort_values(["유형", "품목명", "COLOR", "SIZE"])
        .drop_duplicates("표준SKU")
        .reset_index(drop=True)
    )

    def make_label(row: pd.Series) -> str:
        details = []
        if str(row["COLOR"]).strip() not in {"", "nan"}:
            details.append(str(row["COLOR"]).strip())
        if str(row["SIZE"]).strip() not in {"", "0", "0.0", "nan"}:
            details.append(str(row["SIZE"]).strip())
        option = " / ".join(details)
        option_text = f" · {option}" if option else ""
        return f"{row['품목명']}{option_text}  [{row['표준SKU']}]"

    catalog["선택표시"] = catalog.apply(make_label, axis=1)
    return catalog


st.title("테마상품 매입·매출 대시보드")
st.caption("Rawdata는 수정하지 않고 읽기 전용으로 분석합니다.")

with st.sidebar:
    st.header("데이터 연결")
    st.caption("매출·매입 파일을 모두 올리면 자동으로 분석을 시작합니다.")
    sales_upload = st.file_uploader(
        "매출 Summary",
        type=["xlsm", "xlsx"],
        help="Rawdata 시트가 있는 매출 Summary 파일을 올려주세요.",
    )
    purchase_upload = st.file_uploader(
        "매입 Summary",
        type=["xlsm", "xlsx"],
        help="Purchase_Rawdata 시트가 있는 매입 Summary 파일을 올려주세요.",
    )

sales_bytes = sales_upload.getvalue() if sales_upload else None
purchase_bytes = purchase_upload.getvalue() if purchase_upload else None
sales_path = str(DEFAULT_SALES or "")
purchase_path = str(DEFAULT_PURCHASE or "")

if sales_bytes is None and not sales_path:
    st.info("왼쪽에서 매출 Summary 파일을 올려주세요.")
    st.stop()
if purchase_bytes is None and not purchase_path:
    st.info("왼쪽에서 매입 Summary 파일을 올려주세요.")
    st.stop()

try:
    sales, purchase, quality = load_from_inputs(
        sales_bytes,
        purchase_bytes,
        sales_path,
        purchase_path,
        0 if sales_bytes is not None else Path(sales_path).stat().st_mtime,
        0 if purchase_bytes is not None else Path(purchase_path).stat().st_mtime,
    )
except Exception as exc:
    st.error(f"엑셀 파일을 읽는 중 문제가 발생했습니다: {exc}")
    st.stop()

valid_years = sorted(
    set(sales["연도"].dropna().astype(int)) | set(purchase["연도"].dropna().astype(int))
)
if not valid_years:
    st.error("분석 가능한 날짜가 없습니다.")
    st.stop()

with st.sidebar:
    if sales_upload:
        st.success(f"매출: {sales_upload.name}")
    else:
        st.success(f"매출: {Path(sales_path).name}")
    if purchase_upload:
        st.success(f"매입: {purchase_upload.name}")
    else:
        st.success(f"매입: {Path(purchase_path).name}")

    st.divider()
    st.header("조회 조건")
    base_year = st.selectbox("기준 연도", valid_years, index=len(valid_years) - 1)
    year_sales_months = sales.loc[sales["연도"].eq(base_year), "월"].dropna()
    latest_sales_month = int(year_sales_months.max()) if not year_sales_months.empty else 12
    if st.session_state.get("month_range_year") != base_year:
        st.session_state["month_range"] = (1, latest_sales_month)
        st.session_state["month_range_year"] = base_year
    month_range = st.slider("조회 월", 1, 12, key="month_range")
    category_options = sorted(set(sales["유형"]) | set(purchase["유형"]))
    selected_categories = st.multiselect(
        "상품 유형",
        category_options,
        default=category_options,
    )
    amount_label = st.radio("금액 기준", ["합계(부가세 포함)", "공급가액"], horizontal=False)
    amount_col = "합계" if amount_label.startswith("합계") else "공급가액"

    product_options = sorted(sales["품목명"].dropna().loc[lambda x: x.ne("")].unique())
    selected_product = st.selectbox("품목 선택", ["전체"] + product_options)

if not selected_categories:
    st.warning("상품 유형을 하나 이상 선택해 주세요.")
    st.stop()

sku_filter = None
if selected_product != "전체":
    sku_filter = set(sales.loc[sales["품목명"].eq(selected_product), "표준SKU"])

years_to_compare = [base_year - 1, base_year]
sales_view = filtered_for_years(
    sales, years_to_compare, month_range, selected_categories, sku_filter
)
purchase_view = filtered_for_years(
    purchase, years_to_compare, month_range, selected_categories, sku_filter
)

current_sales = sales_view[sales_view["연도"].eq(base_year)]
previous_sales = sales_view[sales_view["연도"].eq(base_year - 1)]
current_purchase = purchase_view[purchase_view["연도"].eq(base_year)]
previous_purchase = purchase_view[purchase_view["연도"].eq(base_year - 1)]

sales_qty = float(current_sales["수량"].fillna(0).sum())
previous_sales_qty = float(previous_sales["수량"].fillna(0).sum())
sales_amount = float(current_sales[amount_col].fillna(0).sum())
previous_sales_amount = float(previous_sales[amount_col].fillna(0).sum())
purchase_qty = float(current_purchase["수량"].fillna(0).sum())
previous_purchase_qty = float(previous_purchase["수량"].fillna(0).sum())
purchase_amount = float(current_purchase[amount_col].fillna(0).sum())
previous_purchase_amount = float(previous_purchase[amount_col].fillna(0).sum())
average_price = sales_amount / sales_qty if sales_qty else 0
previous_average_price = previous_sales_amount / previous_sales_qty if previous_sales_qty else 0

overview_tab, detail_tab, forecast_tab, quality_tab = st.tabs(
    ["전체 현황", "월 상세", "예측 발주", "데이터 상태"]
)

with overview_tab:
    sales_metric_columns = st.columns(3)
    sales_metric_columns[0].metric(
        "판매수량",
        f"{format_number(sales_qty)}개",
        delta_text(growth(sales_qty, previous_sales_qty)),
    )
    sales_metric_columns[1].metric(
        "매출액",
        format_money(sales_amount),
        delta_text(growth(sales_amount, previous_sales_amount)),
    )
    sales_metric_columns[2].metric(
        "평균 판매단가",
        f"{format_number(average_price)}원",
        delta_text(growth(average_price, previous_average_price)),
    )
    purchase_metric_columns = st.columns(2)
    purchase_metric_columns[0].metric(
        "매입수량",
        f"{format_number(purchase_qty)}개",
        delta_text(growth(purchase_qty, previous_purchase_qty)),
    )
    purchase_metric_columns[1].metric(
        "매입금액",
        format_money(purchase_amount),
        delta_text(growth(purchase_amount, previous_purchase_amount)),
    )

    left, right = st.columns(2)
    with left:
        sales_chart = comparison_chart(
            sales_view,
            base_year,
            amount_col,
            "월별 매출액 전년 비교",
            True,
            month_range,
        )
        event = st.plotly_chart(
            sales_chart,
            width="stretch",
            key="sales_month_chart",
            on_select="rerun",
            selection_mode="points",
        )
    with right:
        st.plotly_chart(
            purchase_sales_flow(sales_view, purchase_view, base_year, month_range),
            width="stretch",
            key="purchase_sales_flow",
        )

    default_month = int(
        current_sales["월"].dropna().max()
        if not current_sales["월"].dropna().empty
        else month_range[1]
    )
    selected_from_chart = selected_month_from_event(event, default_month)
    st.session_state["chart_selected_month"] = selected_from_chart
    st.info(
        f"그래프에서 월을 클릭하면 ‘월 상세’ 탭에서 해당 월의 상품 구성을 볼 수 있습니다. "
        f"현재 선택: {base_year}년 {selected_from_chart}월"
    )

with detail_tab:
    detail_default = int(st.session_state.get("chart_selected_month", month_range[1]))
    detail_month = st.selectbox(
        "상세 조회 월",
        list(range(month_range[0], month_range[1] + 1)),
        index=max(0, min(month_range[1], detail_default) - month_range[0]),
    )
    detail = item_comparison(sales_view, base_year, detail_month, amount_col)

    month_current = sales_view[
        sales_view["연도"].eq(base_year) & sales_view["월"].eq(detail_month)
    ]
    month_previous = sales_view[
        sales_view["연도"].eq(base_year - 1) & sales_view["월"].eq(detail_month)
    ]
    current_qty = float(month_current["수량"].fillna(0).sum())
    previous_qty = float(month_previous["수량"].fillna(0).sum())
    current_amount = float(month_current[amount_col].fillna(0).sum())
    previous_amount = float(month_previous[amount_col].fillna(0).sum())
    qty_growth = growth(current_qty, previous_qty)
    amount_growth = growth(current_amount, previous_amount)
    current_asp = current_amount / current_qty if current_qty else 0
    previous_asp = previous_amount / previous_qty if previous_qty else 0

    detail_metrics = st.columns(4)
    detail_metrics[0].metric("월 판매수량", f"{format_number(current_qty)}개", delta_text(qty_growth))
    detail_metrics[1].metric("월 매출액", format_money(current_amount), delta_text(amount_growth))
    detail_metrics[2].metric(
        "월 평균단가",
        f"{format_number(current_asp)}원",
        delta_text(growth(current_asp, previous_asp)),
    )
    detail_metrics[3].metric("판매 품목 수", f"{len(detail[detail['금년 수량'].ne(0)]):,}개")

    if qty_growth is not None and amount_growth is not None:
        if qty_growth > amount_growth and current_asp < previous_asp:
            st.warning(
                f"판매수량 변화율({qty_growth:+.1%})보다 매출액 변화율({amount_growth:+.1%})이 낮습니다. "
                f"평균 판매단가가 {previous_asp:,.0f}원에서 {current_asp:,.0f}원으로 내려간 영향이 있습니다."
            )
        else:
            st.success(
                f"{base_year}년 {detail_month}월 판매수량은 전년 대비 {qty_growth:+.1%}, "
                f"매출액은 {amount_growth:+.1%}입니다."
            )

    chart_column, table_column = st.columns([1, 1.4])
    with chart_column:
        top_items = detail.head(12).sort_values("금년 금액")
        item_fig = go.Figure(
            go.Bar(
                x=top_items["금년 금액"],
                y=top_items["품목명"],
                orientation="h",
                marker_color="#38BDF8",
                hovertemplate="%{y}<br>%{x:,.0f}원<extra></extra>",
            )
        )
        item_fig.update_layout(
            title=f"{base_year}년 {detail_month}월 주요 매출 품목",
            height=510,
            margin=dict(l=10, r=10, t=55, b=20),
            xaxis_tickformat=",.0f",
            yaxis_title=None,
            xaxis_title=None,
        )
        st.plotly_chart(item_fig, width="stretch")

    with table_column:
        display = detail[
            [
                "품목명",
                "금년 수량",
                "전년 수량",
                "수량 성장률",
                "금년 금액",
                "전년 금액",
                "금액 성장률",
                "금년 평균단가",
                "전년 평균단가",
            ]
        ].copy()
        display["수량 성장률"] = display["수량 성장률"] * 100
        display["금액 성장률"] = display["금액 성장률"] * 100
        st.dataframe(
            display,
            width="stretch",
            hide_index=True,
            height=510,
            column_config={
                "금년 수량": st.column_config.NumberColumn(format="%,.0f"),
                "전년 수량": st.column_config.NumberColumn(format="%,.0f"),
                "수량 성장률": st.column_config.NumberColumn(format="%.1f%%"),
                "금년 금액": st.column_config.NumberColumn(format="%,.0f원"),
                "전년 금액": st.column_config.NumberColumn(format="%,.0f원"),
                "금액 성장률": st.column_config.NumberColumn(format="%.1f%%"),
                "금년 평균단가": st.column_config.NumberColumn(format="%,.0f원"),
                "전년 평균단가": st.column_config.NumberColumn(format="%,.0f원"),
            },
        )

with forecast_tab:
    st.markdown("### 예측 발주")
    st.caption(
        "작년 같은 기간의 실제 판매량에 올해 판매 추세와 안전재고를 반영한 뒤, "
        "현재고와 입고 예정 수량을 차감합니다."
    )

    product_catalog = build_product_catalog(sales)
    product_categories = sorted(product_catalog["유형"].dropna().unique())
    preferred_sku = "MAPSS03ZZZ230"
    preferred_categories = product_catalog.loc[
        product_catalog["표준SKU"].eq(preferred_sku), "유형"
    ]
    default_category = (
        preferred_categories.iloc[0]
        if not preferred_categories.empty
        else product_categories[0]
    )

    selector_left, selector_right = st.columns([1.1, 1])
    with selector_left:
        forecast_category = st.selectbox(
            "상품 유형",
            product_categories,
            index=product_categories.index(default_category),
            key="forecast_category",
        )
        category_catalog = product_catalog[
            product_catalog["유형"].eq(forecast_category)
        ].copy()
        product_labels = dict(
            zip(category_catalog["표준SKU"], category_catalog["선택표시"], strict=False)
        )
        selected_sku = st.selectbox(
            "예측할 품목",
            category_catalog["표준SKU"].tolist(),
            index=(
                category_catalog["표준SKU"].tolist().index(preferred_sku)
                if preferred_sku in category_catalog["표준SKU"].tolist()
                else 0
            ),
            format_func=lambda sku: product_labels.get(sku, sku),
            key="forecast_sku",
        )

    latest_sales_date = sales["거래일자"].dropna().max().normalize()
    today = pd.Timestamp.today().normalize()
    default_forecast_start = max(latest_sales_date + pd.Timedelta(days=1), today)
    default_forecast_end = (
        default_forecast_start + pd.DateOffset(months=3) - pd.Timedelta(days=1)
    )

    with selector_right:
        forecast_dates = st.date_input(
            "예측 기간",
            value=(default_forecast_start.date(), default_forecast_end.date()),
            help="작년의 동일한 달·일 구간과 비교합니다.",
            key="forecast_dates",
        )
        if isinstance(forecast_dates, (tuple, list)) and len(forecast_dates) == 2:
            forecast_start, forecast_end = forecast_dates
        elif isinstance(forecast_dates, (tuple, list)) and len(forecast_dates) == 1:
            forecast_start = forecast_end = forecast_dates[0]
            st.info("종료일을 선택하면 기간 전체를 예측할 수 있습니다.")
        else:
            forecast_start = forecast_end = forecast_dates

        current_stock = int(
            st.number_input(
                "현재 사용 가능한 재고",
                min_value=0,
                value=0,
                step=1,
                key=f"current_stock_{selected_sku}",
            )
        )

    selected_product = category_catalog[
        category_catalog["표준SKU"].eq(selected_sku)
    ].iloc[0]

    growth_as_of = min(
        latest_sales_date,
        pd.Timestamp(forecast_start) - pd.Timedelta(days=1),
    )
    growth_info = calculate_ytd_growth(sales, selected_sku, growth_as_of)
    observed_growth = float(growth_info["growth_rate"])

    assumption_box = st.container(border=True)
    with assumption_box:
        st.markdown("#### 발주 가정")
        assumption_columns = st.columns(4)
        with assumption_columns[0]:
            incoming_stock = int(
                st.number_input(
                    "입고 예정 수량",
                    min_value=0,
                    value=0,
                    step=1,
                    key=f"incoming_stock_{selected_sku}",
                )
            )
        with assumption_columns[1]:
            growth_mode = st.selectbox(
                "성장률 반영 방식",
                ["올해 추세 자동 반영", "작년 수준 유지", "직접 입력"],
                key=f"growth_mode_{selected_sku}",
            )
        with assumption_columns[2]:
            if growth_mode == "직접 입력":
                growth_percent = float(
                    st.number_input(
                        "예상 성장률(%)",
                        min_value=-100.0,
                        value=0.0,
                        step=1.0,
                        key=f"manual_growth_{selected_sku}",
                    )
                )
                applied_growth = growth_percent / 100
            elif growth_mode == "작년 수준 유지":
                applied_growth = 0.0
                st.metric("적용 성장률", "0.0%")
            else:
                applied_growth = observed_growth
                st.metric("적용 성장률", f"{applied_growth:+.1%}")
        with assumption_columns[3]:
            safety_percent = float(
                st.number_input(
                    "안전재고율(%)",
                    min_value=0.0,
                    value=10.0,
                    step=1.0,
                    key=f"safety_rate_{selected_sku}",
                )
            )

        unit_column, trend_column = st.columns([1, 2])
        with unit_column:
            order_unit = int(
                st.number_input(
                    "발주 단위",
                    min_value=1,
                    value=1,
                    step=1,
                    help="박스입수나 최소 주문단위에 맞춰 올림합니다.",
                    key=f"order_unit_{selected_sku}",
                )
            )
        with trend_column:
            if growth_info["has_comparison"]:
                st.caption(
                    f"자동 성장률: {growth_info['current_start']:%Y-%m-%d}~"
                    f"{growth_info['current_end']:%Y-%m-%d} 판매 "
                    f"{growth_info['current_quantity']:,.0f}개 ÷ 전년 동기 "
                    f"{growth_info['previous_quantity']:,.0f}개 − 1 = "
                    f"{observed_growth:+.1%}"
                )
            else:
                st.caption("전년 동기 판매량이 없어 자동 성장률은 0%로 적용합니다.")

    forecast = calculate_order_recommendation(
        sales=sales,
        sku=selected_sku,
        forecast_start=forecast_start,
        forecast_end=forecast_end,
        current_stock=current_stock,
        incoming_stock=incoming_stock,
        growth_rate=applied_growth,
        safety_rate=safety_percent / 100,
        order_unit=order_unit,
    )

    st.markdown(
        f"""
        <div class="recommend-card">
            <div class="recommend-label">추천 발주수량</div>
            <div class="recommend-number">{forecast['recommended_order']:,.0f}개</div>
            <div class="recommend-sub">
                {forecast['forecast_start']:%Y-%m-%d} ~ {forecast['forecast_end']:%Y-%m-%d}
                · {forecast['period_days']:,}일 · 발주단위 {forecast['order_unit']:,}개 적용
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    forecast_metrics = st.columns(3)
    forecast_metrics[0].metric(
        "작년 동일 기간 판매",
        f"{forecast['last_year_demand']:,.0f}개",
    )
    forecast_metrics[1].metric(
        "성장 반영 예상수요",
        f"{forecast['forecast_demand']:,.0f}개",
        f"{applied_growth:+.1%} 적용",
    )
    forecast_metrics[2].metric(
        "안전재고 포함 목표",
        f"{forecast['target_stock']:,.0f}개",
        f"안전재고 {forecast['safety_stock']:,.0f}개",
    )

    stock_metrics = st.columns(3)
    stock_metrics[0].metric("현재고 + 입고예정", f"{forecast['available_stock']:,.0f}개")
    stock_metrics[1].metric("발주 후 예상 잔여재고", f"{forecast['expected_end_stock']:,.0f}개")
    if forecast["coverage_days"] is None:
        coverage_label = "수요 없음"
    else:
        coverage_label = f"약 {forecast['coverage_days']:,.0f}일"
    stock_metrics[2].metric("발주 전 재고 커버기간", coverage_label)

    stock_ratio = (
        min(1.0, forecast["available_stock"] / forecast["target_stock"])
        if forecast["target_stock"] > 0
        else 1.0
    )
    st.progress(
        stock_ratio,
        text=f"현재고·입고예정으로 목표수량의 {stock_ratio:.0%}를 충족합니다.",
    )

    if forecast["last_year_demand"] == 0:
        st.warning(
            "작년 같은 기간의 판매가 0개입니다. 신상품 또는 판매 이력이 짧은 품목이라면 "
            "예상 성장률을 직접 입력하거나 유사 품목을 함께 참고해 주세요."
        )
    elif forecast["recommended_order"] == 0:
        st.success("현재고와 입고 예정 수량으로 예측기간을 충분히 버틸 수 있습니다.")
    elif forecast["expected_stockout"] is not None:
        st.warning(
            f"발주하지 않으면 약 {forecast['expected_stockout']:%Y-%m-%d}에 재고가 소진될 것으로 예상됩니다."
        )
    else:
        st.info("추천 발주수량에는 예상수요와 안전재고, 발주단위 올림이 반영돼 있습니다.")

    weekly_column, stock_column = st.columns([1.35, 1])
    with weekly_column:
        weekly = forecast["weekly"]
        weekly_fig = go.Figure()
        weekly_fig.add_trace(
            go.Bar(
                x=weekly["표시 주차"],
                y=weekly["작년 판매량"],
                name="작년 판매량",
                marker_color="#CBD5E1",
                hovertemplate="%{x}<br>작년 %{y:,.0f}개<extra></extra>",
            )
        )
        weekly_fig.add_trace(
            go.Scatter(
                x=weekly["표시 주차"],
                y=weekly["예상 판매량"],
                name="예상 판매량",
                mode="lines+markers",
                line=dict(color="#2563EB", width=3),
                marker=dict(size=8),
                hovertemplate="%{x}<br>예상 %{y:,.1f}개<extra></extra>",
            )
        )
        weekly_fig.update_layout(
            title="작년 주차별 판매 추이와 예상수요",
            height=410,
            margin=dict(l=15, r=15, t=55, b=20),
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            xaxis_title=None,
            yaxis_title=None,
            yaxis_tickformat=",.0f",
        )
        st.plotly_chart(weekly_fig, width="stretch", key="forecast_weekly_chart")

    with stock_column:
        stock_fig = go.Figure()
        stock_fig.add_trace(
            go.Bar(
                y=["확보 수량"],
                x=[forecast["current_stock"]],
                name="현재고",
                orientation="h",
                marker_color="#60A5FA",
                hovertemplate="현재고 %{x:,.0f}개<extra></extra>",
            )
        )
        stock_fig.add_trace(
            go.Bar(
                y=["확보 수량"],
                x=[forecast["incoming_stock"]],
                name="입고예정",
                orientation="h",
                marker_color="#A7F3D0",
                hovertemplate="입고예정 %{x:,.0f}개<extra></extra>",
            )
        )
        stock_fig.add_trace(
            go.Bar(
                y=["확보 수량"],
                x=[forecast["recommended_order"]],
                name="추천발주",
                orientation="h",
                marker_color="#FBBF24",
                hovertemplate="추천발주 %{x:,.0f}개<extra></extra>",
            )
        )
        stock_fig.add_vline(
            x=forecast["target_stock"],
            line_dash="dash",
            line_color="#DC2626",
            annotation_text=f"목표 {forecast['target_stock']:,.0f}개",
            annotation_position="top",
        )
        stock_fig.update_layout(
            title="목표재고 충족 구성",
            height=410,
            margin=dict(l=15, r=15, t=55, b=20),
            barmode="stack",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            xaxis_title="수량",
            yaxis_title=None,
            xaxis_tickformat=",.0f",
        )
        st.plotly_chart(stock_fig, width="stretch", key="forecast_stock_chart")

    with st.expander("계산 근거 자세히 보기"):
        st.markdown(
            f"""
            **선택 품목:** {selected_product['품목명']}  
            **작년 비교기간:** {forecast['previous_start']:%Y-%m-%d} ~ {forecast['previous_end']:%Y-%m-%d}

            1. 작년 동일 기간 판매량: **{forecast['last_year_demand']:,.0f}개**
            2. 성장률 반영 예상수요: `{forecast['last_year_demand']:,.0f} × (1 {applied_growth:+.1%})` → **{forecast['forecast_demand']:,.0f}개**
            3. 안전재고: `{forecast['forecast_demand']:,.0f} × {forecast['safety_rate']:.1%}` → **{forecast['safety_stock']:,.0f}개**
            4. 목표재고: **{forecast['target_stock']:,.0f}개**
            5. 현재고·입고예정 차감 후 부족량: **{forecast['shortage']:,.0f}개**
            6. 발주단위 {forecast['order_unit']:,}개로 올림: **{forecast['recommended_order']:,.0f}개**
            """
        )

with quality_tab:
    sales_quality = quality["sales"]
    purchase_quality = quality["purchase"]

    if (
        sales_quality["invalid_dates"] == 0
        and purchase_quality["invalid_dates"] == 0
        and sales_quality["category_blank_rows"] == 0
        and purchase_quality["category_blank_rows"] == 0
    ):
        st.success("날짜와 상품 유형은 모두 정상이며 합계행도 Rawdata에 포함되지 않았습니다.")

    quality_metrics = st.columns(5)
    quality_metrics[0].metric("매출 상세행", f"{sales_quality['rows']:,}건")
    quality_metrics[1].metric("매입 상세행", f"{purchase_quality['rows']:,}건")
    quality_metrics[2].metric("SKU 연결률", f"{sales_quality['sku_match_rate']:.1%}")
    quality_metrics[3].metric(
        "매출 금액 공란",
        f"{sales_quality['missing_amount_rows']:,}건",
        delta_color="off",
    )
    quality_metrics[4].metric(
        "매입 금액 공란",
        f"{purchase_quality['missing_amount_rows']:,}건",
        delta_color="off",
    )

    st.markdown("#### 확인이 필요한 항목")
    st.markdown(
        f"""
        - 매출의 `판매처명` 공란 {sales_quality['original_customer_blank_rows']:,}건은 `창고명`을 이용해 화면용 거래처로 보완했습니다.
        - 매출 SKU는 `모델번호 + COLOR + SIZE` 조합과 숫자 사이즈 3자리 보정으로 {sales_quality['sku_matched_rows']:,}건을 매입 품목코드에 연결했습니다.
        - 금액 공란 행은 삭제하지 않고 수량은 포함하며, 금액 합산에서는 0으로 취급합니다.
        - 완전히 동일해 보이는 행도 실제 반복 판매일 수 있으므로 자동 삭제하지 않습니다.
        """
    )

    unmatched = (
        sales[~sales["SKU연결상태"]]
        .groupby(["유형", "모델번호", "품목명"], observed=True)["수량"]
        .sum()
        .reset_index()
        .sort_values("수량", ascending=False)
    )
    with st.expander("매입 품목코드와 연결되지 않은 매출 품목 보기"):
        st.dataframe(unmatched, width="stretch", hide_index=True)

st.caption(
    "음수 수량은 반품으로 포함됩니다. 금액 기준은 왼쪽에서 합계 또는 공급가액으로 바꿀 수 있습니다."
)
