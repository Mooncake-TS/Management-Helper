from __future__ import annotations

import math
from datetime import date

import pandas as pd


def _quantity_sum(frame: pd.DataFrame) -> float:
    return float(pd.to_numeric(frame["수량"], errors="coerce").fillna(0).sum())


def calculate_ytd_growth(
    sales: pd.DataFrame,
    current_sku: str,
    as_of: date | pd.Timestamp,
    previous_sku: str | None = None,
) -> dict:
    as_of_date = pd.Timestamp(as_of).normalize()
    current_start = pd.Timestamp(year=as_of_date.year, month=1, day=1)
    previous_start = current_start - pd.DateOffset(years=1)
    previous_end = as_of_date - pd.DateOffset(years=1)
    comparison_sku = previous_sku or current_sku

    current_qty = _quantity_sum(
        sales[
            sales["표준SKU"].eq(current_sku)
            & sales["거래일자"].between(current_start, as_of_date)
        ]
    )
    previous_qty = _quantity_sum(
        sales[
            sales["표준SKU"].eq(comparison_sku)
            & sales["거래일자"].between(previous_start, previous_end)
        ]
    )

    growth_rate = current_qty / previous_qty - 1 if previous_qty > 0 else 0.0
    return {
        "current_start": current_start,
        "current_end": as_of_date,
        "previous_start": previous_start,
        "previous_end": previous_end,
        "current_quantity": current_qty,
        "previous_quantity": previous_qty,
        "growth_rate": growth_rate,
        "has_comparison": previous_qty > 0,
        "current_sku": current_sku,
        "previous_sku": comparison_sku,
    }


def calculate_order_recommendation(
    sales: pd.DataFrame,
    sku: str,
    forecast_start: date | pd.Timestamp,
    forecast_end: date | pd.Timestamp,
    current_stock: int,
    incoming_stock: int,
    growth_rate: float,
    safety_rate: float,
    order_unit: int,
) -> dict:
    start = pd.Timestamp(forecast_start).normalize()
    end = pd.Timestamp(forecast_end).normalize()
    if end < start:
        raise ValueError("예측 종료일은 시작일보다 빠를 수 없습니다.")
    if current_stock < 0 or incoming_stock < 0:
        raise ValueError("재고와 입고 예정 수량은 0 이상이어야 합니다.")
    if growth_rate < -1:
        raise ValueError("예상 성장률은 -100%보다 낮을 수 없습니다.")
    if safety_rate < 0:
        raise ValueError("안전재고율은 0% 이상이어야 합니다.")
    if order_unit < 1:
        raise ValueError("발주 단위는 1 이상이어야 합니다.")

    previous_start = start - pd.DateOffset(years=1)
    previous_end = end - pd.DateOffset(years=1)
    selected = sales[
        sales["표준SKU"].eq(sku)
        & sales["거래일자"].between(previous_start, previous_end)
    ].copy()

    last_year_demand = max(0.0, _quantity_sum(selected))
    forecast_demand = math.ceil(max(0.0, last_year_demand * (1 + growth_rate)))
    safety_stock = math.ceil(forecast_demand * safety_rate)
    target_stock = forecast_demand + safety_stock
    available_stock = current_stock + incoming_stock
    shortage = max(0, target_stock - available_stock)
    recommended_order = math.ceil(shortage / order_unit) * order_unit if shortage else 0
    expected_end_stock = available_stock + recommended_order - forecast_demand

    period_days = int((end - start).days) + 1
    daily_demand = forecast_demand / period_days if period_days else 0.0
    coverage_days = available_stock / daily_demand if daily_demand > 0 else None
    expected_stockout = None
    if coverage_days is not None and coverage_days < period_days:
        expected_stockout = start + pd.Timedelta(days=max(0, math.floor(coverage_days)))

    number_of_weeks = math.ceil(period_days / 7)
    if selected.empty:
        weekly = pd.DataFrame({"주차": range(1, number_of_weeks + 1), "작년 판매량": 0.0})
    else:
        selected["주차"] = (
            ((selected["거래일자"] - previous_start).dt.days // 7) + 1
        ).astype(int)
        weekly = (
            selected.groupby("주차", observed=True)["수량"]
            .sum()
            .reindex(range(1, number_of_weeks + 1), fill_value=0)
            .rename("작년 판매량")
            .reset_index()
        )

    weekly["예상 판매량"] = (weekly["작년 판매량"] * (1 + growth_rate)).clip(lower=0)
    weekly["예측 주 시작일"] = weekly["주차"].map(
        lambda week: start + pd.Timedelta(days=(int(week) - 1) * 7)
    )
    weekly["표시 주차"] = weekly.apply(
        lambda row: f"{int(row['주차'])}주차 ({row['예측 주 시작일']:%m/%d})",
        axis=1,
    )

    peak_week = None
    if not weekly.empty and weekly["작년 판매량"].max() > 0:
        peak_row = weekly.loc[weekly["작년 판매량"].idxmax()]
        peak_week = {
            "week": int(peak_row["주차"]),
            "label": str(peak_row["표시 주차"]),
            "last_year_quantity": float(peak_row["작년 판매량"]),
        }

    return {
        "forecast_start": start,
        "forecast_end": end,
        "previous_start": previous_start,
        "previous_end": previous_end,
        "period_days": period_days,
        "last_year_demand": last_year_demand,
        "growth_rate": growth_rate,
        "forecast_demand": forecast_demand,
        "safety_rate": safety_rate,
        "safety_stock": safety_stock,
        "target_stock": target_stock,
        "current_stock": current_stock,
        "incoming_stock": incoming_stock,
        "available_stock": available_stock,
        "shortage": shortage,
        "order_unit": order_unit,
        "recommended_order": recommended_order,
        "expected_end_stock": expected_end_stock,
        "daily_demand": daily_demand,
        "coverage_days": coverage_days,
        "expected_stockout": expected_stockout,
        "weekly": weekly,
        "peak_week": peak_week,
    }
