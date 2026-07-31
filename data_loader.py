from __future__ import annotations

import io
import re
from pathlib import Path
from typing import BinaryIO

import pandas as pd


ExcelSource = str | Path | bytes | bytearray | BinaryIO

SALES_SHEET = "Rawdata"
PURCHASE_SHEET = "Purchase_Rawdata"


def _excel_source(source: ExcelSource) -> ExcelSource:
    if isinstance(source, (bytes, bytearray)):
        return io.BytesIO(source)
    return source


def _clean_header(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", "", str(value)).strip()


def _read_raw_sheet(source: ExcelSource, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(
        _excel_source(source),
        sheet_name=sheet_name,
        header=None,
        engine="openpyxl",
        dtype=object,
    )

    header_row = None
    for idx in range(min(20, len(raw))):
        headers = [_clean_header(value) for value in raw.iloc[idx].tolist()]
        if any("일자" in header for header in headers) and "수량" in headers:
            header_row = idx
            break

    if header_row is None:
        raise ValueError(f"'{sheet_name}' 시트에서 헤더 행을 찾지 못했습니다.")

    frame = raw.iloc[header_row + 1 :].copy()
    frame.columns = [str(value).strip() for value in raw.iloc[header_row].tolist()]
    frame = frame.dropna(how="all").reset_index(drop=True)

    unnamed = [column for column in frame.columns if not column or column.startswith("Unnamed")]
    if unnamed:
        frame = frame.drop(columns=unnamed)
    return frame


def _require_columns(frame: pd.DataFrame, required: list[str], source_name: str) -> None:
    missing = [column for column in required if column not in frame.columns]
    if missing:
        raise ValueError(
            f"{source_name}에 필요한 컬럼이 없습니다: {', '.join(missing)}"
        )


def _text(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()


def _numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def _normalize_part(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip().upper()
    if text.endswith(".0"):
        text = text[:-2]
    return re.sub(r"[^A-Z0-9]", "", text)


def _normalize_size(value: object, pad_numeric: bool) -> str:
    text = _normalize_part(value)
    if pad_numeric and text.isdigit():
        return text.zfill(3)
    return text


def _purchase_code(value: object) -> str:
    return _normalize_part(value)


def _sales_sku_candidates(model: object, color: object, size: object) -> list[str]:
    model_part = _normalize_part(model)
    color_part = _normalize_part(color)
    size_raw = _normalize_size(size, False)
    size_padded = _normalize_size(size, True)

    candidates: list[str] = []
    for candidate in (
        model_part + color_part + size_padded,
        model_part + size_padded,
        model_part + color_part + size_raw,
        model_part + size_raw,
        model_part + color_part,
        model_part,
    ):
        if candidate and candidate not in candidates:
            candidates.append(candidate)
    return candidates


def _parse_dates(values: pd.Series) -> pd.Series:
    date_part = _text(values).str.extract(
        r"(\d{4}[/-]\d{2}[/-]\d{2})", expand=False
    )
    return pd.to_datetime(
        date_part.str.replace("/", "-", regex=False),
        errors="coerce",
    )


def _add_calendar_columns(frame: pd.DataFrame) -> pd.DataFrame:
    frame = frame.copy()
    frame["연도"] = frame["거래일자"].dt.year.astype("Int64")
    frame["월"] = frame["거래일자"].dt.month.astype("Int64")
    frame["연월"] = frame["거래일자"].dt.strftime("%Y-%m")
    return frame


def prepare_purchase(raw: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    required = [
        "일자-No.",
        "품목코드",
        "품목명(규격)",
        "수량",
        "단가",
        "공급가액",
        "부가세",
        "합계",
        "거래처명",
        "유형",
    ]
    _require_columns(raw, required, "매입 Rawdata")

    frame = pd.DataFrame(
        {
            "거래구분": "매입",
            "거래번호": _text(raw["일자-No."]),
            "거래일자": _parse_dates(raw["일자-No."]),
            "유형": _text(raw["유형"]),
            "표준SKU": raw["품목코드"].map(_purchase_code),
            "모델번호": "",
            "COLOR": "",
            "SIZE": "",
            "품목명": _text(raw["품목명(규격)"]),
            "거래처명": _text(raw["거래처명"]),
            "창고명": "",
            "수량": _numeric(raw["수량"]),
            "단가": _numeric(raw["단가"]),
            "공급가액": _numeric(raw["공급가액"]),
            "부가세": _numeric(raw["부가세"]),
            "합계": _numeric(raw["합계"]),
        }
    )
    frame["금액누락"] = frame[["단가", "공급가액", "부가세", "합계"]].isna().any(axis=1)
    frame["SKU연결상태"] = True
    invalid_dates = int(frame["거래일자"].isna().sum())
    frame = _add_calendar_columns(frame)

    report = {
        "rows": int(len(frame)),
        "invalid_dates": invalid_dates,
        "missing_amount_rows": int(frame["금액누락"].sum()),
        "negative_quantity_rows": int((frame["수량"] < 0).sum()),
        "exact_duplicate_rows": int(raw.duplicated(keep=False).sum()),
        "category_blank_rows": int(frame["유형"].eq("").sum()),
    }
    return frame, report


def prepare_sales(
    raw: pd.DataFrame,
    purchase_codes: set[str],
) -> tuple[pd.DataFrame, dict]:
    required = [
        "일자",
        "판매처명",
        "창고명",
        "품명 및 규격",
        "모델번호",
        "COLOR",
        "SIZE",
        "수량",
        "단가",
        "공급가액",
        "부가세",
        "합 계",
        "유형",
    ]
    _require_columns(raw, required, "매출 Rawdata")

    sales_customer = _text(raw["판매처명"])
    warehouse = _text(raw["창고명"])
    display_customer = sales_customer.mask(sales_customer.eq(""), warehouse)

    selected_skus: list[str] = []
    sku_matches: list[bool] = []
    for model, color, size in zip(
        raw["모델번호"], raw["COLOR"], raw["SIZE"], strict=False
    ):
        candidates = _sales_sku_candidates(model, color, size)
        selected = next(
            (candidate for candidate in candidates if candidate in purchase_codes),
            candidates[0] if candidates else "",
        )
        selected_skus.append(selected)
        sku_matches.append(selected in purchase_codes)

    frame = pd.DataFrame(
        {
            "거래구분": "매출",
            "거래번호": _text(raw["일자"]),
            "거래일자": _parse_dates(raw["일자"]),
            "유형": _text(raw["유형"]),
            "표준SKU": selected_skus,
            "모델번호": _text(raw["모델번호"]),
            "COLOR": _text(raw["COLOR"]),
            "SIZE": _text(raw["SIZE"]),
            "품목명": _text(raw["품명 및 규격"]),
            "거래처명": display_customer,
            "판매처명_원본": sales_customer,
            "창고명": warehouse,
            "수량": _numeric(raw["수량"]),
            "단가": _numeric(raw["단가"]),
            "공급가액": _numeric(raw["공급가액"]),
            "부가세": _numeric(raw["부가세"]),
            "합계": _numeric(raw["합 계"]),
            "SKU연결상태": sku_matches,
        }
    )
    frame["금액누락"] = frame[["단가", "공급가액", "부가세", "합계"]].isna().any(axis=1)
    invalid_dates = int(frame["거래일자"].isna().sum())
    frame = _add_calendar_columns(frame)

    report = {
        "rows": int(len(frame)),
        "invalid_dates": invalid_dates,
        "missing_amount_rows": int(frame["금액누락"].sum()),
        "negative_quantity_rows": int((frame["수량"] < 0).sum()),
        "exact_duplicate_rows": int(raw.duplicated(keep=False).sum()),
        "category_blank_rows": int(frame["유형"].eq("").sum()),
        "original_customer_blank_rows": int(sales_customer.eq("").sum()),
        "display_customer_blank_rows": int(display_customer.eq("").sum()),
        "sku_matched_rows": int(frame["SKU연결상태"].sum()),
        "sku_unmatched_rows": int((~frame["SKU연결상태"]).sum()),
        "sku_match_rate": float(frame["SKU연결상태"].mean()),
    }
    return frame, report


def load_dashboard_data(
    sales_source: ExcelSource,
    purchase_source: ExcelSource,
) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    sales_raw = _read_raw_sheet(sales_source, SALES_SHEET)
    purchase_raw = _read_raw_sheet(purchase_source, PURCHASE_SHEET)

    purchase, purchase_report = prepare_purchase(purchase_raw)
    purchase_codes = set(purchase["표준SKU"].dropna().astype(str))
    sales, sales_report = prepare_sales(sales_raw, purchase_codes)

    quality = {
        "sales": sales_report,
        "purchase": purchase_report,
        "sales_total_quantity": float(sales["수량"].fillna(0).sum()),
        "sales_total_supply": float(sales["공급가액"].fillna(0).sum()),
        "sales_total_tax": float(sales["부가세"].fillna(0).sum()),
        "sales_total_amount": float(sales["합계"].fillna(0).sum()),
        "purchase_total_quantity": float(purchase["수량"].fillna(0).sum()),
        "purchase_total_supply": float(purchase["공급가액"].fillna(0).sum()),
        "purchase_total_tax": float(purchase["부가세"].fillna(0).sum()),
        "purchase_total_amount": float(purchase["합계"].fillna(0).sum()),
    }
    return sales, purchase, quality
