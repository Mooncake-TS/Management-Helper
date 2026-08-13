from __future__ import annotations

import io
import re
import unicodedata
from functools import lru_cache
from pathlib import Path
from typing import BinaryIO

import pandas as pd


ExcelSource = str | Path | bytes | bytearray | BinaryIO

TAG_MASTER_PATH = Path(__file__).with_name("tag_master.csv")

SALES_REQUIRED = [
    "일자",
    "판매처명",
    "창고명",
    "품명및규격",
    "모델번호",
    "COLOR",
    "SIZE",
    "수량",
    "단가",
    "공급가액",
    "부가세",
    "합계",
]

PURCHASE_REQUIRED = [
    "일자-No.",
    "품목코드",
    "품목명(규격)",
    "수량",
    "단가",
    "공급가액",
    "부가세",
    "합계",
    "거래처명",
]


def _excel_source(source: ExcelSource) -> ExcelSource:
    if isinstance(source, (bytes, bytearray)):
        return io.BytesIO(source)
    return source


def _clean_header(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", "", str(value)).strip()


def _read_erp_sheet(
    source: ExcelSource,
    required: list[str],
    source_name: str,
) -> pd.DataFrame:
    excel = pd.ExcelFile(_excel_source(source), engine="openpyxl")

    for sheet_name in excel.sheet_names:
        raw = pd.read_excel(
            excel,
            sheet_name=sheet_name,
            header=None,
            dtype=object,
        )
        for header_row in range(min(30, len(raw))):
            headers = [_clean_header(value) for value in raw.iloc[header_row].tolist()]
            if not all(column in headers for column in required):
                continue

            frame = raw.iloc[header_row + 1 :].copy()
            keep_indices = [index for index, column in enumerate(headers) if column]
            frame = frame.iloc[:, keep_indices]
            frame.columns = [headers[index] for index in keep_indices]
            frame = frame.loc[:, ~frame.columns.duplicated()]
            return frame.dropna(how="all").reset_index(drop=True)

    raise ValueError(
        f"{source_name}에서 필요한 헤더를 찾지 못했습니다: {', '.join(required)}"
    )


def _require_columns(frame: pd.DataFrame, required: list[str], source_name: str) -> None:
    missing = [column for column in required if column not in frame.columns]
    if missing:
        raise ValueError(f"{source_name}에 필요한 컬럼이 없습니다: {', '.join(missing)}")


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
        r"(\d{4}[/-]\d{1,2}[/-]\d{1,2})", expand=False
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


def _name_key(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    normalized = unicodedata.normalize("NFKC", str(value))
    return re.sub(r"\s+", " ", normalized).strip().casefold()


@lru_cache(maxsize=1)
def _tag_lookup() -> dict[str, str]:
    if not TAG_MASTER_PATH.exists():
        raise FileNotFoundError(f"태그 마스터 파일을 찾지 못했습니다: {TAG_MASTER_PATH.name}")
    master = pd.read_csv(TAG_MASTER_PATH, encoding="utf-8-sig", dtype=str).fillna("")
    _require_columns(master, ["품명", "태그"], "태그 마스터")
    return {
        _name_key(name): str(tag).strip()
        for name, tag in zip(master["품명"], master["태그"], strict=False)
        if _name_key(name) and str(tag).strip()
    }


def _clean_display_size(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return "" if text.upper() in {"", "0", "000", "NAN", "NONE"} else text


def _apply_tags(
    names: pd.Series,
    sizes: pd.Series | None = None,
) -> tuple[pd.Series, pd.Series, pd.Series, pd.Series]:
    lookup = _tag_lookup()
    matched_tags = names.map(lambda value: lookup.get(_name_key(value), ""))
    matched = matched_tags.ne("")
    tags = matched_tags.mask(~matched, "기타용품 > 기타용품").astype(str)

    if sizes is not None:
        dynamic_size = sizes.map(_clean_display_size)
        dynamic_mask = tags.str.endswith("> SIZE 열 사용")
        categories = tags.str.split(">", n=1).str[0].str.strip()
        replacements = categories + " > " + dynamic_size.mask(dynamic_size.eq(""), categories)
        tags = tags.mask(dynamic_mask, replacements)

    split_tags = tags.str.split(">", n=1, expand=True)
    categories = split_tags[0].str.strip()
    subcategories = (
        split_tags[1].str.strip() if split_tags.shape[1] > 1 else categories.copy()
    )
    return tags, categories, subcategories, matched


def _detail_rows(raw: pd.DataFrame, date_column: str, name_column: str) -> pd.DataFrame:
    date_labels = _text(raw[date_column])
    names = _text(raw[name_column])
    subtotal_or_total = date_labels.str.contains("계", na=False)
    return raw.loc[names.ne("") & ~subtotal_or_total].copy().reset_index(drop=True)


def prepare_purchase(raw: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    _require_columns(raw, PURCHASE_REQUIRED, "매입 파일")
    detail = _detail_rows(raw, "일자-No.", "품목명(규격)")

    names = _text(detail["품목명(규격)"])
    tags, categories, subcategories, tag_matched = _apply_tags(names)
    frame = pd.DataFrame(
        {
            "거래구분": "매입",
            "거래번호": _text(detail["일자-No."]),
            "거래일자": _parse_dates(detail["일자-No."]),
            "태그": tags,
            "유형": categories,
            "소분류": subcategories,
            "표준SKU": detail["품목코드"].map(_purchase_code),
            "모델번호": "",
            "COLOR": "",
            "SIZE": "",
            "품목명": names,
            "거래처명": _text(detail["거래처명"]),
            "창고명": "",
            "수량": _numeric(detail["수량"]),
            "단가": _numeric(detail["단가"]),
            "공급가액": _numeric(detail["공급가액"]),
            "부가세": _numeric(detail["부가세"]),
            "합계": _numeric(detail["합계"]),
            "태그연결상태": tag_matched,
        }
    )
    frame["금액공란"] = frame[["단가", "공급가액", "부가세", "합계"]].isna().any(axis=1)
    invalid_dates = int(frame["거래일자"].isna().sum())
    frame = _add_calendar_columns(frame)

    report = {
        "rows": int(len(frame)),
        "invalid_dates": invalid_dates,
        "missing_amount_rows": int(frame["금액공란"].sum()),
        "negative_quantity_rows": int((frame["수량"] < 0).sum()),
        "exact_duplicate_rows": int(detail.duplicated(keep=False).sum()),
        "category_blank_rows": int(frame["유형"].eq("").sum()),
        "tag_matched_rows": int(frame["태그연결상태"].sum()),
        "tag_unmatched_rows": int((~frame["태그연결상태"]).sum()),
    }
    return frame, report


def prepare_sales(
    raw: pd.DataFrame,
    purchase_codes: set[str],
    source_label: str,
) -> tuple[pd.DataFrame, dict]:
    _require_columns(raw, SALES_REQUIRED, source_label)
    detail = _detail_rows(raw, "일자", "품명및규격")

    sales_customer = _text(detail["판매처명"])
    warehouse = _text(detail["창고명"])
    display_customer = sales_customer.mask(sales_customer.eq(""), warehouse)
    names = _text(detail["품명및규격"])
    sizes = _text(detail["SIZE"])
    tags, categories, subcategories, tag_matched = _apply_tags(names, sizes)

    selected_skus: list[str] = []
    sku_matches: list[bool] = []
    for model, color, size in zip(
        detail["모델번호"], detail["COLOR"], detail["SIZE"], strict=False
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
            "거래번호": _text(detail["일자"]),
            "거래일자": _parse_dates(detail["일자"]),
            "태그": tags,
            "유형": categories,
            "소분류": subcategories,
            "표준SKU": selected_skus,
            "모델번호": _text(detail["모델번호"]),
            "COLOR": _text(detail["COLOR"]),
            "SIZE": sizes,
            "품목명": names,
            "거래처명": display_customer,
            "판매처명_원본": sales_customer,
            "창고명": warehouse,
            "수량": _numeric(detail["수량"]),
            "단가": _numeric(detail["단가"]),
            "공급가액": _numeric(detail["공급가액"]),
            "부가세": _numeric(detail["부가세"]),
            "합계": _numeric(detail["합계"]),
            "SKU연결상태": sku_matches,
            "태그연결상태": tag_matched,
            "원본파일": source_label,
        }
    )
    frame["금액공란"] = frame[["단가", "공급가액", "부가세", "합계"]].isna().any(axis=1)
    invalid_dates = int(frame["거래일자"].isna().sum())
    frame = _add_calendar_columns(frame)

    report = {
        "source": source_label,
        "rows": int(len(frame)),
        "invalid_dates": invalid_dates,
        "missing_amount_rows": int(frame["금액공란"].sum()),
        "negative_quantity_rows": int((frame["수량"] < 0).sum()),
        "exact_duplicate_rows": int(detail.duplicated(keep=False).sum()),
        "category_blank_rows": int(frame["유형"].eq("").sum()),
        "original_customer_blank_rows": int(sales_customer.eq("").sum()),
        "display_customer_blank_rows": int(display_customer.eq("").sum()),
        "sku_matched_rows": int(frame["SKU연결상태"].sum()),
        "sku_unmatched_rows": int((~frame["SKU연결상태"]).sum()),
        "tag_matched_rows": int(frame["태그연결상태"].sum()),
        "tag_unmatched_rows": int((~frame["태그연결상태"]).sum()),
    }
    report["sku_match_rate"] = (
        report["sku_matched_rows"] / report["rows"] if report["rows"] else 0.0
    )
    return frame, report


def _combine_sales_reports(reports: list[dict]) -> dict:
    sum_keys = [
        "rows",
        "invalid_dates",
        "missing_amount_rows",
        "negative_quantity_rows",
        "exact_duplicate_rows",
        "category_blank_rows",
        "original_customer_blank_rows",
        "display_customer_blank_rows",
        "sku_matched_rows",
        "sku_unmatched_rows",
        "tag_matched_rows",
        "tag_unmatched_rows",
    ]
    combined = {key: sum(int(report.get(key, 0)) for report in reports) for key in sum_keys}
    combined["sku_match_rate"] = (
        combined["sku_matched_rows"] / combined["rows"] if combined["rows"] else 0.0
    )
    combined["files"] = reports
    return combined


def load_dashboard_data(
    previous_sales_source: ExcelSource,
    current_sales_source: ExcelSource,
    purchase_source: ExcelSource,
) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    previous_sales_raw = _read_erp_sheet(
        previous_sales_source,
        SALES_REQUIRED,
        "전년도 매출 파일",
    )
    current_sales_raw = _read_erp_sheet(
        current_sales_source,
        SALES_REQUIRED,
        "올해 매출 파일",
    )
    purchase_raw = _read_erp_sheet(
        purchase_source,
        PURCHASE_REQUIRED,
        "매입 파일",
    )

    purchase, purchase_report = prepare_purchase(purchase_raw)
    purchase_codes = set(purchase["표준SKU"].dropna().astype(str))
    previous_sales, previous_report = prepare_sales(
        previous_sales_raw,
        purchase_codes,
        "전년도 매출",
    )
    current_sales, current_report = prepare_sales(
        current_sales_raw,
        purchase_codes,
        "올해 매출",
    )
    sales = pd.concat([previous_sales, current_sales], ignore_index=True)
    sales_report = _combine_sales_reports([previous_report, current_report])

    quality = {
        "sales": sales_report,
        "purchase": purchase_report,
        "sales_year_counts": {
            int(year): int(count)
            for year, count in sales["연도"].dropna().astype(int).value_counts().sort_index().items()
        },
        "purchase_year_counts": {
            int(year): int(count)
            for year, count in purchase["연도"].dropna().astype(int).value_counts().sort_index().items()
        },
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
