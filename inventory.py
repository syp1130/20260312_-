# -*- coding: utf-8 -*-
"""재고 분석 및 발주 권장 수량 계산 로직."""
import pandas as pd
from pathlib import Path


def load_inventory(excel_path: str) -> pd.DataFrame:
    """엑셀에서 Inventory 시트를 읽어 DataFrame 반환."""
    df = pd.read_excel(excel_path, sheet_name="Inventory", header=0)
    return df


def load_suppliers(excel_path: str) -> pd.DataFrame:
    """엑셀에서 Suppliers 시트 읽기."""
    df = pd.read_excel(excel_path, sheet_name="Suppliers", header=0)
    return df


def load_email_template(excel_path: str) -> dict:
    """엑셀에서 이메일 템플릿(제목/본문) 읽기."""
    df = pd.read_excel(excel_path, sheet_name="EmailTemplate", header=None)
    subject = ""
    body = ""
    for _, row in df.iterrows():
        v0, v1 = str(row[0]) if pd.notna(row[0]) else "", str(row[1]) if pd.notna(row[1]) else ""
        if "제목" in v0 and v1:
            subject = v1.strip()
        if "본문" in v0 and v1:
            body = v1.strip()
    return {"subject": subject, "body": body}


def analyze_inventory(df: pd.DataFrame) -> pd.DataFrame:
    """
    재고 시트를 분석하여 부족수량, 발주권장수량, 상태, 담당자알림메시지를 계산합니다.
    기준: 현재재고 < 안전재고 → 발주 필요
    발주 권장 수량: MAX(MOQ, 안전재고 - 현재재고)
    """
    df = df.copy()
    # 컬럼명 정리 (공백 제거)
    df.columns = [str(c).strip() for c in df.columns]

    current_col = "현재재고"
    safety_col = "안전재고"
    moq_col = "MOQ"
    unit_col = "단위"
    name_col = "재료명"

    if current_col not in df.columns or safety_col not in df.columns:
        return df

    current = pd.to_numeric(df[current_col], errors="coerce").fillna(0)
    safety = pd.to_numeric(df[safety_col], errors="coerce").fillna(0)
    moq = pd.to_numeric(df[moq_col], errors="coerce").fillna(0)

    shortage = (safety - current).clip(lower=0).astype(int)
    order_qty = shortage.copy()
    # 발주권장 = MAX(MOQ, 안전재고 - 현재재고), 단 부족할 때만
    need_order = current < safety
    order_qty = order_qty.where(~need_order, order_qty.where(order_qty >= moq, moq).astype(int))

    df["부족수량"] = shortage
    df["발주권장수량"] = order_qty.where(need_order, 0).astype(int)
    df["상태"] = need_order.map({True: "발주 필요", False: "정상"})

    def alert_msg(row):
        if row["상태"] != "발주 필요":
            return ""
        u = row.get(unit_col, "개")
        return (
            f"{row[name_col]} 재고 부족 - 현재 {int(row[current_col])}{u}, "
            f"안전재고 {int(row[safety_col])}{u}, 권장발주 {int(row['발주권장수량'])}{u}"
        )

    df["담당자알림메시지"] = df.apply(alert_msg, axis=1)
    return df


def get_order_summary(analyzed: pd.DataFrame) -> dict:
    """발주 요약 통계."""
    need = analyzed[analyzed["상태"] == "발주 필요"]
    total_items = len(analyzed)
    order_items_count = len(need)
    total_order_qty = need["발주권장수량"].sum()
    by_supplier = (
        need.groupby("거래처", as_index=False)
        .agg(
            발주품목수=("재료명", "count"),
            총권장발주수량=("발주권장수량", "sum"),
        )
    )
    # 거래처이메일, 리드타임 병합
    if "거래처이메일" in need.columns and "리드타임(일)" in need.columns:
        first = need.groupby("거래처").first().reset_index()
        merge = by_supplier.merge(
            first[["거래처", "거래처이메일", "리드타임(일)"]],
            on="거래처",
            how="left",
        )
        by_supplier = merge
    return {
        "total_items": int(total_items),
        "order_items_count": int(order_items_count),
        "total_order_qty": int(total_order_qty),
        "by_supplier": by_supplier.to_dict("records"),
        "order_list": need.to_dict("records"),
    }


def run_analysis(excel_path: str) -> tuple[pd.DataFrame, dict]:
    """엑셀 경로로 재고 분석 후 (분석된 DataFrame, 발주 요약) 반환."""
    df = load_inventory(excel_path)
    analyzed = analyze_inventory(df)
    summary = get_order_summary(analyzed)
    return analyzed, summary


def run_analysis_from_items(items: list[dict]) -> tuple[list[dict], dict]:
    """웹에서 입력한 재고 리스트로 분석. (분석된 행 리스트, 발주 요약) 반환."""
    if not items:
        return [], {
            "total_items": 0,
            "order_items_count": 0,
            "total_order_qty": 0,
            "by_supplier": [],
            "order_list": [],
        }
    df = pd.DataFrame(items)
    analyzed = analyze_inventory(df)
    summary = get_order_summary(analyzed)
    # JSON 직렬화 가능한 리스트로 변환
    inv_list = analyzed.fillna("").astype(object).where(analyzed.notna(), None)
    inv_list = inv_list.to_dict("records")
    for row in inv_list:
        for k, v in list(row.items()):
            if hasattr(v, "item"):
                row[k] = v.item()
            elif isinstance(v, (float,)) and (v != v or str(v) == "nan"):
                row[k] = None
    return inv_list, summary
