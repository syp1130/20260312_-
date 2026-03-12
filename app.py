# -*- coding: utf-8 -*-
"""재고·발주 자동화 웹 애플리케이션."""
from pathlib import Path
# .env에서 GMAIL_APP_PASSWORD 로드 (가장 먼저 실행)
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parent / ".env")
except ImportError:
    pass
import os
import json
from flask import Flask, request, render_template, jsonify
from werkzeug.utils import secure_filename

from config import DEFAULT_EXCEL_PATH, SENDER_EMAIL
from inventory import (
    load_inventory,
    load_email_template,
    analyze_inventory,
    get_order_summary,
    run_analysis,
    run_analysis_from_items,
)
from email_sender import send_order_email

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB
UPLOAD_FOLDER = Path(__file__).resolve().parent / "uploads"
UPLOAD_FOLDER.mkdir(exist_ok=True)
app.config["UPLOAD_FOLDER"] = str(UPLOAD_FOLDER)
ALLOWED_EXTENSIONS = {"xlsx", "xls"}


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def get_excel_path():
    """사용할 엑셀 파일 경로 (업로드 파일 우선, 없으면 기본)."""
    # 세션/쿠키 대신 간단히 최근 업로드 파일 사용 가능하도록
    default = Path(__file__).resolve().parent / DEFAULT_EXCEL_PATH
    if default.exists():
        return str(default)
    return None


def _serialize(v):
    if hasattr(v, "item"):
        return v.item()
    if hasattr(v, "__float__") and (v != v or str(v) == "nan"):
        return None
    if isinstance(v, (float,)) and (v != v or str(v) == "nan"):
        return None
    return v


@app.route("/")
def index():
    return render_template("index.html", sender_email=SENDER_EMAIL)


@app.route("/api/master-inventory", methods=["GET"])
def api_master_inventory():
    """기준 재고 데이터(엑셀)를 불러와 웹 입력용으로 반환."""
    excel_path = get_excel_path()
    if not excel_path or not Path(excel_path).exists():
        return jsonify({"error": "기본 엑셀 파일(domino_inventory_training.xlsx)이 없습니다."}), 400
    try:
        df = load_inventory(excel_path)
        df.columns = [str(c).strip() for c in df.columns]
        rows = df.to_dict("records")
        out = []
        for row in rows:
            out.append({str(k): _serialize(v) for k, v in row.items()})
        return jsonify({"items": out})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/analyze", methods=["POST"])
def api_analyze():
    """엑셀 파일 업로드, 또는 JSON(웹 입력 재고)으로 재고 분석."""
    # JSON으로 웹 입력 재고 전달된 경우
    if request.content_type and "application/json" in request.content_type:
        data = request.get_json() or {}
        items = data.get("items") or data.get("inventory")
        if items:
            try:
                # 숫자 등 타입 정리
                for row in items:
                    for k in ("현재재고", "안전재고", "MOQ", "부족수량", "발주권장수량", "리드타임(일)"):
                        if k in row and row[k] is not None and row[k] != "":
                            try:
                                row[k] = int(float(row[k]))
                            except (TypeError, ValueError):
                                pass
                inv_list, summary = run_analysis_from_items(items)
                inv_list = [{str(k): _serialize(v) for k, v in r.items()} for r in inv_list]
                order_list = [{str(k): _serialize(v) for k, v in r.items()} for r in summary["order_list"]]
                by_supplier = [{str(k): _serialize(v) for k, v in r.items()} for r in summary["by_supplier"]]
                return jsonify({
                    "inventory": inv_list,
                    "summary": {
                        "total_items": summary["total_items"],
                        "order_items_count": summary["order_items_count"],
                        "total_order_qty": summary["total_order_qty"],
                        "by_supplier": by_supplier,
                        "order_list": order_list,
                    },
                    "excel_path": None,
                })
            except Exception as e:
                return jsonify({"error": str(e)}), 500

    excel_path = None
    if "file" in request.files:
        f = request.files["file"]
        if f and f.filename and allowed_file(f.filename):
            filename = secure_filename(f.filename)
            path = Path(app.config["UPLOAD_FOLDER"]) / filename
            f.save(path)
            excel_path = str(path)
    if not excel_path:
        excel_path = get_excel_path()
    if not excel_path or not Path(excel_path).exists():
        return jsonify({"error": "엑셀 파일을 업로드하거나 기본 파일(domino_inventory_training.xlsx)을 넣어주세요."}), 400
    try:
        analyzed, summary = run_analysis(excel_path)
        inv_list = []
        for _, row in analyzed.iterrows():
            inv_list.append({str(k): _serialize(v) for k, v in row.items()})
        order_list = [{str(k): _serialize(v) for k, v in row.items()} for row in summary["order_list"]]
        by_supplier = [{str(k): _serialize(v) for k, v in row.items()} for row in summary["by_supplier"]]
        return jsonify({
            "inventory": inv_list,
            "summary": {
                "total_items": summary["total_items"],
                "order_items_count": summary["order_items_count"],
                "total_order_qty": summary["total_order_qty"],
                "by_supplier": by_supplier,
                "order_list": order_list,
            },
            "excel_path": excel_path,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/send-orders", methods=["POST"])
def api_send_orders():
    """거래처별로 발주 이메일 발송. JSON으로 order_list/by_supplier 전달 시 엑셀 없이 발송."""
    data = request.get_json() or {}
    order_list = data.get("order_list")
    by_supplier = data.get("by_supplier")
    excel_path = data.get("excel_path") or get_excel_path()

    if order_list is not None and by_supplier is not None:
        # 웹 입력 기준 발송: 템플릿만 엑셀에서 로드
        template_path = get_excel_path()
        if not template_path or not Path(template_path).exists():
            return jsonify({"error": "이메일 템플릿을 위해 기본 엑셀 파일이 필요합니다."}), 400
        try:
            templates = load_email_template(template_path)
            results = []
            for sup in by_supplier:
                supplier_name = sup.get("거래처", "")
                to_email = sup.get("거래처이메일", "")
                if not to_email:
                    results.append({"supplier": supplier_name, "ok": False, "message": "이메일 없음"})
                    continue
                items = [r for r in order_list if r.get("거래처") == supplier_name]
                ok, msg = send_order_email(
                    to_email=to_email,
                    supplier_name=supplier_name,
                    item_list=items,
                    template_subject=templates["subject"],
                    template_body=templates["body"],
                )
                results.append({"supplier": supplier_name, "to": to_email, "ok": ok, "message": msg})
            return jsonify({"results": results})
        except Exception as e:
            return jsonify({"error": str(e)}), 500

    if not excel_path or not Path(excel_path).exists():
        return jsonify({"error": "분석에 사용한 엑셀 경로가 없습니다. 먼저 재고 분석을 실행하세요."}), 400
    try:
        templates = load_email_template(excel_path)
        _, summary = run_analysis(excel_path)
        results = []
        for sup in summary["by_supplier"]:
            supplier_name = sup.get("거래처", "")
            to_email = sup.get("거래처이메일", "")
            if not to_email:
                results.append({"supplier": supplier_name, "ok": False, "message": "이메일 없음"})
                continue
            items = [r for r in summary["order_list"] if r.get("거래처") == supplier_name]
            ok, msg = send_order_email(
                to_email=to_email,
                supplier_name=supplier_name,
                item_list=items,
                template_subject=templates["subject"],
                template_body=templates["body"],
            )
            results.append({"supplier": supplier_name, "to": to_email, "ok": ok, "message": msg})
        return jsonify({"results": results})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
