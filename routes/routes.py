from flask import Blueprint, request, jsonify
from Services.graph_services import GraphServices

graph_bp = Blueprint("graph", __name__)

@graph_bp.post("/copy_excel")
def copy_excel():
    data = request.get_json() or {}
    required = ["original_name", "copy_name", "folder_path"]
    faltan = [k for k in required if k not in data]
    if faltan:
        return jsonify({"error": f"Faltan campos: {', '.join(faltan)}"}), 400
    try:
        gs = GraphServices()
        gs.copy_excel(data["original_name"], data["copy_name"], data["folder_path"])
        return jsonify({"message": "Excel copiado exitosamente"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@graph_bp.post("/fill_excel")
def fill_excel():
    data = request.get_json() or {}
    required = ["file_path", "worksheet_name", "data"]
    faltan = [k for k in required if k not in data]
    if faltan:
        return jsonify({"error": f"Faltan campos: {', '.join(faltan)}"}), 400
    try:
        gs = GraphServices()
        gs.fill_excel(data["file_path"], data["worksheet_name"], data["data"])
        return jsonify({"message": "Excel llenado exitosamente"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500