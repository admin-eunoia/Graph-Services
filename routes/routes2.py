
from flask import Blueprint, request, jsonify, send_file
from io import BytesIO
import base64
import uuid

from Services.excel_live_writer import ExcelLiveWriter
from Services.excel_section_writer import procesar_excel_completo

"""
Flask Blueprint exposing minimal Excel copy/fill endpoints.

Each endpoint accepts JSON (or multipart for `process-in-memory`) and returns
simple JSON responses including a `correlation_id` for tracing.

Endpoints:
- POST /excel/api/copy-template: create a copy of a template registered for a client.
- POST /excel/api/fill-section: fill a key-value section in an existing file.
- POST /excel/api/fill-table: fill a table (multiple rows) in an existing file.
- POST /excel/api/process: process multiple sections in an existing file.
- POST /excel/api/process-in-memory: upload a template (file or base64) and return processed .xlsx.

Examples (narrative):
- copy-template: client `eunoia` has template `bodas`; create a copy named "Boda de Diego".
- fill-table: take file created above and write the guest list table into the `invitados` section.
- process-in-memory: upload a wedding template, send sections config and get back the filled Excel file.
"""

bp = Blueprint("excel_api", __name__, url_prefix="/excel/api")


def _new_cid():
	return str(uuid.uuid4())


@bp.route("/copy-template", methods=["POST"])
def copy_template():
	"""Copy a template into a new file for the client.

	Required JSON params:
	- client_key (str): client identifier (required)
	- dest_file_name (str): destination filename to create (required)

	Optional JSON params:
	- template_key (str): which template to copy (optional; if omitted the active template is used)
	- file_key (str): custom id to store for the created file (optional)
	- context_data (object): metadata to attach when creating the file (optional)

	Example (narrative): client `eunoia` has template `bodas`. Calling this endpoint with
	`dest_file_name` = "Boda de Diego.xlsx" will copy the `bodas` template and create
	a new file named "Boda de Diego.xlsx" registered under `eunoia`.
	"""
	cid = _new_cid()
	try:
		payload = request.get_json(force=True)
		client_key = payload.get("client_key")
		dest_file_name = payload.get("dest_file_name")
		template_key = payload.get("template_key")
		file_key = payload.get("file_key")
		context_data = payload.get("context_data")

		if not client_key or not dest_file_name:
			return jsonify({"error": "client_key and dest_file_name are required", "correlation_id": cid}), 400

		with ExcelLiveWriter(client_key=client_key) as writer:
			item_id, web_url, excel_file_id = writer.copy_template(
				dest_file_name=dest_file_name,
				template_key=template_key,
				file_key=file_key,
				context_data=context_data,
			)

		return jsonify({"message": "OK", "item_id": item_id, "web_url": web_url, "excel_file_id": excel_file_id, "correlation_id": cid})

	except Exception as e:
		return jsonify({"error": str(e), "correlation_id": cid}), 500


@bp.route("/fill-section", methods=["POST"])
def fill_section():
	"""Fill a simple key-value section in an existing file.

	Required JSON params:
	- client_key (str): client identifier (required)
	- file_key (str): the target file identifier (required)
	- datos (object): dict of field_key -> value to write (required)
	- section_key (str): the logical section to fill (optional if file has a single section)

	Example (narrative): after copying the wedding template for `eunoia`, call this
	endpoint with `file_key` set to the created file and `datos` containing
	{"nombre": "Boda de Diego", "fecha": "2026-02-14"} to populate the header row.
	"""
	cid = _new_cid()
	try:
		payload = request.get_json(force=True)
		client_key = payload.get("client_key")
		file_key = payload.get("file_key")
		section_key = payload.get("section_key")
		datos = payload.get("datos")

		if not client_key or not file_key or datos is None:
			return jsonify({"error": "client_key, file_key and datos are required", "correlation_id": cid}), 400

		with ExcelLiveWriter(client_key=client_key) as writer:
			writer.llenar_seccion(file_key=file_key, datos=datos, section_key=section_key)

		return jsonify({"message": "OK", "written_fields": len(datos) if isinstance(datos, dict) else None, "correlation_id": cid})

	except Exception as e:
		return jsonify({"error": str(e), "correlation_id": cid}), 500


@bp.route("/fill-table", methods=["POST"])
def fill_table():
	"""Fill a table section (multiple rows) in an existing file.

	Required JSON params:
	- client_key (str): client identifier (required)
	- file_key (str): the target file identifier (required)
	- datos (array): list of row objects (required)
	- section_key (str): logical section that represents the table (recommended)

	Example (narrative): using the `Boda de Diego.xlsx` file, call this endpoint with
	`section_key` = "invitados" and `datos` = [ {"nombre": "Ana", "asistira": true}, ... ]
	to write the guest list table into the sheet.
	"""
	cid = _new_cid()
	try:
		payload = request.get_json(force=True)
		client_key = payload.get("client_key")
		file_key = payload.get("file_key")
		section_key = payload.get("section_key")
		datos = payload.get("datos")

		if not client_key or not file_key or not isinstance(datos, list):
			return jsonify({"error": "client_key, file_key and datos (list) are required", "correlation_id": cid}), 400

		with ExcelLiveWriter(client_key=client_key) as writer:
			writer.llenar_tabla(file_key=file_key, datos=datos, section_key=section_key)

		return jsonify({"message": "OK", "rows_written": len(datos), "correlation_id": cid})

	except Exception as e:
		return jsonify({"error": str(e), "correlation_id": cid}), 500


@bp.route("/process", methods=["POST"])
def process_excel():
	"""Process multiple sections in an existing file.

	Required JSON params:
	- client_key (str): client identifier (required)
	- file_key (str): the target file identifier (required)
	- secciones (object): { section_key: data } where data is dict or list depending on section (required)

	Example (narrative): process the `Boda de Diego.xlsx` file by sending
	secciones={"header": {"nombre": "Boda de Diego"}, "invitados": [ ... ]}
	which will call the appropriate fill methods for each configured section.
	"""
	cid = _new_cid()
	try:
		payload = request.get_json(force=True)
		client_key = payload.get("client_key")
		file_key = payload.get("file_key")
		secciones = payload.get("secciones")

		if not client_key or not file_key or secciones is None:
			return jsonify({"error": "client_key, file_key and secciones are required", "correlation_id": cid}), 400

		with ExcelLiveWriter(client_key=client_key) as writer:
			writer.procesar_excel(file_key=file_key, secciones=secciones)

		return jsonify({"message": "OK", "processed_sections": len(secciones), "correlation_id": cid})

	except Exception as e:
		return jsonify({"error": str(e), "correlation_id": cid}), 500


@bp.route("/process-in-memory", methods=["POST"])
def process_in_memory():
	"""Process a template in-memory and return the filled Excel file.

	Two options to provide the template:
	1) Multipart/form-data upload:
		- field `template` (file): the .xlsx template file (required)
		- field `secciones` (string): JSON string of sections (required)
		- field `configuracion` (string): JSON string describing markers/columns (required)

	2) JSON body with base64:
		- `template_b64` (str): base64 encoded .xlsx file (required if not multipart)
		- `secciones` (object): same as above (required)
		- `configuracion` (object): same as above (required)

	Example (narrative): upload the `bodas.xlsx` template and provide a `configuracion`
	that maps section keys to markers/columns. The endpoint returns the processed
	.xlsx file (attachment), e.g. a filled "Boda de Diego.xlsx" file.
	"""

	cid = _new_cid()
	try:
		# Support multipart file upload (field 'template') or JSON base64 (template_b64)
		secciones = None
		configuracion = None

		if request.content_type and request.content_type.startswith("multipart/"):
			template_file = request.files.get("template")
			secciones_raw = request.form.get("secciones")
			configuracion_raw = request.form.get("configuracion")

			if template_file:
				template_bytes = template_file.read()
			else:
				return jsonify({"error": "template file not provided", "correlation_id": cid}), 400

			if secciones_raw:
				import json
				secciones = json.loads(secciones_raw)
			if configuracion_raw:
				import json
				configuracion = json.loads(configuracion_raw)

		else:
			payload = request.get_json(force=True)
			template_b64 = payload.get("template_b64")
			secciones = payload.get("secciones")
			configuracion = payload.get("configuracion")

			if not template_b64:
				return jsonify({"error": "template_b64 is required when not uploading multipart file", "correlation_id": cid}), 400

			template_bytes = base64.b64decode(template_b64)

		if secciones is None or configuracion is None:
			return jsonify({"error": "secciones and configuracion are required", "correlation_id": cid}), 400

		# Process in-memory and return file bytes
		output_io = procesar_excel_completo(template_bytes=template_bytes, secciones=secciones, configuracion=configuracion)

		output_io.seek(0)
		filename = "processed.xlsx"
		return send_file(output_io, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name=filename)

	except Exception as e:
		return jsonify({"error": str(e), "correlation_id": cid}), 500

