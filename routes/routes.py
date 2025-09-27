# routes/routes.py
import re
import time
import uuid
from datetime import datetime
from string import Formatter
from flask import Blueprint, request, jsonify
from Postgress.Tables import (
    TenantCredentials,
    TenantUsers,
    StorageTargets,
    Templates,
    RenderLogs,
    RenderStatus,
)
from Auth.Microsoft_Graph_Auth import MicrosoftGraphAuthenticator
from Services.graph_services import GraphServices, GraphAPIError, validate_cell_map_or_raise

MAX_CLIENT_FIELD_LENGTH = 100  # coincide con tamaño en DB
MAX_TEMPLATE_KEY_LENGTH = 100
MAX_DEST_FILE_NAME_LENGTH = 255  # límite en RenderLogs.dest_file_name
MAX_DATA_ENTRIES = 500  # evita payloads que saturen Graph
MAX_TARGET_ALIAS_LENGTH = 100
MAX_DATA_VALUE_LENGTH = 2048  # chars máximos por valor string
MAX_DATA_TOTAL_LENGTH = 20000  # suma máxima de caracteres representados
MAX_NAMING_ENTRIES = 30
MAX_NAMING_KEY_LENGTH = 64
MAX_NAMING_VALUE_LENGTH = 512
ALLOWED_IDENTIFIER_RE = re.compile(r"^[A-Za-z0-9._-]+$")

graph_bp = Blueprint("graph", __name__)

# ------------------ Helpers ------------------ #
def _require(body, keys):
    for k in keys:
        if k not in body:
            return f"Falta campo requerido: {k}"
    return None

def _new_correlation_id():
    return str(uuid.uuid4())


def _validate_string_field(field_name: str, value, *, max_length: int):
    if not isinstance(value, str) or not value.strip():
        return f"{field_name} debe ser un string no vacío"
    if len(value.strip()) > max_length:
        return f"{field_name} excede longitud máxima ({max_length})"
    return None


def _validate_data_dict(data):
    if not isinstance(data, dict):
        return "data debe ser un diccionario"
    if not data:
        return "data no puede estar vacío"
    if len(data) > MAX_DATA_ENTRIES:
        return f"data excede el máximo de {MAX_DATA_ENTRIES} celdas"
    total_chars = 0
    for key, value in data.items():
        if isinstance(value, str):
            if len(value) > MAX_DATA_VALUE_LENGTH:
                return f"Valor de '{key}' excede {MAX_DATA_VALUE_LENGTH} caracteres"
            total_chars += len(value)
        elif isinstance(value, (int, float, bool)) or value is None:
            total_chars += len(str(value))
        else:
            return f"Tipo no soportado para '{key}': {type(value).__name__}"
        if total_chars > MAX_DATA_TOTAL_LENGTH:
            return f"Suma total de caracteres en data excede {MAX_DATA_TOTAL_LENGTH}"
    return None


def _pattern_fields(pattern: str) -> set[str]:
    formatter = Formatter()
    return {
        field
        for literal, field, format_spec, conversion in formatter.parse(pattern)
        if field
    }


def _build_dest_file_name(template, body) -> str:
    """Construye el nombre final de archivo a partir del patrón configurable.

    Soporta placeholders dinamicos definidos en DB (dest_file_pattern) y valores
    enviados en el payload bajo `naming`. Permite cambiar formatos por tenant sin
    redeploy mientras se mantenga la sanitización final.
    """
    pattern = (template.dest_file_pattern or "{template_key}_{client_name_sanitized}.xlsx").strip()
    if not pattern:
        pattern = "{template_key}_{client_name_sanitized}.xlsx"

    naming = body.get("naming", {})
    if naming is None:
        naming = {}
    if not isinstance(naming, dict):
        raise ValueError("naming debe ser un objeto (dict)")

    raw_client_name = str(body["client_name"]).strip()
    client_name_sanitized = re.sub(r"[^A-Za-z0-9._-]+", "_", raw_client_name).strip("_")
    if not client_name_sanitized:
        client_name_sanitized = "cliente"

    context = {
        "client_key": body["client_key"],
        "template_key": template.template_key,
        "client_name": raw_client_name,
        "client_name_sanitized": client_name_sanitized,
    }
    # Permite que naming sobrescriba context si se requiere algo distinto
    context.update(naming)

    required = _pattern_fields(pattern)
    missing = [k for k in required if k not in context]
    if missing:
        raise ValueError(f"Faltan campos para naming: {', '.join(missing)}")

    try:
        raw_name = pattern.format(**context)
    except Exception as exc:
        raise ValueError(f"dest_file_pattern inválido: {exc}") from exc

    sanitized = re.sub(r"[^A-Za-z0-9._-]+", "_", str(raw_name)).strip("_")
    if not sanitized:
        sanitized = "archivo"
    if not sanitized.lower().endswith(".xlsx"):
        sanitized = f"{sanitized}.xlsx"
    if len(sanitized) > MAX_DEST_FILE_NAME_LENGTH:
        sanitized = sanitized[:MAX_DEST_FILE_NAME_LENGTH]
        sanitized = sanitized.rstrip("._-") or "archivo.xlsx"
        if not sanitized.lower().endswith(".xlsx"):
            sanitized = f"{sanitized}.xlsx"
    if any(sep in sanitized for sep in ("/", "\\")) or ".." in sanitized:
        raise ValueError("Nombre de archivo resultante inválido")
    return sanitized


def _validate_naming_dict(naming):
    if naming in (None, {}):
        return None, {}
    if not isinstance(naming, dict):
        return "naming debe ser un objeto (dict)", {}
    if len(naming) > MAX_NAMING_ENTRIES:
        return f"naming excede el máximo de {MAX_NAMING_ENTRIES} claves", {}

    sanitized = {}
    for raw_key, raw_value in naming.items():
        if not isinstance(raw_key, str) or not raw_key.strip():
            return "naming contiene claves inválidas", {}
        key = raw_key.strip()
        if len(key) > MAX_NAMING_KEY_LENGTH:
            return f"Clave '{key}' excede {MAX_NAMING_KEY_LENGTH} caracteres", {}
        if not ALLOWED_IDENTIFIER_RE.match(key):
            return f"Clave '{key}' contiene caracteres no permitidos", {}

        value = raw_value
        if isinstance(value, str):
            value = value.strip()
            if len(value) > MAX_NAMING_VALUE_LENGTH:
                return f"Valor de '{key}' excede {MAX_NAMING_VALUE_LENGTH} caracteres", {}
        elif not isinstance(value, (int, float, bool)) and value is not None:
            return f"Tipo no soportado para '{key}' en naming", {}

        sanitized[key] = value

    return None, sanitized


# ------------------ ROUTES ------------------ #

@graph_bp.post("/excel/render-upload")
def render_upload():
    """
    1) Lee tenant/template/target de DB
    2) Descarga template desde OneDrive del cliente
    3) Rellena en memoria con 'data'
    4) Sube el archivo final al destino (conflictBehavior del template)
    5) Inserta log en render_logs
    """
    body = request.get_json() or {}
    err = _require(body, ["client_key", "template_key", "data", "client_name"])
    if err:
        return jsonify({"error": err}), 400

    err = _validate_string_field("client_key", body.get("client_key"), max_length=MAX_CLIENT_FIELD_LENGTH)
    if err:
        return jsonify({"error": err}), 400

    err = _validate_string_field("template_key", body.get("template_key"), max_length=MAX_TEMPLATE_KEY_LENGTH)
    if err:
        return jsonify({"error": err}), 400

    err = _validate_string_field("client_name", body.get("client_name"), max_length=MAX_CLIENT_FIELD_LENGTH)
    if err:
        return jsonify({"error": err}), 400

    err = _validate_data_dict(body.get("data"))
    if err:
        return jsonify({"error": err}), 400

    try:
        validate_cell_map_or_raise(body["data"])
    except ValueError as ve:
        return jsonify({"error": str(ve)}), 400

    naming_err, sanitized_naming = _validate_naming_dict(body.get("naming"))
    if naming_err:
        return jsonify({"error": naming_err}), 400
    body["naming"] = sanitized_naming

    target_alias = body.get("target_alias")
    if target_alias is not None:
        err = _validate_string_field("target_alias", target_alias, max_length=MAX_TARGET_ALIAS_LENGTH)
        if err:
            return jsonify({"error": err}), 400
        alias_clean = target_alias.strip()
        if alias_clean and not ALLOWED_IDENTIFIER_RE.match(alias_clean):
            return jsonify({"error": "target_alias contiene caracteres no permitidos"}), 400
        body["target_alias"] = alias_clean or None

    target_label = body.get("target_label")
    if target_label is not None:
        err = _validate_string_field("target_label", target_label, max_length=MAX_TARGET_ALIAS_LENGTH)
        if err:
            return jsonify({"error": err}), 400
        label_clean = target_label.strip()
        if label_clean and not ALLOWED_IDENTIFIER_RE.match(label_clean):
            return jsonify({"error": "target_label contiene caracteres no permitidos"}), 400
        body["target_label"] = label_clean or None

    corr_id = _new_correlation_id()
    t0 = time.perf_counter()

    # Sesión de DB provista por el middleware (main.py)
    db = request.environ.get("db_session")

    dest_file_name = None
    template = None

    try:
        # 1) Config del tenant
        creds = (
            db.query(TenantCredentials)
            .filter_by(client_key=body["client_key"], enabled=True)
            .first()
        )
        storage_query = (
            db.query(StorageTargets)
            .filter_by(client_key=body["client_key"], tenant_id=creds.id)
        )
        label = body.get("target_label")
        target_alias = body.get("target_alias")
        tenant_user = None
        if target_alias:
            tenant_user = (
                db.query(TenantUsers)
                .filter_by(tenant_id=creds.id, alias=target_alias)
                .first()
            )
            if not tenant_user:
                return jsonify({"error": f"target_alias '{target_alias}' no está configurado"}), 400
            storage = storage_query.filter_by(tenant_user_id=tenant_user.id).first()
            if not storage:
                return jsonify({"error": f"No hay destino configurado para el alias '{target_alias}'"}), 400
        else:
            if label:
                storage = storage_query.filter_by(label=label).first()
            else:
                storage = storage_query.first()

        template = (
            db.query(Templates)
            .filter_by(client_key=body["client_key"], template_key=body["template_key"])
            .first()
        )
        if not creds or not storage or not template:
            return jsonify({"error": "Configuración incompleta en DB"}), 400

        # 2) Token MSAL
        auth = MicrosoftGraphAuthenticator(
            creds.tenant_id, creds.app_client_id, creds.app_client_secret
        )
        token = auth.get_access_token()
        gs = GraphServices(access_token=token, correlation_id=corr_id)

        # 3) Descargar template
        full_template_path = f"{template.template_folder_path.strip('/')}/{template.template_file_name}"
        tpl_bytes, ms_id_download = gs.download_file_bytes(
            full_template_path,
            target_user_id=storage.target_user_id if not storage.use_drive_id else None,
            drive_id=storage.drive_id if storage.use_drive_id else None,
        )

        # 4) Render en memoria
        filled_bytes = gs.render_in_memory(tpl_bytes, body["data"])

        # 5) Calcular nombre de archivo final según patrón configurable
        try:
            dest_file_name = _build_dest_file_name(template, body)
        except ValueError as ve:
            return jsonify({"error": str(ve)}), 400

        dest_path = f"{storage.default_dest_folder_path.strip('/')}/{dest_file_name}"
        upload_result, ms_id_upload = gs.upload_file_bytes(
            filled_bytes,
            dest_path,
            conflict_behavior=template.default_conflict_behavior,
            target_user_id=storage.target_user_id if not storage.use_drive_id else None,
            drive_id=storage.drive_id if storage.use_drive_id else None,
        )

        # 6) Guardar log
        duration_ms = int((time.perf_counter() - t0) * 1000)
        
        log_row = RenderLogs(
            client_key=body["client_key"],
            template_id=template.id,
            template_key=template.template_key,
            data_json=body["data"],
            result_drive_item_id=upload_result.get("id"),
            result_web_url=upload_result.get("webUrl"),
            status=RenderStatus.SUCCESS,
            requested_by=body.get("requested_by", "eco-agent"),
            created_at=datetime.utcnow(),
            duration_ms=duration_ms,
            dest_file_name=dest_file_name,
        )

        db.add(log_row)
        # commit/rollback lo maneja tu teardown; si prefieres commit explícito, descomenta:
        # db.commit()

        return jsonify(
            {
                "message": "OK",
                "driveItem": upload_result,
                "correlation_id": corr_id,
                "duration_ms": duration_ms,
                "ms_request_ids": {
                    "download_template": ms_id_download,
                    "upload_file": ms_id_upload,
                },
            }
        ), 200

    except GraphAPIError as ge:
        try:
            duration_ms = int((time.perf_counter() - t0) * 1000)
            log_row = RenderLogs(
                client_key=body.get("client_key"),
                template_id=template.id if template else None,
                template_key=body.get("template_key", "__unknown_template__"),
                data_json=body.get("data", {}),
                status=RenderStatus.ERROR,
                error_message=f"{ge.message} | ms-request-id={ge.ms_request_id}",
                requested_by=body.get("requested_by", "eco-agent"),
                created_at=datetime.utcnow(),
                duration_ms=duration_ms,
                dest_file_name=dest_file_name,
            )
            db.add(log_row)
        except Exception:
            pass

        status = ge.status_code if 400 <= (ge.status_code or 0) <= 599 else 502
        payload = {
            "error": ge.message,
            "correlation_id": corr_id,
        }
        if ge.ms_request_id:
            payload["ms_request_id"] = ge.ms_request_id
        return jsonify(payload), status

    except Exception as e:
        try:
            duration_ms = int((time.perf_counter() - t0) * 1000)
            log_row = RenderLogs(
                client_key=body.get("client_key"),
                template_id=template.id if template else None,
                template_key=body.get("template_key", "__unknown_template__"),
                data_json=body.get("data", {}),
                status=RenderStatus.ERROR,
                error_message=str(e),
                requested_by=body.get("requested_by", "eco-agent"),
                created_at=datetime.utcnow(),
                duration_ms=duration_ms,
                dest_file_name=dest_file_name,
            )
            db.add(log_row)
        except Exception:
            pass

        return jsonify({"error": str(e), "correlation_id": corr_id}), 500


@graph_bp.post("/excel/write-cells")
def write_cells():
    """
    1) Lee tenant/target de DB
    2) Valida el mapa de celdas (A1 / Hoja!A1)
    3) Escribe directamente en el Excel existente vía Graph Excel API
    4) Inserta log en render_logs (sin template_id)
    """
    body = request.get_json() or {}
    err = _require(body, ["client_key", "dest_file_name", "data"])
    if err:
        return jsonify({"error": err}), 400

    err = _validate_string_field("client_key", body.get("client_key"), max_length=MAX_CLIENT_FIELD_LENGTH)
    if err:
        return jsonify({"error": err}), 400

    err = _validate_string_field("dest_file_name", body.get("dest_file_name"), max_length=MAX_DEST_FILE_NAME_LENGTH)
    if err:
        return jsonify({"error": err}), 400

    err = _validate_data_dict(body.get("data"))
    if err:
        return jsonify({"error": err}), 400

    # Validación de direcciones y valores
    try:
        validate_cell_map_or_raise(body["data"])
    except ValueError as ve:
        return jsonify({"error": str(ve)}), 400

    dest_file_name = body["dest_file_name"].strip()
    if any(sep in dest_file_name for sep in ("/", "\\")) or dest_file_name.startswith(".") or ".." in dest_file_name:
        return jsonify({"error": "dest_file_name inválido"}), 400
    if not dest_file_name.lower().endswith(".xlsx"):
        return jsonify({"error": "dest_file_name debe terminar en .xlsx"}), 400
    body["dest_file_name"] = dest_file_name

    target_alias = body.get("target_alias")
    if target_alias is not None:
        err = _validate_string_field("target_alias", target_alias, max_length=MAX_TARGET_ALIAS_LENGTH)
        if err:
            return jsonify({"error": err}), 400
        alias_clean = target_alias.strip()
        if alias_clean and not ALLOWED_IDENTIFIER_RE.match(alias_clean):
            return jsonify({"error": "target_alias contiene caracteres no permitidos"}), 400
        body["target_alias"] = alias_clean or None

    target_label = body.get("target_label")
    if target_label is not None:
        err = _validate_string_field("target_label", target_label, max_length=MAX_TARGET_ALIAS_LENGTH)
        if err:
            return jsonify({"error": err}), 400
        label_clean = target_label.strip()
        if label_clean and not ALLOWED_IDENTIFIER_RE.match(label_clean):
            return jsonify({"error": "target_label contiene caracteres no permitidos"}), 400
        body["target_label"] = label_clean or None

    corr_id = _new_correlation_id()
    t0 = time.perf_counter()

    db = request.environ.get("db_session")

    try:
        creds = (
            db.query(TenantCredentials)
            .filter_by(client_key=body["client_key"], enabled=True)
            .first()
        )
        storage_query = (
            db.query(StorageTargets)
            .filter_by(client_key=body["client_key"], tenant_id=creds.id)
        )
        label = body.get("target_label")
        target_alias = body.get("target_alias")
        tenant_user = None
        if target_alias:
            tenant_user = (
                db.query(TenantUsers)
                .filter_by(tenant_id=creds.id, alias=target_alias)
                .first()
            )
            if not tenant_user:
                return jsonify({"error": f"target_alias '{target_alias}' no está configurado"}), 400
            storage = storage_query.filter_by(tenant_user_id=tenant_user.id).first()
            if not storage:
                return jsonify({"error": f"No hay destino configurado para el alias '{target_alias}'"}), 400
        else:
            if label:
                storage = storage_query.filter_by(label=label).first()
            else:
                storage = storage_query.first()
        
        if not creds or not storage:
            return jsonify({"error": "Configuración incompleta en DB"}), 400

        # Token MSAL
        auth = MicrosoftGraphAuthenticator(
            creds.tenant_id, creds.app_client_id, creds.app_client_secret
        )
        token = auth.get_access_token()
        gs = GraphServices(access_token=token, correlation_id=corr_id)

        # Ruta al archivo existente
        full_dest_path = f"{storage.default_dest_folder_path.strip('/')}/{body['dest_file_name']}"

        # Escribir todas las celdas
        result, ms_request_ids = gs.write_cells_graph(
            full_dest_path=full_dest_path,
            data=body["data"],
            target_user_id=storage.target_user_id if not storage.use_drive_id else None,
            drive_id=storage.drive_id if storage.use_drive_id else None,
        )

        written = result.get("written", {})
        success_cells = [cell for cell, info in written.items() if info.get("status") == "ok"]
        error_cells = {cell: info for cell, info in written.items() if info.get("status") == "error"}

        # Log (sin template_id porque es edición directa)
        duration_ms = int((time.perf_counter() - t0) * 1000)
        log_row = RenderLogs(
            client_key=body["client_key"],
            template_id=None,  # importante: None, no 0
            template_key="__manual_write__",
            data_json=body["data"],
            status=RenderStatus.SUCCESS,
            requested_by=body.get("requested_by", "eco-agent"),
            created_at=datetime.utcnow(),
            duration_ms=duration_ms,
            dest_file_name=body["dest_file_name"]
        )
        db.add(log_row)
        # db.commit()  # opcional; tu teardown ya hace commit si no hay excepción

        response_payload = {
            "message": "OK" if not error_cells else "Parcial",
            "written": written,
            "correlation_id": corr_id,
            "duration_ms": duration_ms,
            "ms_request_ids": ms_request_ids,
        }

        status_code = 200 if not error_cells else 207  # 207 Multi-Status para resultados mixtos
        return jsonify(response_payload), status_code

    except GraphAPIError as ge:
        try:
            duration_ms = int((time.perf_counter() - t0) * 1000)
            log_row = RenderLogs(
                client_key=body.get("client_key"),
                template_id=None,
                template_key="__manual_write__",
                data_json=body.get("data", {}),
                status=RenderStatus.ERROR,
                error_message=f"{ge.message} | ms-request-id={ge.ms_request_id}",
                requested_by=body.get("requested_by", "eco-agent"),
                created_at=datetime.utcnow(),
                duration_ms=duration_ms,
                dest_file_name=body.get("dest_file_name"),
            )
            db.add(log_row)
        except Exception:
            pass

        status = ge.status_code if 400 <= (ge.status_code or 0) <= 599 else 502
        payload = {
            "error": ge.message,
            "correlation_id": corr_id,
        }
        if ge.ms_request_id:
            payload["ms_request_id"] = ge.ms_request_id
        return jsonify(payload), status

    except Exception as e:
        # Intentar registrar el error también
        try:
            duration_ms = int((time.perf_counter() - t0) * 1000)
            log_row = RenderLogs(
                client_key=body.get("client_key"),
                template_id=None,
                template_key="__manual_write__",
                data_json=body.get("data", {}),
                status=RenderStatus.ERROR,
                error_message=str(e),
                requested_by=body.get("requested_by", "eco-agent"),
                created_at=datetime.utcnow(),
                duration_ms=duration_ms,
                dest_file_name=body["dest_file_name"]
            )
            db.add(log_row)
            # db.commit()  # opcional, tu teardown gestionará rollback si vuelve a fallar
        except Exception:
            pass

        return jsonify({"error": str(e), "correlation_id": corr_id}), 500
