# routes/routes.py
import time
from datetime import datetime
from typing import Optional, Tuple
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
from Services.graph_services import GraphServices, GraphAPIError
from validators.payload import (
    ALLOWED_IDENTIFIER_RE,
    MAX_CLIENT_FIELD_LENGTH,
    MAX_TENANT_NAME_LENGTH,
    MAX_DEST_FILE_NAME_LENGTH,
    MAX_TARGET_ALIAS_LENGTH,
    MAX_TEMPLATE_KEY_LENGTH,
    build_dest_file_name,
    new_correlation_id,
    require_fields,
    validate_cell_map_or_raise,
    validate_data_dict,
    validate_location_selector,
    validate_naming_dict,
    validate_string_field,
)

graph_bp = Blueprint("graph", __name__)

# ------------------ Helpers ------------------ #
def _join_storage_path(*segments: str) -> str:
    cleaned = []
    for segment in segments:
        if segment is None:
            continue
        stripped = str(segment).strip()
        if not stripped:
            continue
        cleaned.append(stripped.strip("/"))
    if not cleaned:
        return ""
    return "/".join(cleaned)


def _resolve_graph_target(storage) -> Tuple[Optional[str], Optional[str]]:
    """Devuelve (drive_id, target_user_id) soportando los campos nuevos y legacy."""
    if storage is None:
        return None, None

    location_type = getattr(storage, "location_type", None)
    location_identifier = getattr(storage, "location_identifier", None)
    if location_type in {"drive", "user"} and location_identifier:
        if location_type == "drive":
            return location_identifier, None
        return None, location_identifier

    return None, getattr(storage, "target_user_id", None)


# ------------------ ROUTES ------------------ #

@graph_bp.post("/excel/render-upload")
def render_upload():
    """
    1) Lee tenant/template/target/tenantuser de DB
    2) Descarga template desde OneDrive del cliente
    3) Rellena en memoria con 'data'
    4) Sube el archivo final al destino (conflictBehavior del template)
    5) Inserta log en render_logs

    curl -X POST http://localhost:8000/graph/excel/render-upload 
        -H "Content-Type: application/json" 
        -H "X-Api-Key: opcional-si-quieres-saltar-rate-limit" 
        -d '{
            "client_key": "contoso",
            "template_key": "reporte-mensual",
            "tenant_name": "Contoso S.A.",
            "data": {
            "A1": "Contoso",
            "Hoja1!B3": 42,
            "Hoja1!C5": "Ingreso"
            },
            "naming": {                                       opcional
            "periodo": "2024-05",
            "sucursal": "mexico"
            },
            "target_alias": "finance",                        opcional
            "location_type": "drive",                         opcional
            "location_identifier": "b!oZ1234567890abcdef",    opcional
            "requested_by": "api-client-123"                  opcional
    }'
    """
    body = request.get_json() or {}
    err = require_fields(body, ["client_key", "template_key", "data", "tenant_name"])
    if err:
        return jsonify({"error": err}), 400

    err = validate_string_field("client_key", body.get("client_key"), max_length=MAX_CLIENT_FIELD_LENGTH)
    if err:
        return jsonify({"error": err}), 400

    err = validate_string_field("template_key", body.get("template_key"), max_length=MAX_TEMPLATE_KEY_LENGTH)
    if err:
        return jsonify({"error": err}), 400

    err = validate_string_field("tenant_name", body.get("tenant_name"), max_length=MAX_TENANT_NAME_LENGTH)
    if err:
        return jsonify({"error": err}), 400

    err = validate_data_dict(body.get("data"))
    if err:
        return jsonify({"error": err}), 400

    try:
        validate_cell_map_or_raise(body["data"])
    except ValueError as ve:
        return jsonify({"error": str(ve)}), 400

    naming_provided = body.get("naming") not in (None, {})
    naming_err, sanitized_naming = validate_naming_dict(body.get("naming"))
    if naming_err:
        return jsonify({"error": naming_err}), 400
    body["naming"] = sanitized_naming

    target_alias = body.get("target_alias")
    if target_alias is not None:
        err = validate_string_field("target_alias", target_alias, max_length=MAX_TARGET_ALIAS_LENGTH)
        if err:
            return jsonify({"error": err}), 400
        alias_clean = target_alias.strip()
        if alias_clean and not ALLOWED_IDENTIFIER_RE.match(alias_clean):
            return jsonify({"error": "target_alias contiene caracteres no permitidos"}), 400
        body["target_alias"] = alias_clean or None

    loc_err, (normalized_type, ident_clean) = validate_location_selector(body.get("location_type"), body.get("location_identifier"))
    if loc_err:
        return jsonify({"error": loc_err}), 400
    body["location_type"] = normalized_type
    body["location_identifier"] = ident_clean

    corr_id = new_correlation_id()
    t0 = time.perf_counter()

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
        if not creds:
            return jsonify({"error": "Configuración incompleta en DB"}), 400

        storage_query = (
            db.query(StorageTargets)
            .filter_by(client_key=body["client_key"], tenant_id=creds.id)
        )
        target_alias = body.get("target_alias")
        location_type = body.get("location_type")
        location_identifier = body.get("location_identifier")
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
            if location_type and location_identifier:
                storage = (
                    storage_query
                    .filter_by(location_type=location_type, location_identifier=location_identifier)
                    .first()
                )
                if not storage:
                    return jsonify({"error": "No hay destino configurado para la ubicación indicada"}), 400
            else:
                storage = storage_query.first()

        template = (
            db.query(Templates)
            .filter_by(client_key=body["client_key"], template_key=body["template_key"])
            .first()
        )
        if not storage or not template:
            return jsonify({"error": "Configuración incompleta en DB"}), 400

        # 2) Token MSAL
        auth = MicrosoftGraphAuthenticator(
            creds.tenant_id, creds.app_client_id, creds.app_client_secret
        )
        token = auth.get_access_token()
        gs = GraphServices(access_token=token, correlation_id=corr_id)

        # 3) Descargar template
        full_template_path = _join_storage_path(
            template.template_folder_path,
            template.template_file_name,
        )
        drive_id, target_user_graph_id = _resolve_graph_target(storage)

        tpl_bytes, ms_id_download = gs.download_file_bytes(
            full_template_path,
            target_user_id=target_user_graph_id,
            drive_id=drive_id,
        )

        # 4) Render en memoria
        filled_bytes = gs.render_in_memory(tpl_bytes, body["data"])

        # 5) Calcular nombre de archivo final según patrón configurable
        try:
            dest_file_name = build_dest_file_name(template, body, naming_provided=naming_provided)
        except ValueError as ve:
            return jsonify({"error": str(ve)}), 400

        dest_path = _join_storage_path(storage.default_dest_folder_path, dest_file_name)
        upload_result, ms_id_upload = gs.upload_file_bytes(
            filled_bytes,
            dest_path,
            conflict_behavior=template.default_conflict_behavior,
            target_user_id=target_user_graph_id,
            drive_id=drive_id,
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

    curl -X POST http://localhost:8000/graph/excel/write-cells \
        -H "Content-Type: application/json" \
        -H "X-Api-Key: opcional-si-quieres-saltar-rate-limit" \
        -d '{
            "client_key": "contoso",
            "dest_file_name": "reporte-abril.xlsx",
            "data": {
            "A1": "Contoso S.A.",
            "Resumen!B3": 12850,
            "Resumen!C4": "OK"
            },
            "target_alias": "finance",                              opcional
            "location_type": "user",                                opcional
            "location_identifier": "finance-team@contoso.com",      opcional
            "requested_by": "api-client-123"                        opcional
    }'
    """
    body = request.get_json() or {}
    err = require_fields(body, ["client_key", "tenant_name", "dest_file_name", "data"])
    if err:
        return jsonify({"error": err}), 400

    err = validate_string_field("client_key", body.get("client_key"), max_length=MAX_CLIENT_FIELD_LENGTH)
    if err:
        return jsonify({"error": err}), 400

    err = validate_string_field("dest_file_name", body.get("dest_file_name"), max_length=MAX_DEST_FILE_NAME_LENGTH)
    if err:
        return jsonify({"error": err}), 400

    err = validate_data_dict(body.get("data"))
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
        err = validate_string_field("target_alias", target_alias, max_length=MAX_TARGET_ALIAS_LENGTH)
        if err:
            return jsonify({"error": err}), 400
        alias_clean = target_alias.strip()
        if alias_clean and not ALLOWED_IDENTIFIER_RE.match(alias_clean):
            return jsonify({"error": "target_alias contiene caracteres no permitidos"}), 400
        body["target_alias"] = alias_clean or None

    loc_err, (normalized_type, ident_clean) = validate_location_selector(
        body.get("location_type"), body.get("location_identifier")
    )
    if loc_err:
        return jsonify({"error": loc_err}), 400
    body["location_type"] = normalized_type
    body["location_identifier"] = ident_clean

    corr_id = new_correlation_id()
    t0 = time.perf_counter()

    db = request.environ.get("db_session")

    try:
        creds = (
            db.query(TenantCredentials)
            .filter_by(client_key=body["client_key"], enabled=True)
            .first()
        )
        if not creds:
            return jsonify({"error": "Configuración incompleta en DB"}), 400

        storage_query = (
            db.query(StorageTargets)
            .filter_by(client_key=body["client_key"], tenant_id=creds.id)
        )
        target_alias = body.get("target_alias")
        location_type = body.get("location_type")
        location_identifier = body.get("location_identifier")
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
            if location_type and location_identifier:
                storage = (
                    storage_query
                    .filter_by(location_type=location_type, location_identifier=location_identifier)
                    .first()
                )
                if not storage:
                    return jsonify({"error": "No hay destino configurado para la ubicación indicada"}), 400
            else:
                storage = storage_query.first()

        if not storage:
            return jsonify({"error": "Configuración incompleta en DB"}), 400

        # Token MSAL
        auth = MicrosoftGraphAuthenticator(
            creds.tenant_id, creds.app_client_id, creds.app_client_secret
        )
        token = auth.get_access_token()
        gs = GraphServices(access_token=token, correlation_id=corr_id)

        # Ruta al archivo existente
        full_dest_path = _join_storage_path(
            storage.default_dest_folder_path,
            body["dest_file_name"],
        )

        # Escribir todas las celdas
        drive_id, target_user_graph_id = _resolve_graph_target(storage)

        result, ms_request_ids = gs.write_cells_graph(
            full_dest_path=full_dest_path,
            data=body["data"],
            target_user_id=target_user_graph_id,
            drive_id=drive_id,
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
    err = validate_string_field("tenant_name", body.get("tenant_name"), max_length=MAX_TENANT_NAME_LENGTH)
    if err:
        return jsonify({"error": err}), 400
