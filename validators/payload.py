import re
import uuid
from datetime import datetime
from string import Formatter

MAX_CLIENT_FIELD_LENGTH = 100
MAX_TENANT_NAME_LENGTH = 100
MAX_TEMPLATE_KEY_LENGTH = 100
MAX_DEST_FILE_NAME_LENGTH = 255
MAX_DATA_ENTRIES = 500
MAX_DATA_VALUE_LENGTH = 2048
MAX_DATA_TOTAL_LENGTH = 20000
MAX_NAMING_ENTRIES = 30
MAX_NAMING_KEY_LENGTH = 64
MAX_NAMING_VALUE_LENGTH = 512
MAX_TARGET_ALIAS_LENGTH = 100
MAX_LOCATION_IDENTIFIER_LENGTH = 200
ALLOWED_IDENTIFIER_RE = re.compile(r"^[A-Za-z0-9._-]+$")
CELL_REF_RE = re.compile(r"^(?:([^!]+)!)?([A-Za-z]+[1-9][0-9]*)$")
DRIVE_ID_PATTERN = re.compile(r"^[A-Za-z0-9!._-]{16,}$")
USER_GUID_PATTERN = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", re.IGNORECASE)
USER_UPN_PATTERN = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")


def require_fields(body, keys):
    for key in keys:
        if key not in body:
            return f"Falta campo requerido: {key}"
    return None


def new_correlation_id():
    return str(uuid.uuid4())


def validate_string_field(field_name: str, value, *, max_length: int):
    if not isinstance(value, str) or not value.strip():
        return f"{field_name} debe ser un string no vacío"
    if len(value.strip()) > max_length:
        return f"{field_name} excede longitud máxima ({max_length})"
    return None


def validate_data_dict(data):
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


def pattern_fields(pattern: str) -> set[str]:
    formatter = Formatter()
    return {field for _, field, _, _ in formatter.parse(pattern) if field}


def build_dest_file_name(template, body, *, naming_provided: bool = False) -> str:
    pattern = (template.dest_file_pattern or "{template_key}_{tenant_name_sanitized}.xlsx").strip()
    if not pattern:
        pattern = "{template_key}_{tenant_name_sanitized}.xlsx"

    if "tenant_name" not in body:
        raise ValueError("Debe proporcionarse tenant_name")
    raw_tenant_name = str(body["tenant_name"]).strip()
    tenant_name_sanitized = re.sub(r"[^A-Za-z0-9._-]+", "_", raw_tenant_name).strip("_")
    if not tenant_name_sanitized:
        tenant_name_sanitized = "tenant"

    naming = body.get("naming", {}) or {}
    if not isinstance(naming, dict):
        raise ValueError("naming debe ser un objeto (dict)")

    context = {
        "client_key": body["client_key"],
        "template_key": template.template_key,
        "tenant_name": raw_tenant_name,
        "tenant_name_sanitized": tenant_name_sanitized,
    }
    context.update(naming)

    required = pattern_fields(pattern)
    missing = [k for k in required if k not in context]
    if missing:
        if naming_provided:
            raise ValueError(f"Faltan campos para naming: {', '.join(missing)}")
        context.setdefault("timestamp", datetime.utcnow().strftime("%Y%m%d-%H%M%S"))
        pattern = "{template_key}_{tenant_name_sanitized}_{timestamp}.xlsx"
        required = pattern_fields(pattern)
        missing = [k for k in required if k not in context]
        if missing:
            raise ValueError("No se pudo construir nombre de archivo")
    else:
        context.setdefault("timestamp", datetime.utcnow().strftime("%Y%m%d-%H%M%S"))

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


def validate_naming_dict(naming):
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


def validate_cell_map_or_raise(data):
    if not isinstance(data, dict) or not data:
        raise ValueError("data debe ser un diccionario no vacío de {celda: valor}.")

    for cell, value in data.items():
        if not isinstance(cell, str) or not CELL_REF_RE.match(cell):
            raise ValueError(f"Dirección de celda inválida: '{cell}'. Usa 'A1' o 'Hoja!B2'.")
        if not (isinstance(value, (str, int, float, bool)) or value is None):
            raise ValueError(f"Valor no soportado para '{cell}': {type(value).__name__}")


def validate_location_selector(location_type, location_identifier):
    if location_type is None and location_identifier is None:
        return None, (None, None)
    if location_type is None or location_identifier is None:
        return "Debes proporcionar ambos location_type y location_identifier", (None, None)
    if not isinstance(location_type, str):
        return "location_type debe ser 'drive' o 'user'", (None, None)

    normalized_type = location_type.strip().lower()
    if normalized_type not in {"drive", "user"}:
        return "location_type debe ser 'drive' o 'user'", (None, None)

    if not isinstance(location_identifier, str) or not location_identifier.strip():
        return "location_identifier debe ser string no vacío", (None, None)

    ident_clean = location_identifier.strip()
    if len(ident_clean) > MAX_LOCATION_IDENTIFIER_LENGTH:
        return f"location_identifier excede {MAX_LOCATION_IDENTIFIER_LENGTH} caracteres", (None, None)

    if normalized_type == "drive":
        if not DRIVE_ID_PATTERN.match(ident_clean):
            return "location_identifier para drive no tiene formato válido", (None, None)
    else:
        if not (USER_GUID_PATTERN.match(ident_clean) or USER_UPN_PATTERN.match(ident_clean)):
            return "location_identifier para user debe ser GUID o UPN", (None, None)

    return None, (normalized_type, ident_clean)
