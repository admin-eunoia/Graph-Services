# Graph Services API

Flask API que automatiza la generación y edición de archivos Excel sobre Microsoft Graph. Incluye:

- Autenticación MSAL per tenant con `MicrosoftGraphAuthenticator` (`Auth/Microsoft_Graph_Auth.py`).
- Cliente de alto nivel `GraphServices` (`Services/graph_services.py`) con reintentos y telemetría.
- Validaciones de payload (`validators/payload.py`) y modelos SQLAlchemy para la persistencia de configuración (`Postgress/Tables.py`).

## Requisitos

- Python 3.10+
- PostgreSQL accesible con credenciales creadas.
- Aplicación registrada en Azure AD con permisos de Microsoft Graph para OneDrive/Excel.

## Variables de entorno

Configura un archivo `.env` en la raíz con los valores necesarios:

```env
FLASK_SECRET_KEY=super-secret
PORT=8000
RATE_LIMITS="120 per minute; 5000 per hour"
TRUSTED_API_KEY=opcional-para-saltar-rate-limit

DB_USER=postgres
DB_PASSWORD=postgres
DB_HOST=localhost
DB_PORT=5432
DB_NAME=graph_services
```

Las credenciales de Microsoft Graph se almacenan por tenant en la tabla `tenant_credentials`, por lo que no se necesitan aquí, pero deben existir en la base de datos antes de consumir el servicio.

## Instalación

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Inicializa la base de datos y arranca la aplicación:

```bash
python main.py
```

Al iniciar, se crea el esquema público en PostgreSQL (si no existe), se generan las tablas declaradas y se aplican límites de tasa configurables mediante `flask-limiter`.

## Endpoints

Todas las rutas están bajo el prefijo `/graph`.

### `POST /graph/excel/render-upload`

1. Valida el payload (`client_key`, `template_key`, `tenant_name`, `data`, opcional `naming` y selectores de destino).
2. Obtiene las credenciales y definición de template/almacenamiento desde PostgreSQL.
3. Descarga el template de OneDrive (drive o usuario) y lo rellena en memoria.
4. Construye el nombre final mediante `dest_file_pattern` y sube el archivo rellenado.
5. Registra el resultado en `render_logs`, incluyendo duración, request IDs de Microsoft y URL generada.

Payload mínimo de ejemplo:

```json
{
  "client_key": "demo",
  "template_key": "reporte-mensual",
  "tenant_name": "Contoso",
  "data": {
    "A1": "Contoso",
    "Resumen!B3": 42
  }
}
```

### `POST /graph/excel/write-cells`

Actualiza celdas en un archivo existente sin pasar por un template:

1. Valida que `dest_file_name` termine en `.xlsx` y el mapa de celdas.
2. Resuelve la ubicación del archivo en Graph (drive o usuario) a partir de la configuración en PostgreSQL.
3. Para cada celda hace `PATCH` vía Graph Excel API, acumulando los request IDs y resultados (`ok`/`error`).
4. Guarda un log con `template_key="__manual_write__"` y los datos utilizados.

Si alguna celda falla, responde con `207 Multi-Status` para que el cliente pueda inspeccionar qué direcciones fallaron.

## Estructura de la base de datos

`Postgress/Tables.py` define las entidades principales:

- `TenantCredentials`: credenciales y metadatos por cliente.
- `TenantUsers`: aliases de usuarios impersonados.
- `StorageTargets`: destino en OneDrive/Drive para cada combinación (cliente, usuario, ubicación).
- `Templates`: describe carpetas, archivo base, patrón de nombres y comportamiento de conflicto.
- `RenderLogs`: historial de ejecuciones, duración, estado y errores.

Cada endpoint abre una sesión SQLAlchemy por request (`main.py`) y delega el commit/rollback al `teardown_request`.

## Validaciones y helpers

- `validators/payload.py` centraliza longitudes máximas, sanitización de nombres, validación de `location_type` (`drive`/`user`) y construcción del nombre final.
- `Services/excel_render.py` usa `openpyxl` para rellenar workbooks en memoria, teniendo en cuenta celdas combinadas.
- `GraphServices` encapsula llamadas HTTP (`requests`), añade reintentos exponenciales para códigos 423/429/50x y registra los `request-id` de Microsoft en las respuestas.

## Desarrollo

- Ejecuta pruebas unitarias con `pytest` (si agregas tests):

```bash
pytest
```

- Mantén las configuraciones sensibles fuera del repositorio y usa Azure Key Vault u otro mecanismo seguro para los secretos de clientes.
