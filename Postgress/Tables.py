from sqlalchemy import Column, String, Boolean, JSON, DateTime, ForeignKey, Integer, Enum, UniqueConstraint, Index, CheckConstraint
from sqlalchemy.ext.declarative import declarative_base
from datetime import datetime
import enum
from sqlalchemy.dialects.postgresql import ENUM as PGEnum

Base = declarative_base()


class RenderStatus(str, enum.Enum):
    SUCCESS = "success"
    ERROR = "error"
    PENDING = "pending"


class TenantCredentials(Base):
    __tablename__ = "tenant_credentials"
    id = Column(Integer, primary_key=True, autoincrement=True)
    client_key = Column(String(100), unique=True, nullable=False)  # length sugerido

    # Auth
    tenant_id = Column(String(64), nullable=False)
    app_client_id = Column(String(64), nullable=False)
    app_client_secret = Column(String(256), nullable=False)  # ojo: manejar cifrado/vault

    tenant_name = Column(String(200), nullable=False)

    enabled = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)


class TenantUsers(Base):
    __tablename__ = "tenant_users"
    id = Column(Integer, primary_key=True, autoincrement=True)
    tenant_id = Column(Integer, ForeignKey("tenant_credentials.id"), nullable=False)
    alias = Column(String(100), nullable=False)
    email = Column(String(200), nullable=True)
    first_name = Column(String(100), nullable=True)
    last_name = Column(String(100), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)

    __table_args__ = (
        UniqueConstraint('tenant_id', 'alias', name='uq_tenant_users_alias'),
        Index('ix_tenant_users_tenant_alias', 'tenant_id', 'alias'),
    )


class StorageTargets(Base):
    __tablename__ = "storage_targets"
    id = Column(Integer, primary_key=True, autoincrement=True)
    client_key = Column(String(100), ForeignKey("tenant_credentials.client_key"), nullable=False)
    tenant_id = Column(Integer, ForeignKey("tenant_credentials.id"), nullable=False)

    # Definición del destino
    location_type = Column(Enum('drive', 'user', name='storage_target_location_type'), nullable=False)
    location_identifier = Column(String(200), nullable=False)
    default_dest_folder_path = Column(String(500), nullable=False)

    # Metadata
    tenant_user_id = Column(Integer, ForeignKey("tenant_users.id"), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)

    __table_args__ = (
        UniqueConstraint('client_key', 'location_type', 'location_identifier', name='uq_storage_targets_location'),
        UniqueConstraint('client_key', 'tenant_user_id', name='uq_storage_targets_clientkey_user'),
        CheckConstraint("char_length(location_identifier) > 0", name="ck_storage_targets_identifier_not_empty"),
    )


class Templates(Base):
    __tablename__ = "templates"
    id = Column(Integer, primary_key=True, autoincrement=True)
    client_key = Column(String(100), ForeignKey("tenant_credentials.client_key"), nullable=False)

    template_key = Column(String(100), nullable=False)
    __table_args__ = (
        UniqueConstraint('client_key', 'template_key', name='uq_templates_clientkey_templatekey'),
    )
    description = Column(String(500), nullable=True)
    template_folder_path = Column(String(500), nullable=False)
    template_file_name = Column(String(255), nullable=False)
    template_version = Column(String(50), nullable=True, default="1.0")
    dest_file_pattern = Column(String(255), nullable=False)
    default_conflict_behavior = Column(String(10), default="fail")
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    
    # cell_mapping: Configuración de secciones dinámicas para escritura basada en marcadores
    # Schema esperado:
    # {
    #   "sections": {
    #     "header": {
    #       "marker": "DATOS DEL CLIENTE:",          // Texto a buscar en el Excel
    #       "sheet": "Hoja1",                        // Nombre de la hoja (opcional, default: hoja activa)
    #       "start_row_offset": 1,                   // Filas después del marcador donde escribir
    #       "start_col_offset": 0,                   // Columnas después del marcador
    #       "is_table": false,                       // false = key-value, true = tabla de filas
    #       "columns": {                             // Mapeo de campos a offset de columnas
    #         "nombre": 0,                           // Columna relativa al marcador
    #         "rfc": 1,
    #         "direccion": 2
    #       }
    #     },
    #     "items": {
    #       "marker": "DETALLE DE PAGOS:",
    #       "sheet": "Hoja1",
    #       "start_row_offset": 2,                   // Filas después del marcador (para headers)
    #       "start_col_offset": 0,
    #       "is_table": true,                        // Permite múltiples filas de datos
    #       "columns": {
    #         "fecha": 0,
    #         "concepto": 1,
    #         "monto": 2,
    #         "estado": 3
    #       },
    #       "merge_ranges": [                        // Opcional: rangos a re-aplicar por fila
    #         "A{row}:B{row}",                       // {row} se reemplaza con número de fila
    #         "C{row}:D{row}"
    #       ]
    #     },
    #     "footer": {
    #       "marker": "TOTALES:",
    #       "start_row_offset": 1,
    #       "is_table": false,
    #       "columns": {
    #         "subtotal": 1,
    #         "iva": 2,
    #         "total": 3
    #       }
    #     }
    #   }
    # }
    #
    # Notas:
    # - Si cell_mapping es null o no contiene "sections", el sistema usa modo legacy (referencias estáticas)
    # - Los valores pueden ser strings, números, booleanos, null, o fórmulas (strings con "=")
    # - Las fórmulas se detectan automáticamente si empiezan con "="
    # - start_row_offset: 0 = misma fila del marcador, 1 = siguiente, etc.
    # - start_col_offset: 0 = misma columna del marcador, 1 = siguiente, etc.
    # - is_table=true permite insertar filas dinámicamente si hay más datos que filas en template
    # - merge_ranges: patron con {row} para re-aplicar merges en cada fila insertada
    cell_mapping = Column(JSON, nullable=True)
    

class RenderLogs(Base):
    __tablename__ = "render_logs"
    id = Column(Integer, primary_key=True, autoincrement=True)
    client_key = Column(String(100), ForeignKey("tenant_credentials.client_key"), nullable=False)

    template_id = Column(Integer, ForeignKey("templates.id"), nullable=True)
    template_key = Column(String(100), nullable=False)

    data_json = Column(JSON, nullable=False)
    result_drive_item_id = Column(String(200), nullable=True)
    result_web_url = Column(String(1024), nullable=True)

    dest_file_name = Column(String(255), nullable=True)

    status = Column(
        PGEnum(RenderStatus, name="renderstatus", schema="public", create_type=True),
        nullable=False,
        default=RenderStatus.PENDING
    )
    error_message = Column(String(1000), nullable=True)
    requested_by = Column(String(200), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    duration_ms = Column(Integer, nullable=True)
