from sqlalchemy import Column, String, Boolean, JSON, DateTime, ForeignKey, Integer, UniqueConstraint, Index, CheckConstraint, Text
from sqlalchemy.ext.declarative import declarative_base
from datetime import datetime
import enum
from sqlalchemy.dialects.postgresql import ENUM as PGEnum

Base = declarative_base()

class RenderStatus(str, enum.Enum):
    success = "success"
    error = "error"
    pending = "pending"
    partial = "partial"


class OperationType(str, enum.Enum):
    copy_template = "copy_template"
    write_section = "write_section"
    write_table = "write_table"
    insert_rows = "insert_rows"
    update_cell = "update_cell"
    search_marker = "search_marker"
    apply_merge = "apply_merge"


class LocationType(enum.Enum):
    drive = "drive"
    user = "user"


class DataType(str, enum.Enum):
    text = "text"
    number = "number"
    date = "date"
    boolean = "boolean"
    formula = "formula"

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

    location_type = Column(
        PGEnum(LocationType, name="locationtype", schema="public", create_type=True),
        nullable=False
    )
    location_identifier = Column(String(200), nullable=False)
    default_dest_folder_path = Column(String(500), nullable=False)

    tenant_user_id = Column(Integer, ForeignKey("tenant_users.id"), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)

    __table_args__ = (
        UniqueConstraint('client_key', 'location_type', 'location_identifier', name='uq_storage_targets_location'),
        UniqueConstraint('client_key', 'tenant_user_id', name='uq_storage_targets_clientkey_user'),
    CheckConstraint("char_length(location_identifier) > 0", name="ck_storage_targets_identifier_not_empty"),
    Index('ix_storage_targets_client_key', 'client_key'),
    Index('ix_storage_targets_tenant_id', 'tenant_id'),
    )

class Templates(Base):
    __tablename__ = "templates"
    id = Column(Integer, primary_key=True, autoincrement=True)
    client_key = Column(String(100), ForeignKey("tenant_credentials.client_key"), nullable=False)
    template_key = Column(String(100), nullable=False)
    description = Column(String(500), nullable=True)
    template_version = Column(String(50), nullable=True, default="1.0")
    template_folder_path = Column(String(500), nullable=False, comment='Ruta donde está el template vacío')
    template_file_name = Column(String(255), nullable=False, comment='Nombre del archivo template')
    dest_file_pattern = Column(String(255), nullable=False, comment='Patrón con variables: {cliente}_{fecha}.xlsx')
    default_sheet_name = Column(String(100), nullable=True, comment='Hoja por defecto (null = primera hoja)')
    is_active = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    __table_args__ = (
        UniqueConstraint('client_key', 'template_key', name='uq_templates_clientkey_templatekey'),
        Index('ix_templates_client_key', 'client_key'),
        Index('ix_templates_template_key', 'template_key'),
        Index('ix_templates_active', 'is_active'),
    )

class ExcelFiles(Base):
    __tablename__ = "excel_files"
    id = Column(Integer, primary_key=True, autoincrement=True)
    client_key = Column(String(100), ForeignKey("tenant_credentials.client_key"), nullable=False)
    template_id = Column(Integer, ForeignKey("templates.id"), nullable=False)
    storage_target_id = Column(Integer, ForeignKey("storage_targets.id"), nullable=False)
    
    file_key = Column(String(100), nullable=False, unique=True)
    file_folder_path = Column(String(500), nullable=False)
    file_name = Column(String(255), nullable=False)
    
    item_id = Column(String(200), nullable=True)
    web_url = Column(String(1024), nullable=True)
    context_data = Column(JSON, nullable=True)
    
    is_active = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    __table_args__ = (
        UniqueConstraint('storage_target_id', 'file_folder_path', 'file_name', name='uq_excel_files_location'),
    Index('ix_excel_files_client_key', 'client_key'),
    Index('ix_excel_files_template_id', 'template_id'),
    Index('ix_excel_files_file_key', 'file_key'),
        Index('ix_excel_files_item_id', 'item_id'),
        Index('ix_excel_files_active', 'is_active'),
    )

class ExcelSections(Base):
    __tablename__ = "excel_sections"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    template_id = Column(Integer, ForeignKey("templates.id", ondelete="CASCADE"), nullable=False)
    client_key = Column(String(100), ForeignKey("tenant_credentials.client_key"), nullable=False)
    section_key = Column(String(100), nullable=False)
    section_name = Column(String(200), nullable=True)
    description = Column(Text, nullable=True)
    marker_text = Column(String(255), nullable=False)
    sheet_name = Column(String(100), nullable=True)
    
    is_table = Column(Boolean, default=False, nullable=False)
    row_offset = Column(Integer, default=1, nullable=False)
    column_offset = Column(Integer, default=0, nullable=False)
    merge_ranges = Column(JSON, nullable=True)
    order_index = Column(Integer, default=0, nullable=False)
    is_active = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    __table_args__ = (
        UniqueConstraint('template_id', 'section_key', name='uq_sections_template_section'),
    Index('ix_sections_template_id', 'template_id'),
    Index('ix_sections_client_key', 'client_key'),
    Index('ix_sections_section_key', 'section_key'),
        Index('ix_sections_marker', 'marker_text'),
        Index('ix_sections_active', 'is_active'),
        Index('ix_sections_order', 'order_index'),
    )

class ExcelFields(Base):
    __tablename__ = "excel_fields"
    id = Column(Integer, primary_key=True, autoincrement=True)
    section_id = Column(Integer, ForeignKey("excel_sections.id", ondelete="CASCADE"), nullable=False)
    template_id = Column(Integer, ForeignKey("templates.id", ondelete="CASCADE"), nullable=False)
    client_key = Column(String(100), ForeignKey("tenant_credentials.client_key"), nullable=False)
    
    field_key = Column(String(100), nullable=False)
    field_name = Column(String(200), nullable=True)
    column_offset = Column(Integer, nullable=False)
    
    data_type = Column(
        PGEnum(DataType, name="datatype", schema="public", create_type=True),
        nullable=False,
        default=DataType.text
    )
    is_required = Column(Boolean, default=False, nullable=False)
    default_value = Column(String(500), nullable=True)
    format_pattern = Column(String(100), nullable=True)
    validation_rules = Column(JSON, nullable=True)
    description = Column(Text, nullable=True)
    example_value = Column(String(200), nullable=True)
    is_active = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)

    __table_args__ = (
        UniqueConstraint('section_id', 'field_key', name='uq_fields_section_field'),
    Index('ix_fields_section_id', 'section_id'),
    Index('ix_fields_template_id', 'template_id'),
    Index('ix_fields_client_key', 'client_key'),
        Index('ix_fields_field_key', 'field_key'),
        Index('ix_fields_offset', 'column_offset'),
    )

class GraphTokens(Base):
    __tablename__ = "graph_tokens"
    id = Column(Integer, primary_key=True, autoincrement=True)
    client_key = Column(String(100), ForeignKey("tenant_credentials.client_key"), nullable=False)
    tenant_id = Column(Integer, ForeignKey("tenant_credentials.id"), nullable=False)
    user_email = Column(String(255), nullable=True)
    
    access_token = Column(Text, nullable=False)
    refresh_token = Column(Text, nullable=True)
    token_type = Column(String(50), default="Bearer", nullable=False)
    
    expires_at = Column(DateTime, nullable=False)
    scope = Column(String(1000), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)
    last_used_at = Column(DateTime, nullable=True)
    use_count = Column(Integer, default=0, nullable=False)

    __table_args__ = (
        UniqueConstraint('client_key', 'user_email', name='uq_tokens_client_user'),
    Index('ix_tokens_client_key', 'client_key'),
    Index('ix_tokens_expires_at', 'expires_at'),
    Index('ix_tokens_tenant_id', 'tenant_id'),
        Index('ix_tokens_last_used', 'last_used_at'),
        CheckConstraint('expires_at > created_at', name='ck_tokens_expires_after_created'),
    )

class OperationLogs(Base):
    __tablename__ = "operation_logs"
    id = Column(Integer, primary_key=True, autoincrement=True)
    operation_id = Column(String(100), unique=True, nullable=False)
    correlation_id = Column(String(100), nullable=True)
    client_key = Column(String(100), ForeignKey("tenant_credentials.client_key"), nullable=False)
    template_id = Column(Integer, ForeignKey("templates.id"), nullable=True)
    excel_file_id = Column(Integer, ForeignKey("excel_files.id"), nullable=True)
    operation_type = Column(
        PGEnum(OperationType, name="operationtype", schema="public", create_type=True),
        nullable=False
    )
    
    section_id = Column(Integer, ForeignKey("excel_sections.id"), nullable=True)
    sheet_name = Column(String(100), nullable=True)
    marker_text = Column(String(255), nullable=True)
    marker_found = Column(Boolean, nullable=True)
    marker_position = Column(String(50), nullable=True)
    rows_affected = Column(Integer, nullable=True)
    cells_affected = Column(Integer, nullable=True)
    
    input_data = Column(JSON, nullable=True)
    output_data = Column(JSON, nullable=True)
    
    status = Column(
        PGEnum(RenderStatus, name="renderstatus", schema="public", create_type=True),
        nullable=False,
        default=RenderStatus.pending
    )
    error_message = Column(Text, nullable=True)
    error_code = Column(String(50), nullable=True)
    error_stack_trace = Column(Text, nullable=True)
    
    ms_request_ids = Column(JSON, nullable=True)
    retry_count = Column(Integer, default=0, nullable=False)
    duration_ms = Column(Integer, nullable=True)
    
    requested_by = Column(String(200), nullable=True)
    executed_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    
    __table_args__ = (
        Index('ix_operation_logs_operation_id', 'operation_id'),
        Index('ix_operation_logs_correlation_id', 'correlation_id'),
        Index('ix_operation_logs_client_key', 'client_key'),
        Index('ix_operation_logs_template_id', 'template_id'),
        Index('ix_operation_logs_excel_file_id', 'excel_file_id'),
        Index('ix_operation_logs_operation_type', 'operation_type'),
        Index('ix_operation_logs_status', 'status'),
        Index('ix_operation_logs_executed_at', 'executed_at'),
        Index('ix_operation_logs_error_code', 'error_code'),
    )
