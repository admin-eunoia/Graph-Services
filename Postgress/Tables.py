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
    use_drive_id = Column(Boolean, default=False, nullable=False)
    drive_id = Column(String(200), nullable=True)
    target_user_id = Column(String(200), nullable=True)
    default_dest_folder_path = Column(String(500), nullable=False)

    # Metadata
    label = Column(String(100), nullable=True)
    tenant_user_id = Column(Integer, ForeignKey("tenant_users.id"), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)

    __table_args__ = (
        UniqueConstraint('client_key', 'label', name='uq_storage_targets_clientkey_label'),
        UniqueConstraint('client_key', 'tenant_user_id', name='uq_storage_targets_clientkey_user'),
        CheckConstraint(
            "(drive_id IS NOT NULL) OR (target_user_id IS NOT NULL)",
            name="ck_storage_targets_has_location"
        ),
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
    # Opcional:
    # updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)


class RenderLogs(Base):
    __tablename__ = "render_logs"
    id = Column(Integer, primary_key=True, autoincrement=True)
    client_key = Column(String, ForeignKey("tenant_credentials.client_key"), nullable=False)

    template_id = Column(Integer, ForeignKey("templates.id"), nullable=True)
    template_key = Column(String, nullable=False)

    data_json = Column(JSON, nullable=False)
    result_drive_item_id = Column(String, nullable=True)
    result_web_url = Column(String, nullable=True)

    dest_file_name = Column(String(255), nullable=True)

    status = Column(
        PGEnum(RenderStatus, name="renderstatus", schema="public", create_type=True),
        nullable=False,
        default=RenderStatus.PENDING
    )
    error_message = Column(String, nullable=True)
    requested_by = Column(String, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    duration_ms = Column(Integer, nullable=True)
