from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker
from dotenv import load_dotenv
import os

load_dotenv()

DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")
DB_HOST = os.getenv("DB_HOST")
DB_PORT = os.getenv("DB_PORT")
DB_NAME = os.getenv("DB_NAME")

DATABASE_URL = f"postgresql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}"

# echo=True para ver SQL en consola; ponlo en False en prod
engine = create_engine(DATABASE_URL, echo=True)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

def init_db():
    # Asegura schema y search_path ⇒ evita el error “no schema has been selected to create in”
    with engine.connect() as conn:
        conn.execute(text("CREATE SCHEMA IF NOT EXISTS public"))
        conn.execute(text("SET search_path TO public"))
        conn.commit()

    # Importa modelos DESPUÉS de crear engine (evita referencias circulares)
    from Postgress.Tables import (
        Base,
        TenantCredentials,
        TenantUsers,
        StorageTargets,
        Templates,
        RenderLogs,
    )
    Base.metadata.create_all(bind=engine)

