"""
Prueba de excel_live_writer.py - Edición EN VIVO
"""
from Services.excel_live_writer import llenar_seccion_live, llenar_tabla_live, procesar_excel_live, insertar_filas_live
from Services.graph_services import GraphServices
from Auth.Microsoft_Graph_Auth import MicrosoftGraphAuthenticator
import os
from dotenv import load_dotenv

# Cargar variables de entorno
load_dotenv()

# Configuración - EDITA AQUÍ
TENANT_ID = os.getenv("MICROSOFT_TENANT_ID", "TU_TENANT_ID_AQUI")
CLIENT_ID = os.getenv("MICROSOFT_CLIENT_ID", "TU_CLIENT_ID_AQUI")
CLIENT_SECRET = os.getenv("MICROSOFT_CLIENT_SECRET", "TU_CLIENT_SECRET_AQUI")
TARGET_USER_ID = "RodrigoAguilera@Eunoia8.onmicrosoft.com"

# Verificar configuración
if TENANT_ID == "TU_TENANT_ID_AQUI" or CLIENT_ID == "TU_CLIENT_ID_AQUI":
    print("=" * 60)
    print("⚠️  CONFIGURACIÓN REQUERIDA")
    print("=" * 60)
    print("\nEdita test_live_writer.py y configura:")
    print("  - TENANT_ID")
    print("  - CLIENT_ID")
    print("  - CLIENT_SECRET")
    print("\nO usa variables de entorno:")
    print("  export MICROSOFT_TENANT_ID='tu-tenant-id'")
    print("  export MICROSOFT_CLIENT_ID='tu-client-id'")
    print("  export MICROSOFT_CLIENT_SECRET='tu-secret'")
    exit(1)

# Archivo a editar (debe existir en OneDrive)
ARCHIVO = "Prueba WAMAN/Prueba #1 WAMAN.xlsx"

# Datos de prueba
datos_cliente = {
    "nombre": "ACME Corporation LIVE",
    "rol": "Cliente Premium"
}

datos_seguimiento = [
    {"fecha": "2025-11-22", "medio": "API LIVE", "comentarios": "Editado en tiempo real"},
    {"fecha": "2025-11-23", "medio": "Sin descargar", "comentarios": "Funciona abierto!"}
]

configuracion = {
    "cliente": {
        "marker": "DATOS DEL CLIENTE",
        "es_tabla": False,
        "columnas": {"nombre": 0, "rol": 1}
    },
    "seguimiento": {
        "marker": "SEGUIMIENTO",
        "es_tabla": True,
        "columnas": {"fecha": 0, "medio": 1, "comentarios": 2}
    }
}


def test_una_celda():
    """Prueba: Editar una sola celda."""
    print("=" * 60)
    print("🔥 TEST 1: EDITAR UNA CELDA EN VIVO")
    print("=" * 60)
    
    auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = auth.get_access_token()
    client = GraphServices(token)
    
    print(f"\n📝 Editando archivo: {ARCHIVO}")
    print("   Solo el campo 'nombre' en DATOS DEL CLIENTE\n")
    
    try:
        llenar_seccion_live(
            client,
            file_path=ARCHIVO,
            marker="DATOS DEL EVENTO",
            datos={"fecha": "69-69-69"},
            columnas={"fecha": 0},
            target_user_id=TARGET_USER_ID
        )
        
        print(f"\n✅ ¡Éxito! Celda editada EN VIVO")
        print(f"   Abre el archivo en OneDrive y verás el cambio")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()


def test_seccion_completa():
    """Prueba: Editar sección completa."""
    print("\n" + "=" * 60)
    print("🔥 TEST 2: EDITAR SECCIÓN COMPLETA")
    print("=" * 60)
    
    auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = auth.get_access_token()
    client = GraphServices(token)
    
    print(f"\n📝 Editando archivo: {ARCHIVO}")
    print("   Campos 'nombre' y 'rol' en DATOS DEL CLIENTE\n")
    
    try:
        llenar_seccion_live(
            client,
            file_path=ARCHIVO,
            marker="DATOS DEL CLIENTE",
            datos=datos_cliente,
            columnas={"nombre": 0, "rol": 1},
            target_user_id=TARGET_USER_ID
        )
        
        print(f"\n✅ ¡Éxito! Sección editada EN VIVO")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()


def test_tabla():
    """Prueba: Llenar tabla."""
    print("\n" + "=" * 60)
    print("🔥 TEST 3: LLENAR TABLA EN VIVO")
    print("=" * 60)
    
    auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = auth.get_access_token()
    client = GraphServices(token)
    
    print(f"\n📝 Editando archivo: {ARCHIVO}")
    print("   Tabla SEGUIMIENTO con 2 filas\n")
    
    try:
        llenar_tabla_live(
            client,
            file_path=ARCHIVO,
            marker="SEGUIMIENTO",
            datos=datos_seguimiento,
            columnas={"fecha": 0, "medio": 1, "comentarios": 2},
            target_user_id=TARGET_USER_ID
        )
        
        print(f"\n✅ ¡Éxito! Tabla llenada EN VIVO")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()


def test_todo():
    """Prueba: Procesar todo."""
    print("\n" + "=" * 60)
    print("🔥 TEST 4: PROCESAR TODO EN VIVO")
    print("=" * 60)
    
    auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = auth.get_access_token()
    client = GraphServices(token)
    
    print(f"\n📝 Editando archivo: {ARCHIVO}")
    print("   Cliente + Seguimiento\n")
    
    try:
        procesar_excel_live(
            client,
            file_path=ARCHIVO,
            secciones={
                "cliente": datos_cliente,
                "seguimiento": datos_seguimiento
            },
            configuracion=configuracion,
            target_user_id=TARGET_USER_ID
        )
        
        print(f"\n🎉 ¡TODO procesado EN VIVO!")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()


def test_insertar_filas():
    """Prueba: Insertar filas en posición específica."""
    print("\n" + "=" * 60)
    print("🔥 TEST 5: INSERTAR FILAS EN POSICIÓN ESPECÍFICA")
    print("=" * 60)
    
    auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = auth.get_access_token()
    client = GraphServices(token)
    
    print(f"\n📝 Editando archivo: {ARCHIVO}")
    print("   Insertando 3 filas nuevas en la fila 25\n")
    
    # Datos nuevos a insertar
    datos_nuevos = [
        {"fecha": "2025-12-01", "medio": "INSERTADO 1", "comentarios": "Primera fila nueva"},
        {"fecha": "2025-12-02", "medio": "INSERTADO 2", "comentarios": "Segunda fila nueva"},
        {"fecha": "2025-12-03", "medio": "INSERTADO 3", "comentarios": "Tercera fila nueva"}
    ]
    
    try:
        insertar_filas_live(
            client,
            file_path=ARCHIVO,
            fila_inicio=25,  # Inserta en la fila 25
            datos=datos_nuevos,
            columnas={"fecha": 0, "medio": 1, "comentarios": 2},
            target_user_id=TARGET_USER_ID,
            columna_inicio=1,  # Desde la columna A
            merges_a_aplicar=["A:C"]  # Mergear columnas A-C en cada fila nueva
        )
        
        print(f"\n✅ ¡3 filas insertadas correctamente EN VIVO en la fila 25!")
        print(f"   Las filas existentes se movieron hacia abajo")
        print(f"   Los merges de celdas se aplicaron automáticamente")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    # Ejecuta TODAS las pruebas
    
    test_una_celda()
    test_seccion_completa()
    test_tabla()
    test_todo()
    test_insertar_filas()
