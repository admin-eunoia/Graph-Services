"""
Script de prueba para excel_section_writer.py con OneDrive.

Configura aquí las rutas de OneDrive y archivos para probar las funciones.
"""
from Services.excel_section_writer import copiar_template, llenar_seccion, guardar_excel, procesar_excel_completo
from Services.excel_live_writer import llenar_seccion_live, llenar_tabla_live, procesar_excel_live
from Services.graph_services import GraphServices
from Auth.Microsoft_Graph_Auth import MicrosoftGraphAuthenticator
import os

# ==========================================
# CONFIGURACIÓN - EDITA AQUÍ TUS CREDENCIALES
# ==========================================

# Credenciales de Azure AD (puedes usar variables de entorno)
TENANT_ID = os.getenv("MICROSOFT_TENANT_ID", "tu-tenant-id-aqui")
CLIENT_ID = os.getenv("MICROSOFT_CLIENT_ID", "tu-client-id-aqui")
CLIENT_SECRET = os.getenv("MICROSOFT_CLIENT_SECRET")

# Configuración de OneDrive
TARGET_USER_ID = "RodrigoAguilera@Eunoia8.onmicrosoft.com"  # Email del usuario de OneDrive
# O usa DRIVE_ID en lugar de TARGET_USER_ID si prefieres especificar el drive directamente
DRIVE_ID = None  # Ejemplo: "b!abc123..." (opcional, deja None para usar TARGET_USER_ID)

# Rutas en OneDrive (sin barra inicial)
RUTA_TEMPLATE = "Prueba WAMAN/Prueba #1 WAMAN.xlsx"      # Ruta completa del template en OneDrive
RUTA_SALIDA = "Prueba/resultado.xlsx"      # Ruta completa del archivo de salida en OneDrive


# ==========================================
# DATOS DE PRUEBA
# ==========================================

# Ejemplo 1: Datos para sección simple (key-value)
datos_cliente = {
    "nombre": "ACME Corporation",
    "rol": "ACM123456ABC"
}

# Ejemplo 2: Datos para tabla (múltiples filas)
datos_seguimiento = [
    {"fecha": "2025-01-15", "medio": "Email", "comentarios": "Primer contacto"},
    {"fecha": "2025-02-15", "medio": "Teléfono", "comentarios": "Seguimiento"},
    {"fecha": "2025-03-15", "medio": "WhatsApp", "comentarios": "Confirmación"}
]


# ==========================================
# CONFIGURACIÓN DE SECCIONES
# ==========================================

configuracion = {
    "cliente": {
        "marker": "DATOS DEL CLIENTE",     # Texto que buscará en el Excel
        "es_tabla": False,
        "columnas": {
            "nombre": 0,      # Columna A (offset 0)
            "rol": 1          # Columna B (offset 1)
        }
    },
    "seguimiento": {
        "marker": "SEGUIMIENTO",
        "es_tabla": True,
        "columnas": {
            "fecha": 0,       # Columna A
            "medio": 1,       # Columna B
            "comentarios": 2  # Columna C
        }
    }
}


# ==========================================
# FUNCIÓN PRINCIPAL
# ==========================================

def main():
    """Ejecuta la prueba de escritura en Excel con OneDrive."""
    
    print("=" * 60)
    print("PRUEBA DE EXCEL WRITER CON ONEDRIVE")
    print("=" * 60)
    print(f"\n📂 Template OneDrive: {RUTA_TEMPLATE}")
    print(f"📂 Salida OneDrive: {RUTA_SALIDA}")
    print(f"👤 Usuario: {TARGET_USER_ID if TARGET_USER_ID else f'Drive ID: {DRIVE_ID}'}\n")
    
    try:
        # Obtener token de acceso
        print("🔑 Obteniendo token de acceso...")
        auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        token = auth.get_access_token()
        
        # Crear cliente de Graph API
        client = GraphServices(token)
        
        # Descargar template desde OneDrive
        print("📥 Descargando template desde OneDrive...")
        template_bytes, _ = client.download_file_bytes(
            RUTA_TEMPLATE,
            target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
            drive_id=DRIVE_ID
        )
        print(f"   ✓ Descargado: {len(template_bytes)} bytes")
        
        # Procesar Excel - OPCIÓN TODO-EN-UNO
        print("\n✏️  Procesando Excel (TODO-EN-UNO)...")
        output = procesar_excel_completo(
            template_bytes,
            secciones={
                "cliente": datos_cliente,
                "seguimiento": datos_seguimiento
            },
            configuracion=configuracion
        )
        
        # Subir resultado a OneDrive
        print("📤 Subiendo resultado a OneDrive...")
        
        # Intentar subir, si falla por bloqueo, usar rename
        try:
            result, _ = client.upload_file_bytes(
                output.read(),
                RUTA_SALIDA,
                conflict_behavior="replace",  # replace, fail, o rename
                target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
                drive_id=DRIVE_ID
            )
            ruta_final = RUTA_SALIDA
        except Exception as upload_error:
            if "423" in str(upload_error) or "Locked" in str(upload_error):
                print("   ⚠️  Archivo bloqueado, guardando con nombre alternativo...")
                output.seek(0)  # Reiniciar el buffer
                result, _ = client.upload_file_bytes(
                    output.read(),
                    RUTA_SALIDA,
                    conflict_behavior="rename",  # Crear nuevo archivo con nombre diferente
                    target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
                    drive_id=DRIVE_ID
                )
                ruta_final = result.get('name', RUTA_SALIDA)
            else:
                raise
        
        print(f"\n✅ ¡Éxito! Archivo creado en OneDrive:")
        print(f"   {ruta_final}")
        print(f"   ID: {result.get('id', 'N/A')}")
        
    except ValueError as e:
        print(f"\n❌ Error de validación: {e}")
        print("\n💡 Verifica que tu template tenga los marcadores:")
        for seccion, config in configuracion.items():
            print(f"   - '{config['marker']}'")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        print(f"\n💡 Verifica:")
        print(f"   1. La ruta del template existe en OneDrive: {RUTA_TEMPLATE}")
        print(f"   2. El usuario/drive es correcto: {TARGET_USER_ID or DRIVE_ID}")
        print(f"   3. Tienes permisos de lectura/escritura")


# ==========================================
# OPCIÓN B: MÉTODO PASO A PASO (ALTERNATIVA)
# ==========================================

def test_paso_a_paso():
    """
    PRUEBA PASO A PASO: Cada función individual.
    Usa esta opción para probar cada función por separado.
    """
    print("=" * 60)
    print("PRUEBA PASO A PASO - FUNCIONES INDIVIDUALES")
    print("=" * 60)
    
    # Obtener token
    print("\n🔑 Paso 0: Obteniendo token...")
    auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = auth.get_access_token()
    client = GraphServices(token)
    print("   ✓ Token obtenido")
    
    # Descargar template
    print("\n📥 Paso 1: Descargando template desde OneDrive...")
    template_bytes, _ = client.download_file_bytes(
        RUTA_TEMPLATE,
        target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
        drive_id=DRIVE_ID
    )
    print(f"   ✓ Descargado: {len(template_bytes)} bytes")
    
    # FUNCIÓN 1: Copiar template
    print("\n📋 Paso 2: copiar_template()")
    wb = copiar_template(template_bytes)
    print(f"   ✓ Workbook creado, hoja activa: {wb.active.title}")
    
    # FUNCIÓN 2A: Llenar sección cliente (simple)
    print("\n✏️  Paso 3: llenar_seccion() - DATOS DEL CLIENTE (simple)")
    llenar_seccion(
        wb,
        marker="DATOS DEL CLIENTE",
        datos=datos_cliente,
        es_tabla=False,
        columnas={"nombre": 0, "rol": 1}
    )
    print("   ✓ Sección cliente llenada")
    
    # FUNCIÓN 2B: Llenar sección seguimiento (tabla)
    print("\n✏️  Paso 4: llenar_seccion() - SEGUIMIENTO (tabla)")
    llenar_seccion(
        wb,
        marker="SEGUIMIENTO",
        datos=datos_seguimiento,
        es_tabla=True,
        columnas={"fecha": 0, "medio": 1, "comentarios": 2}
    )
    print("   ✓ Sección seguimiento llenada")
    
    # FUNCIÓN 3: Guardar Excel
    print("\n💾 Paso 5: guardar_excel()")
    output = guardar_excel(wb)
    print(f"   ✓ Excel guardado en memoria: {len(output.getvalue())} bytes")
    
    # Subir a OneDrive
    print("\n📤 Paso 6: Subiendo resultado a OneDrive...")
    result, _ = client.upload_file_bytes(
        output.read(),
        RUTA_SALIDA,
        conflict_behavior="replace",
        target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
        drive_id=DRIVE_ID
    )
    
    print(f"\n✅ ¡Éxito! Archivo creado en OneDrive:")
    print(f"   {RUTA_SALIDA}")
    print(f"   ID: {result.get('id', 'N/A')}")


def test_solo_cliente():
    """Prueba solo la sección DATOS DEL CLIENTE."""
    print("=" * 60)
    print("PRUEBA: SOLO DATOS DEL CLIENTE")
    print("=" * 60)
    
    auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = auth.get_access_token()
    client = GraphServices(token)
    
    template_bytes, _ = client.download_file_bytes(
        RUTA_TEMPLATE,
        target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
        drive_id=DRIVE_ID
    )
    
    wb = copiar_template(template_bytes)
    
    print("\n✏️  Llenando DATOS DEL CLIENTE...")
    llenar_seccion(
        wb,
        marker="DATOS DEL CLIENTE",
        datos=datos_cliente,
        es_tabla=False,
        columnas={"nombre": 0, "rol": 1}
    )
    
    output = guardar_excel(wb)
    
    result, _ = client.upload_file_bytes(
        output.read(),
        "Prueba/resultado_solo_cliente.xlsx",
        conflict_behavior="replace",
        target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
        drive_id=DRIVE_ID
    )
    
    print(f"✅ Archivo creado: Prueba/resultado_solo_cliente.xlsx")


def test_solo_seguimiento():
    """Prueba solo la sección SEGUIMIENTO."""
    print("=" * 60)
    print("PRUEBA: SOLO SEGUIMIENTO")
    print("=" * 60)
    
    auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = auth.get_access_token()
    client = GraphServices(token)
    
    template_bytes, _ = client.download_file_bytes(
        RUTA_TEMPLATE,
        target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
        drive_id=DRIVE_ID
    )
    
    wb = copiar_template(template_bytes)
    
    print("\n✏️  Llenando SEGUIMIENTO...")
    llenar_seccion(
        wb,
        marker="SEGUIMIENTO",
        datos=datos_seguimiento,
        es_tabla=True,
        columnas={"fecha": 0, "medio": 1, "comentarios": 2}
    )
    
    output = guardar_excel(wb)
    
    result, _ = client.upload_file_bytes(
        output.read(),
        "Prueba/resultado_solo_seguimiento.xlsx",
        conflict_behavior="replace",
        target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
        drive_id=DRIVE_ID
    )
    
    print(f"✅ Archivo creado: Prueba/resultado_solo_seguimiento.xlsx")


def test_solo_copiar():
    """Prueba solo copiar el template (sin llenar nada)."""
    print("=" * 60)
    print("PRUEBA: SOLO COPIAR TEMPLATE")
    print("=" * 60)
    
    auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = auth.get_access_token()
    client = GraphServices(token)
    
    print("\n📥 Descargando template...")
    template_bytes, _ = client.download_file_bytes(
        RUTA_TEMPLATE,
        target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
        drive_id=DRIVE_ID
    )
    print(f"   ✓ Descargado: {len(template_bytes)} bytes")
    
    print("\n📋 Copiando template...")
    wb = copiar_template(template_bytes)
    print(f"   ✓ Workbook creado")
    print(f"   ✓ Hoja activa: {wb.active.title}")
    print(f"   ✓ Dimensiones: {wb.active.max_row} filas x {wb.active.max_column} columnas")
    
    print("\n💾 Guardando...")
    output = guardar_excel(wb)
    print(f"   ✓ Guardado en memoria: {len(output.getvalue())} bytes")
    
    print("\n📤 Subiendo a OneDrive...")
    try:
        result, _ = client.upload_file_bytes(
            output.read(),
            "Prueba/copia_sin_editar.xlsx",
            conflict_behavior="replace",
            target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
            drive_id=DRIVE_ID
        )
        print(f"✅ Archivo creado: Prueba/copia_sin_editar.xlsx")
    except Exception as e:
        if "423" in str(e):
            output.seek(0)
            result, _ = client.upload_file_bytes(
                output.read(),
                "Prueba/copia_sin_editar.xlsx",
                conflict_behavior="rename",
                target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
                drive_id=DRIVE_ID
            )
            print(f"✅ Archivo creado (renombrado): {result.get('name', 'N/A')}")
        else:
            raise


def test_una_celda():
    """Prueba llenar una sola celda."""
    print("=" * 60)
    print("PRUEBA: LLENAR UNA SOLA CELDA")
    print("=" * 60)
    
    auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = auth.get_access_token()
    client = GraphServices(token)
    
    template_bytes, _ = client.download_file_bytes(
        RUTA_TEMPLATE,
        target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
        drive_id=DRIVE_ID
    )
    
    wb = copiar_template(template_bytes)
    
    print("\n✏️  Llenando solo el campo 'nombre' en DATOS DEL CLIENTE...")
    llenar_seccion(
        wb,
        marker="DATOS DEL CLIENTE",
        datos={"nombre": "Solo este valor"},
        es_tabla=False,
        columnas={"nombre": 0}  # Solo nombre en columna A
    )
    
    output = guardar_excel(wb)
    
    try:
        result, _ = client.upload_file_bytes(
            output.read(),
            "Prueba/resultado_una_celda.xlsx",
            conflict_behavior="replace",
            target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
            drive_id=DRIVE_ID
        )
        print(f"✅ Archivo creado: Prueba/resultado_una_celda.xlsx")
    except Exception as e:
        if "423" in str(e):
            output.seek(0)
            result, _ = client.upload_file_bytes(
                output.read(),
                "Prueba/resultado_una_celda.xlsx",
                conflict_behavior="rename",
                target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
                drive_id=DRIVE_ID
            )
            print(f"✅ Archivo creado (renombrado): {result.get('name', 'N/A')}")
        else:
            raise


def test_muchas_celdas():
    """Prueba llenar muchas celdas en múltiples secciones."""
    print("=" * 60)
    print("PRUEBA: LLENAR MUCHAS CELDAS")
    print("=" * 60)
    
    auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = auth.get_access_token()
    client = GraphServices(token)
    
    template_bytes, _ = client.download_file_bytes(
        RUTA_TEMPLATE,
        target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
        drive_id=DRIVE_ID
    )
    
    wb = copiar_template(template_bytes)
    
    print("\n✏️  Llenando DATOS DEL CLIENTE (2 campos)...")
    llenar_seccion(
        wb,
        marker="DATOS DEL CLIENTE",
        datos={"nombre": "Empresa XYZ", "rol": "Cliente Premium"},
        es_tabla=False,
        columnas={"nombre": 0, "rol": 1}
    )
    
    print("\n✏️  Llenando SEGUIMIENTO (10 filas)...")
    datos_muchos = [
        {"fecha": f"2025-01-{i+1:02d}", "medio": f"Medio {i+1}", "comentarios": f"Comentario largo número {i+1}"}
        for i in range(10)
    ]
    llenar_seccion(
        wb,
        marker="SEGUIMIENTO",
        datos=datos_muchos,
        es_tabla=True,
        columnas={"fecha": 0, "medio": 1, "comentarios": 2}
    )
    
    output = guardar_excel(wb)
    
    try:
        result, _ = client.upload_file_bytes(
            output.read(),
            "Prueba/resultado_muchas_celdas.xlsx",
            conflict_behavior="replace",
            target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
            drive_id=DRIVE_ID
        )
        print(f"✅ Archivo creado: Prueba/resultado_muchas_celdas.xlsx")
    except Exception as e:
        if "423" in str(e):
            output.seek(0)
            result, _ = client.upload_file_bytes(
                output.read(),
                "Prueba/resultado_muchas_celdas.xlsx",
                conflict_behavior="rename",
                target_user_id=TARGET_USER_ID if not DRIVE_ID else None,
                drive_id=DRIVE_ID
            )
            print(f"✅ Archivo creado (renombrado): {result.get('name', 'N/A')}")
        else:
            raise


# ==========================================
# PRUEBAS EN VIVO (sin descargar/subir)
# ==========================================

def test_una_celda_live():
    """🔥 OPCIÓN 8: Llenar una celda EN VIVO usando API."""
    print("=" * 60)
    print("🔥 PRUEBA: LLENAR UNA CELDA EN VIVO")
    print("=" * 60)
    print("Edita el archivo directamente sin descargarlo")
    print("Funciona incluso si alguien lo tiene abierto!")
    print()
    
    # Archivo a editar (debe existir)
    ARCHIVO = "Prueba/archivo_live.xlsx"
    
    auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = auth.get_access_token()
    client = GraphServices(token)
    
    print(f"📝 Editando: {ARCHIVO}")
    
    try:
        llenar_seccion_live(
            client,
            file_path=ARCHIVO,
            marker="DATOS DEL CLIENTE",
            datos={"nombre": "✨ EDITADO EN VIVO"},
            columnas={"nombre": 0},
            target_user_id=TARGET_USER_ID,
            drive_id=DRIVE_ID
        )
        
        print(f"\n✅ ¡Celda editada EN VIVO!")
        print(f"   Abre {ARCHIVO} y verás el cambio inmediatamente")
        
    except ValueError as e:
        print(f"\n❌ Error: {e}")
        print(f"\n💡 Asegúrate de que:")
        print(f"   1. El archivo '{ARCHIVO}' existe")
        print(f"   2. Tiene el marcador 'DATOS DEL CLIENTE'")
    except Exception as e:
        print(f"\n❌ Error: {e}")


def test_tabla_live():
    """🔥 OPCIÓN 9: Llenar tabla EN VIVO usando API."""
    print("=" * 60)
    print("🔥 PRUEBA: LLENAR TABLA EN VIVO")
    print("=" * 60)
    
    ARCHIVO = "Prueba/archivo_live.xlsx"
    
    auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = auth.get_access_token()
    client = GraphServices(token)
    
    print(f"📝 Editando: {ARCHIVO}")
    
    try:
        llenar_tabla_live(
            client,
            file_path=ARCHIVO,
            marker="SEGUIMIENTO",
            datos=[
                {"fecha": "2025-11-22", "medio": "LIVE API", "comentarios": "Editado en vivo!"},
                {"fecha": "2025-11-23", "medio": "Sin descargar", "comentarios": "Magia de API"},
                {"fecha": "2025-11-24", "medio": "Tiempo real", "comentarios": "Funciona abierto"}
            ],
            columnas={"fecha": 0, "medio": 1, "comentarios": 2},
            target_user_id=TARGET_USER_ID,
            drive_id=DRIVE_ID
        )
        
        print(f"\n✅ ¡Tabla llenada EN VIVO!")
        print(f"   {ARCHIVO} se actualizó sin descargarlo")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")


def test_todo_live():
    """🔥 OPCIÓN 10: Procesar todas las secciones EN VIVO."""
    print("=" * 60)
    print("🔥 PRUEBA: PROCESAR TODO EN VIVO")
    print("=" * 60)
    
    ARCHIVO = "Prueba/archivo_live.xlsx"
    
    auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
    token = auth.get_access_token()
    client = GraphServices(token)
    
    print(f"📝 Editando: {ARCHIVO}")
    
    try:
        procesar_excel_live(
            client,
            file_path=ARCHIVO,
            secciones={
                "cliente": datos_cliente,
                "seguimiento": datos_seguimiento
            },
            configuracion=configuracion,
            target_user_id=TARGET_USER_ID,
            drive_id=DRIVE_ID
        )
        
        print(f"\n🎉 ¡TODO editado EN VIVO!")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")


if __name__ == "__main__":
    # Elige UNA de estas opciones:
    
    # OPCIÓN 1: Todo-en-uno (más simple)
    # main()
    
    # OPCIÓN 2: Paso a paso (ver cada función)
    # test_paso_a_paso()
    
    # OPCIÓN 3: Solo cliente
    # test_solo_cliente()
    
    # OPCIÓN 4: Solo seguimiento
    # test_solo_seguimiento()
    
    # OPCIÓN 5: Solo copiar template (sin llenar nada)
    # test_solo_copiar()
    
    # OPCIÓN 6: Llenar una sola celda
    # test_una_celda()
    
    # OPCIÓN 7: Llenar muchas celdas
    # test_muchas_celdas()
    
    # ========================================
    # 🔥 EDICIÓN EN VIVO (sin descargar/subir)
    # ========================================
    
    # OPCIÓN 8: Llenar una celda EN VIVO
    test_una_celda_live()
    
    # OPCIÓN 9: Llenar tabla EN VIVO
    # test_tabla_live()
    
    # OPCIÓN 10: Procesar todo EN VIVO
    # test_todo_live()
