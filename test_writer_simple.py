"""
Test simple para ExcelLiveWriter - Prueba cada función individualmente
Ahora usa la API simplificada que obtiene todo desde DB con solo client_key
"""
from Services.excel_live_writer import ExcelLiveWriter

# Configuración - Solo necesitas el client_key
CLIENT_KEY = "eunoia"

# Opcionales: solo si tienes múltiples templates/archivos y necesitas especificar
TEMPLATE_KEY = "waman_prueba"  # Opcional si solo tienes un template activo
FILE_KEY = "test_file_002"      # Opcional si solo tienes un archivo activo


def test_1_copy_template():
    """Test 1: Copiar template (usa el único template activo automáticamente)"""
    print("\n" + "="*60)
    print("TEST 1: Copiar Template")
    print("="*60)
    
    with ExcelLiveWriter(client_key=CLIENT_KEY) as writer:
        # Si solo tienes UN template activo, no necesitas template_key
        item_id, web_url, file_id = writer.copy_template(
            dest_file_name="Test_Simple2.xlsx",
            # template_key=TEMPLATE_KEY,  # Opcional
            file_key=FILE_KEY,
            context_data={"test": "simple", "fecha": "2025-11-25"}
        )
        
        print(f"\n✓ Archivo creado:")
        print(f"  - Item ID: {item_id}")
        print(f"  - Web URL: {web_url}")
        print(f"  - File ID en DB: {file_id}")
        
        return item_id, web_url, file_id


def test_2_buscar_marcador():
    """Test 2: Buscar un marcador (usa el archivo más reciente automáticamente)"""
    print("\n" + "="*60)
    print("TEST 2: Buscar Marcador")
    print("="*60)
    
    with ExcelLiveWriter(client_key=CLIENT_KEY) as writer:
        # Si solo especificas section_key, usa el archivo más reciente
        fila, columna = writer.buscar_marcador(
            # file_key=FILE_KEY,  # Opcional si solo tienes un archivo
            section_key="cliente"
        )
        
        print(f"\n✓ Marcador encontrado en:")
        print(f"  - Fila: {fila}")
        print(f"  - Columna: {columna}")
        
        return fila, columna


def test_3_llenar_seccion():
    """Test 3: Llenar una sección"""
    print("\n" + "="*60)
    print("TEST 3: Llenar Sección")
    print("="*60)
    
    datos = {
        "nombre": "Empresa Test SA",
        "rfc": "TEST123456ABC",
        "direccion": "Calle Principal 123",
        "telefono": "555-1234",
        "email": "test@example.com"
    }
    
    with ExcelLiveWriter(client_key=CLIENT_KEY) as writer:
        # Ahora file_key es OBLIGATORIO
        writer.llenar_seccion(
            file_key=FILE_KEY,
            datos=datos,
            section_key="cliente"
        )
        
        print(f"\n✓ Sección llenada con {len(datos)} campos")


def test_4_llenar_tabla():
    """Test 4: Llenar una tabla"""
    print("\n" + "="*60)
    print("TEST 4: Llenar Tabla")
    print("="*60)
    
    datos_tabla = [
        {
            "fecha": "2025-01-15",
            "medio": "Email",
            "comentarios": "Primer contacto establecido"
        },
        {
            "fecha": "2025-02-20",
            "medio": "Teléfono",
            "comentarios": "Seguimiento de propuesta"
        },
        {
            "fecha": "2025-03-10",
            "medio": "Reunión",
            "comentarios": "Cierre de contrato"
        }
    ]
    
    with ExcelLiveWriter(client_key=CLIENT_KEY) as writer:
        writer.llenar_tabla(
            file_key=FILE_KEY,
            datos=datos_tabla,
            section_key="seguimiento"
        )
        
        print(f"\n✓ Tabla llenada con {len(datos_tabla)} filas")


def test_5_insertar_filas():
    """Test 5: Insertar nuevas filas"""
    print("\n" + "="*60)
    print("TEST 5: Insertar Filas")
    print("="*60)
    
    nuevas_filas = [
        {
            "fecha": "2025-04-05",
            "medio": "WhatsApp",
            "comentarios": "Consulta adicional"
        },
        {
            "fecha": "2025-04-12",
            "medio": "Email",
            "comentarios": "Envío de documentación"
        }
    ]
    
    with ExcelLiveWriter(client_key=CLIENT_KEY) as writer:
        writer.insertar_filas(
            file_key=FILE_KEY,
            fila_inicio=25,
            datos=nuevas_filas,
            section_key="seguimiento"
        )
        
        print(f"\n✓ Insertadas {len(nuevas_filas)} filas en posición 25")


def test_6_procesar_excel():
    """Test 6: Procesar múltiples secciones"""
    print("\n" + "="*60)
    print("TEST 6: Procesar Excel Completo")
    print("="*60)
    
    secciones = {
        "cliente": {
            "nombre": "Procesado en Lote SA",
            "rfc": "PROC987654XYZ",
            "direccion": "Av. Secundaria 456",
            "telefono": "555-9876",
            "email": "lote@example.com"
        },
        "seguimiento": [
            {
                "fecha": "2025-05-01",
                "medio": "Zoom",
                "comentarios": "Presentación inicial"
            },
            {
                "fecha": "2025-05-15",
                "medio": "Teams",
                "comentarios": "Demo del producto"
            }
        ]
    }
    
    with ExcelLiveWriter(client_key=CLIENT_KEY) as writer:
        writer.procesar_excel(
            file_key=FILE_KEY,
            secciones=secciones
        )
        
        print(f"\n✓ Procesadas {len(secciones)} secciones")


if __name__ == "__main__":
    print("\n" + "="*60)
    print("TESTS DE ExcelLiveWriter")
    
    try:
        # Comenta/descomenta los tests que quieras ejecutar
        
        # Test 1: Copiar template y crear archivo en DB
        #test_1_copy_template()
        
        # Test 2: Buscar un marcador
        # test_2_buscar_marcador()
        
        # Test 3: Llenar una sección simple
        test_3_llenar_seccion()
        
        # Test 4: Llenar una tabla con múltiples filas
        # test_4_llenar_tabla()
        
        # Test 5: Insertar filas nuevas
        # test_5_insertar_filas()
        
        # Test 6: Procesar múltiples secciones (sobrescribe datos anteriores)
        # test_6_procesar_excel()
        
        print("\n" + "✅"*30)
        print("TESTS COMPLETADOS")
        print("✅"*30)
        
    except Exception as e:
        print("\n" + "❌"*30)
        print(f"ERROR: {e}")
        print("❌"*30)
        raise
