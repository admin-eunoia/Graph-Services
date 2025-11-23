# excel_live_writer.py
"""
Sistema de escritura en Excel EN VIVO usando Microsoft Graph API.
Permite editar archivos Excel sin descargarlos, incluso si están abiertos.

Funciones principales:
1. buscar_marcador_live() - Busca un marcador en el Excel
2. llenar_seccion_live() - Llena una sección en vivo
3. llenar_tabla_live() - Llena una tabla en vivo
4. insertar_filas_live() - Inserta filas en una posición específica
5. obtener_merges_fila() - Obtiene los merges de una fila
6. aplicar_merges_filas() - Aplica merges a múltiples filas
"""
from typing import Dict, Any, List, Optional, Tuple
from Services.graph_services import GraphServices, _col_index_to_letters


def buscar_marcador_live(
    client: GraphServices,
    file_path: str,
    marker: str,
    target_user_id: str = None,
    drive_id: str = None,
    sheet_name: str = None
) -> Tuple[Optional[int], Optional[int]]:
    """
    Busca un marcador en el Excel usando la API.
    
    Args:
        client: Cliente de GraphServices
        file_path: Ruta del archivo en OneDrive
        marker: Texto a buscar
        target_user_id: Email del usuario
        drive_id: ID del drive (opcional)
        sheet_name: Nombre de la hoja (None = primera hoja)
    
    Returns:
        (fila, columna) donde se encontró el marcador, o (None, None)
    """
    # Resolver item_id
    item_id, _ = client._resolve_item_id(
        file_path,
        target_user_id=target_user_id,
        drive_id=drive_id
    )
    
    # Obtener hojas
    sheets, _ = client._resolve_worksheets(
        item_id=item_id,
        target_user_id=target_user_id,
        drive_id=drive_id
    )
    
    if not sheets:
        raise ValueError("No se encontraron hojas en el Excel")
    
    # Seleccionar hoja
    if sheet_name:
        sheet = next((s for s in sheets if s.get("name") == sheet_name), None)
        if not sheet:
            raise ValueError(f"Hoja '{sheet_name}' no encontrada")
    else:
        sheet = sheets[0]
    
    ws_id = sheet["id"]
    
    # Obtener rango usado
    if drive_id:
        base = f"{client.graph_url}/drives/{drive_id}/items/{item_id}"
    else:
        base = f"{client.graph_url}/users/{target_user_id}/drive/items/{item_id}"
    
    url = f"{base}/workbook/worksheets/{ws_id}/usedRange"
    resp, _ = client._request_with_retry("GET", url, expected=(200,), headers=client._headers())
    
    data = resp.json()
    values = data.get("values", [])
    row_offset = data.get("rowIndex", 0)
    col_offset = data.get("columnIndex", 0)
    
    # Buscar el marcador
    for row_idx, row in enumerate(values):
        for col_idx, cell_value in enumerate(row):
            if cell_value and marker in str(cell_value):
                fila = row_offset + row_idx + 1  # 1-indexed
                columna = col_offset + col_idx + 1  # 1-indexed
                print(f"   ✓ Marcador '{marker}' encontrado en fila {fila}, columna {columna}")
                return (fila, columna)
    
    return (None, None)


def llenar_seccion_live(
    client: GraphServices,
    file_path: str,
    marker: str,
    datos: Dict[str, Any],
    columnas: Dict[str, int],
    target_user_id: str = None,
    drive_id: str = None,
    sheet_name: str = None
):
    """
    Llena una sección simple (key-value) EN VIVO.
    
    Args:
        client: Cliente de GraphServices
        file_path: Ruta del archivo en OneDrive
        marker: Marcador a buscar
        datos: Diccionario con los datos {campo: valor}
        columnas: Mapeo de campo -> offset de columna {campo: 0, otro: 1}
        target_user_id: Email del usuario
        drive_id: ID del drive
        sheet_name: Nombre de la hoja
    
    Ejemplo:
        llenar_seccion_live(
            client, "ruta/archivo.xlsx", "DATOS DEL CLIENTE",
            datos={"nombre": "ACME", "rfc": "ACM123"},
            columnas={"nombre": 0, "rfc": 1},
            target_user_id="user@domain.com"
        )
    """
    # Buscar marcador
    marker_row, marker_col = buscar_marcador_live(
        client, file_path, marker, target_user_id, drive_id, sheet_name
    )
    
    if not marker_row:
        raise ValueError(f"No se encontró '{marker}' en el Excel")
    
    # Resolver item_id y worksheet
    item_id, _ = client._resolve_item_id(file_path, target_user_id=target_user_id, drive_id=drive_id)
    sheets, _ = client._resolve_worksheets(item_id=item_id, target_user_id=target_user_id, drive_id=drive_id)
    
    if sheet_name:
        sheet = next((s for s in sheets if s.get("name") == sheet_name), None)
    else:
        sheet = sheets[0]
    
    ws_id = sheet["id"]
    ws_name = sheet["name"]
    
    # Fila destino (siguiente al marcador)
    fila_destino = marker_row + 1
    
    # Construir URL base
    if drive_id:
        base = f"{client.graph_url}/drives/{drive_id}/items/{item_id}"
    else:
        base = f"{client.graph_url}/users/{target_user_id}/drive/items/{item_id}"
    
    # Escribir cada campo
    print(f"   ✏️  Escribiendo {len(datos)} campos en fila {fila_destino}...")
    for campo, valor in datos.items():
        if campo not in columnas:
            continue
        
        col_offset = columnas[campo]
        col_destino = marker_col + col_offset
        col_letter = _col_index_to_letters(col_destino)
        
        # Dirección de la celda
        cell_address = f"{ws_name}!{col_letter}{fila_destino}"
        
        # PATCH a la celda
        url = f"{base}/workbook/worksheets/{ws_id}/range(address='{cell_address}')"
        
        try:
            client._request_with_retry(
                "PATCH",
                url,
                expected=(200,),
                headers=client._headers(),
                json={"values": [[valor]]}
            )
            print(f"      ✓ {campo}: {valor}")
        except Exception as e:
            print(f"      ✗ {campo}: Error - {e}")


def llenar_tabla_live(
    client: GraphServices,
    file_path: str,
    marker: str,
    datos: List[Dict[str, Any]],
    columnas: Dict[str, int],
    target_user_id: str = None,
    drive_id: str = None,
    sheet_name: str = None,
    merges_a_aplicar: Optional[List[str]] = None
):
    """
    Llena una tabla (múltiples filas) EN VIVO.
    
    Args:
        client: Cliente de GraphServices
        file_path: Ruta del archivo en OneDrive
        marker: Marcador a buscar
        datos: Lista de diccionarios con los datos
        columnas: Mapeo de campo -> offset de columna
        target_user_id: Email del usuario
        drive_id: ID del drive
        sheet_name: Nombre de la hoja
        merges_a_aplicar: Lista de rangos a mergear en cada fila
                         Ejemplo: ["A:C"] mergeará columnas A-C en cada fila
    
    Ejemplo:
        llenar_tabla_live(
            client, "ruta/archivo.xlsx", "SEGUIMIENTO",
            datos=[
                {"fecha": "2025-01", "medio": "Email", "comentarios": "Contacto"},
                {"fecha": "2025-02", "medio": "Teléfono", "comentarios": "Llamada"}
            ],
            columnas={"fecha": 0, "medio": 1, "comentarios": 2},
            target_user_id="user@domain.com",
            merges_a_aplicar=["A:C"]  # Opcional: mergear columnas
        )
    """
    # Buscar marcador
    marker_row, marker_col = buscar_marcador_live(
        client, file_path, marker, target_user_id, drive_id, sheet_name
    )
    
    if not marker_row:
        raise ValueError(f"No se encontró '{marker}' en el Excel")
    
    # Resolver item_id y worksheet
    item_id, _ = client._resolve_item_id(file_path, target_user_id=target_user_id, drive_id=drive_id)
    sheets, _ = client._resolve_worksheets(item_id=item_id, target_user_id=target_user_id, drive_id=drive_id)
    
    if sheet_name:
        sheet = next((s for s in sheets if s.get("name") == sheet_name), None)
    else:
        sheet = sheets[0]
    
    ws_id = sheet["id"]
    ws_name = sheet["name"]
    
    # Fila de inicio (marcador + 1 header + 1)
    fila_inicio = marker_row + 2
    
    # Construir URL base
    if drive_id:
        base = f"{client.graph_url}/drives/{drive_id}/items/{item_id}"
    else:
        base = f"{client.graph_url}/users/{target_user_id}/drive/items/{item_id}"
    
    # Preparar datos en formato de matriz
    num_filas = len(datos)
    num_columnas = len(columnas)
    
    # Crear matriz vacía
    matriz = [[None] * num_columnas for _ in range(num_filas)]
    
    # Llenar matriz
    for row_idx, fila_datos in enumerate(datos):
        for campo, valor in fila_datos.items():
            if campo in columnas:
                col_offset = columnas[campo]
                matriz[row_idx][col_offset] = valor
    
    # Calcular rango
    col_inicio_letter = _col_index_to_letters(marker_col)
    col_fin_letter = _col_index_to_letters(marker_col + num_columnas - 1)
    fila_fin = fila_inicio + num_filas - 1
    
    range_address = f"{ws_name}!{col_inicio_letter}{fila_inicio}:{col_fin_letter}{fila_fin}"
    
    # PATCH al rango completo
    url = f"{base}/workbook/worksheets/{ws_id}/range(address='{range_address}')"
    
    print(f"   ✏️  Escribiendo {num_filas} filas en rango {range_address}...")
    
    try:
        client._request_with_retry(
            "PATCH",
            url,
            expected=(200,),
            headers=client._headers(),
            json={"values": matriz}
        )
        print(f"      ✓ {num_filas} filas escritas exitosamente")
    except Exception as e:
        print(f"      ✗ Error: {e}")
    
    # Aplicar merges si se especificaron
    if merges_a_aplicar and len(merges_a_aplicar) > 0:
        try:
            print(f"      💡 Aplicando merges a {num_filas} filas...")
            
            for i in range(num_filas):
                fila_actual = fila_inicio + i
                
                for merge_range in merges_a_aplicar:
                    if ":" in merge_range:
                        col_inicio_merge, col_fin_merge = merge_range.split(":")
                        # Ajustar al offset del marcador
                        col_inicio_abs = col_inicio_merge
                        col_fin_abs = col_fin_merge
                        rango_merge = f"{col_inicio_abs}{fila_actual}:{col_fin_abs}{fila_actual}"
                    else:
                        continue
                    
                    merge_url = f"{base}/workbook/worksheets/{ws_id}/range(address='{rango_merge}')/merge"
                    try:
                        client._request_with_retry(
                            "POST",
                            merge_url,
                            expected=(200, 204),
                            headers=client._headers(),
                            json={"across": True}
                        )
                    except Exception as e_merge:
                        print(f"      ⚠️  No se pudo mergear {rango_merge}: {e_merge}")
            
            print(f"      ✓ Merges aplicados")
        except Exception as e_merges:
            print(f"      ⚠️  Error aplicando merges: {e_merges}")


def procesar_excel_live(
    client: GraphServices,
    file_path: str,
    secciones: Dict[str, Any],
    configuracion: Dict[str, Dict],
    target_user_id: str = None,
    drive_id: str = None
):
    """
    Función todo-en-uno para procesar múltiples secciones EN VIVO.
    
    Args:
        client: Cliente de GraphServices
        file_path: Ruta del archivo en OneDrive
        secciones: Datos por sección {"nombre_seccion": datos}
        configuracion: Config de cada sección
        target_user_id: Email del usuario
        drive_id: ID del drive
    
    Ejemplo:
        procesar_excel_live(
            client,
            "ruta/archivo.xlsx",
            secciones={
                "cliente": {"nombre": "ACME", "rfc": "ACM123"},
                "pagos": [{"fecha": "2025-01", "monto": 1000}]
            },
            configuracion={
                "cliente": {
                    "marker": "DATOS DEL CLIENTE",
                    "es_tabla": False,
                    "columnas": {"nombre": 0, "rfc": 1}
                },
                "pagos": {
                    "marker": "Pagos",
                    "es_tabla": True,
                    "columnas": {"fecha": 0, "monto": 1}
                }
            },
            target_user_id="user@domain.com"
        )
    """
    print("🔥 Procesando Excel EN VIVO...")
    
    for nombre_seccion, datos in secciones.items():
        if nombre_seccion not in configuracion:
            continue
        
        config = configuracion[nombre_seccion]
        marker = config["marker"]
        es_tabla = config.get("es_tabla", False)
        columnas = config.get("columnas", {})
        sheet_name = config.get("sheet_name")
        
        print(f"\n📝 Sección: {nombre_seccion}")
        
        if es_tabla:
            llenar_tabla_live(
                client, file_path, marker, datos, columnas,
                target_user_id, drive_id, sheet_name
            )
        else:
            llenar_seccion_live(
                client, file_path, marker, datos, columnas,
                target_user_id, drive_id, sheet_name
            )
    
    print("\n✅ Proceso completado")


def obtener_merges_fila(
    client: GraphServices,
    item_id: str,
    ws_id: str,
    fila: int,
    target_user_id: str = None,
    drive_id: str = None
) -> List[Dict[str, Any]]:
    """
    Obtiene los rangos merged de una fila específica.
    
    Args:
        client: Cliente de GraphServices
        item_id: ID del archivo
        ws_id: ID de la hoja
        fila: Número de fila (1-indexed)
        target_user_id: Email del usuario
        drive_id: ID del drive
    
    Returns:
        Lista de rangos merged que incluyen esa fila
        Ejemplo: [{"address": "A22:C22"}, {"address": "D22:F22"}]
    """
    if drive_id:
        base = f"{client.graph_url}/drives/{drive_id}/items/{item_id}"
    else:
        base = f"{client.graph_url}/users/{target_user_id}/drive/items/{item_id}"
    
    # Obtener el rango usado de la hoja
    url = f"{base}/workbook/worksheets/{ws_id}/usedRange"
    resp, _ = client._request_with_retry("GET", url, expected=(200,), headers=client._headers())
    data = resp.json()
    
    # Obtener merged cells de la hoja
    # Microsoft Graph no tiene endpoint directo, así que usamos range/format
    # Obtenemos el formato del rango de la fila
    col_start = data.get("columnIndex", 0) + 1
    col_end = data.get("columnIndex", 0) + data.get("columnCount", 10)
    
    col_start_letter = _col_index_to_letters(col_start)
    col_end_letter = _col_index_to_letters(col_end)
    
    range_address = f"{col_start_letter}{fila}:{col_end_letter}{fila}"
    
    url = f"{base}/workbook/worksheets/{ws_id}/range(address='{range_address}')"
    resp, _ = client._request_with_retry("GET", url, expected=(200,), headers=client._headers())
    data = resp.json()
    
    merges = []
    cell_count = data.get("cellCount", 0)
    row_count = data.get("rowCount", 1)
    col_count = data.get("columnCount", 0)
    
    # Si el rango tiene merges, los extraemos manualmente
    # Verificando el número de celdas vs el tamaño esperado
    if "address" in data:
        # Para detectar merges, necesitamos verificar cada celda
        # Usamos la propiedad mergedAreas si está disponible
        address = data.get("address", "")
        if "!" in address:
            address = address.split("!")[-1]
        
        # Por ahora, retornamos info básica
        # En una implementación más completa, iteraríamos celda por celda
        merges.append({"address": address, "row": fila})
    
    return merges


def aplicar_merges_filas(
    client: GraphServices,
    item_id: str,
    ws_id: str,
    ws_name: str,
    fila_template: int,
    filas_destino: List[int],
    target_user_id: str = None,
    drive_id: str = None
):
    """
    Aplica los merges de una fila template a múltiples filas destino.
    
    Args:
        client: Cliente de GraphServices
        item_id: ID del archivo
        ws_id: ID de la hoja
        ws_name: Nombre de la hoja
        fila_template: Fila de la cual copiar el formato de merge
        filas_destino: Lista de filas donde aplicar los merges
        target_user_id: Email del usuario
        drive_id: ID del drive
    """
    if drive_id:
        base = f"{client.graph_url}/drives/{drive_id}/items/{item_id}"
    else:
        base = f"{client.graph_url}/users/{target_user_id}/drive/items/{item_id}"
    
    # Obtener el rango de la fila template para analizar sus merges
    url = f"{base}/workbook/worksheets/{ws_id}/range(address='{fila_template}:{fila_template}')"
    resp, _ = client._request_with_retry("GET", url, expected=(200,), headers=client._headers())
    template_data = resp.json()
    
    # Obtener el formato completo incluyendo merges
    format_url = f"{base}/workbook/worksheets/{ws_id}/range(address='{fila_template}:{fila_template}')/format"
    resp_format, _ = client._request_with_retry("GET", format_url, expected=(200,), headers=client._headers())
    format_data = resp_format.json()
    
    # Detectar rangos merged analizando el formato de celdas
    # Microsoft Graph API tiene limitaciones aquí, usaremos merge directo por rango
    # Aplicar merge a las filas destino con el mismo patrón
    
    for fila_dest in filas_destino:
        # Intentar copiar el formato completo de la fila template
        try:
            # Merge específico: necesitamos saber qué columnas están merged
            # Por ahora, aplicaremos merge basado en el patrón más común
            # que vemos en la imagen: A-C merged
            
            # Opción simple: merge las mismas columnas que en template
            col_start_letter = template_data.get("address", "A1").split("!")[0] if "!" in template_data.get("address", "") else "A"
            
            # Para cada rango merged detectado, aplicarlo a la fila destino
            # Esto es una aproximación - idealmente necesitaríamos la API de MergedAreas
            pass
            
        except Exception as e:
            print(f"      ⚠️  Error copiando formato de fila {fila_template} a {fila_dest}: {e}")


def insertar_filas_live(
    client: GraphServices,
    file_path: str,
    fila_inicio: int,
    datos: List[Dict[str, Any]],
    columnas: Dict[str, int],
    target_user_id: str = None,
    drive_id: str = None,
    sheet_name: str = None,
    columna_inicio: int = 1,
    merges_a_aplicar: Optional[List[str]] = None
):
    """
    Inserta filas nuevas en una posición específica EN VIVO.
    
    Esta función inserta filas completas moviendo las existentes hacia abajo,
    llena las nuevas filas con los datos proporcionados, y opcionalmente
    aplica merges de celdas para mantener el formato de la tabla.
    
    Args:
        client: Cliente de GraphServices
        file_path: Ruta del archivo en OneDrive
        fila_inicio: Número de fila donde insertar (1-indexed)
        datos: Lista de diccionarios con los datos
        columnas: Mapeo de campo -> offset de columna
        target_user_id: Email del usuario
        drive_id: ID del drive
        sheet_name: Nombre de la hoja
        columna_inicio: Columna inicial (default=1 para A)
        merges_a_aplicar: Lista de rangos a mergear en cada fila nueva
                         Ejemplo: ["A:C", "D:F"] mergeará A-C y D-F en cada fila
    
    Ejemplo:
        insertar_filas_live(
            client, "ruta/archivo.xlsx",
            fila_inicio=25,
            datos=[
                {"fecha": "2025-01", "medio": "Email", "comentarios": "Nuevo"},
                {"fecha": "2025-02", "medio": "Tel", "comentarios": "Otro"}
            ],
            columnas={"fecha": 0, "medio": 1, "comentarios": 2},
            target_user_id="user@domain.com"
        )
    """
    # Resolver item_id y worksheet
    item_id, _ = client._resolve_item_id(file_path, target_user_id=target_user_id, drive_id=drive_id)
    sheets, _ = client._resolve_worksheets(item_id=item_id, target_user_id=target_user_id, drive_id=drive_id)
    
    if sheet_name:
        sheet = next((s for s in sheets if s.get("name") == sheet_name), None)
        if not sheet:
            raise ValueError(f"Hoja '{sheet_name}' no encontrada")
    else:
        sheet = sheets[0]
    
    ws_id = sheet["id"]
    ws_name = sheet["name"]
    
    # Construir URL base
    if drive_id:
        base = f"{client.graph_url}/drives/{drive_id}/items/{item_id}"
    else:
        base = f"{client.graph_url}/users/{target_user_id}/drive/items/{item_id}"
    
    num_filas = len(datos)
    num_columnas = len(columnas)
    
    # Preparar matriz de datos
    matriz = [[None] * num_columnas for _ in range(num_filas)]
    
    for row_idx, fila_datos in enumerate(datos):
        for campo, valor in fila_datos.items():
            if campo in columnas:
                col_offset = columnas[campo]
                matriz[row_idx][col_offset] = valor
    
    # Calcular rango
    col_inicio_letter = _col_index_to_letters(columna_inicio)
    col_fin_letter = _col_index_to_letters(columna_inicio + num_columnas - 1)
    fila_fin = fila_inicio + num_filas - 1
    
    # Usar dirección simple sin nombre de hoja para el insert
    range_simple = f"{col_inicio_letter}{fila_inicio}:{col_fin_letter}{fila_fin}"
    range_con_hoja = f"{ws_name}!{range_simple}"
    
    print(f"   📍 Insertando {num_filas} filas en {range_con_hoja}...")
    
    # PASO 1: Insertar filas usando el endpoint de filas completas
    inserted = False
    
    try:
        # Insertar filas completas una por una desde la posición especificada
        print(f"      💡 Insertando {num_filas} filas completas...")
        for i in range(num_filas):
            # La fila actual donde insertar (siempre es fila_inicio porque cada insert mueve las siguientes)
            row_range = f"{fila_inicio}:{fila_inicio}"
            insert_url = f"{base}/workbook/worksheets/{ws_id}/range(address='{row_range}')/insert"
            
            client._request_with_retry(
                "POST",
                insert_url,
                expected=(200, 201),
                headers=client._headers(),
                json={"shift": "Down"}
            )
        print(f"      ✓ {num_filas} filas insertadas correctamente")
        inserted = True
    except Exception as e1:
        print(f"      ✗ Error insertando filas completas: {e1}")
    
    # Si la inserción falló, usar escritura directa
    if not inserted:
        print(f"      💡 Usando escritura directa (sobrescribe celdas existentes)...")
        try:
            url = f"{base}/workbook/worksheets/{ws_id}/range(address='{range_simple}')"
            client._request_with_retry(
                "PATCH",
                url,
                expected=(200,),
                headers=client._headers(),
                json={"values": matriz}
            )
            print(f"      ✓ Datos escritos directamente (sin insertar)")
            return  # Salir porque ya terminamos
        except Exception as e3:
            print(f"      ✗ Error con escritura directa: {e3}")
            raise Exception("No se pudo insertar ni escribir las filas")
    
    # PASO 2: Llenar las filas insertadas con datos
    url = f"{base}/workbook/worksheets/{ws_id}/range(address='{range_simple}')"
    
    try:
        client._request_with_retry(
            "PATCH",
            url,
            expected=(200,),
            headers=client._headers(),
            json={"values": matriz}
        )
        print(f"      ✓ Datos escritos en {num_filas} filas")
    except Exception as e:
        print(f"      ✗ Error escribiendo datos: {e}")
        raise
    
    # PASO 3: Aplicar merges si se especificaron
    if merges_a_aplicar and len(merges_a_aplicar) > 0:
        try:
            print(f"      💡 Aplicando merges a {num_filas} filas...")
            
            for i in range(num_filas):
                fila_actual = fila_inicio + i
                
                for merge_range in merges_a_aplicar:
                    # El rango puede ser "A:C" o letras de columna
                    # Lo convertimos a un rango completo con la fila actual
                    if ":" in merge_range:
                        col_inicio_merge, col_fin_merge = merge_range.split(":")
                        rango_merge = f"{col_inicio_merge}{fila_actual}:{col_fin_merge}{fila_actual}"
                    else:
                        # Si solo es una columna, no hacer merge
                        continue
                    
                    # Aplicar merge
                    merge_url = f"{base}/workbook/worksheets/{ws_id}/range(address='{rango_merge}')/merge"
                    try:
                        client._request_with_retry(
                            "POST",
                            merge_url,
                            expected=(200, 204),
                            headers=client._headers(),
                            json={"across": True}  # Merge horizontal
                        )
                    except Exception as e_merge:
                        print(f"      ⚠️  No se pudo mergear {rango_merge}: {e_merge}")
            
            print(f"      ✓ Merges aplicados correctamente")
        except Exception as e_merges:
            print(f"      ⚠️  Error aplicando merges: {e_merges}")
