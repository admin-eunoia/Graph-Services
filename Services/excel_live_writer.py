# excel_live_writer.py
"""
Escritura en Excel EN VIVO usando Microsoft Graph API.
Lee toda la configuración desde la base de datos automáticamente.
"""
from typing import Dict, Any, List, Optional, Tuple
from Services.graph_services import GraphServices, _col_index_to_letters
from sqlalchemy.orm import Session
from Postgress.connection import SessionLocal
from Postgress.Tables import (
    TenantCredentials,
    StorageTargets,
    Templates,
    ExcelFiles,
    ExcelSections,
    ExcelFields,
    OperationLogs,
    OperationType,
    RenderStatus
)
from Auth.Microsoft_Graph_Auth import MicrosoftGraphAuthenticator
from datetime import datetime
import uuid


class ExcelLiveWriter:
    """Wrapper para operaciones de Excel usando Graph API con configuración desde DB."""
    
    def __init__(self, client_key: str, correlation_id: str = None):
        """
        Inicializa el writer.
        
        Args:
            client_key: Clave del cliente en tenant_credentials
            correlation_id: ID opcional para tracking
        """
        self.client_key = client_key
        self.correlation_id = correlation_id
        self.db = SessionLocal()
        self.client = self._init_graph_client()
    
    def _init_graph_client(self) -> GraphServices:
        """Inicializa el cliente de Graph API."""
        creds = self.db.query(TenantCredentials).filter_by(
            client_key=self.client_key,
            enabled=True
        ).first()
        
        if not creds:
            raise ValueError(f"Cliente '{self.client_key}' no encontrado")
        
        auth = MicrosoftGraphAuthenticator(
            creds.tenant_id,
            creds.app_client_id,
            creds.app_client_secret
        )
        token = auth.get_access_token()
        return GraphServices(access_token=token, correlation_id=self.correlation_id)
    
    def close(self):
        """Cierra la sesión de base de datos."""
        if self.db:
            self.db.close()
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
    
    def _get_file_context(self, file_key: str = None, section_key: str = None):
        """
        Obtiene contexto completo desde DB.
        
        Returns:
            (excel_file, section, fields, storage, file_path, drive_id, target_user_id)
        """
        excel_file = self._get_file(file_key)
        
        section = None
        fields = None
        if section_key:
            section = self.db.query(ExcelSections).filter_by(
                client_key=self.client_key,
                template_id=excel_file.template_id,
                section_key=section_key,
                is_active=True
            ).first()
            
            if not section:
                raise ValueError(f"Sección '{section_key}' no encontrada")
            
            fields = self.db.query(ExcelFields).filter_by(
                section_id=section.id,
                is_active=True
            ).all()
        
        storage = self.db.query(StorageTargets).filter_by(
            id=excel_file.storage_target_id
        ).first()
        
        if not storage:
            raise ValueError("Storage no encontrado")
        
        file_path = f"{excel_file.file_folder_path}/{excel_file.file_name}".replace("//", "/")
        
        drive_id = None
        target_user_id = None
        if storage.location_type.value.upper() == "DRIVE":
            drive_id = storage.location_identifier
        else:
            target_user_id = storage.location_identifier
        
        return excel_file, section, fields, storage, file_path, drive_id, target_user_id
    
    def _get_template(self, template_key: str = None):
        """Obtiene un template. Si no se especifica template_key, usa el único activo."""
        query = self.db.query(Templates).filter_by(
            client_key=self.client_key,
            is_active=True
        )
        
        if template_key:
            template = query.filter_by(template_key=template_key).first()
            if not template:
                raise ValueError(f"Template '{template_key}' no encontrado")
        else:
            templates = query.all()
            if len(templates) == 0:
                raise ValueError(f"No hay templates activos para '{self.client_key}'")
            elif len(templates) > 1:
                keys = [t.template_key for t in templates]
                raise ValueError(f"Hay {len(templates)} templates activos. Especifica template_key: {keys}")
            template = templates[0]
        
        return template
    
    def _get_file(self, file_key: str = None):
        """Obtiene un archivo. Si no se especifica file_key, usa el más reciente activo."""
        query = self.db.query(ExcelFiles).filter_by(
            client_key=self.client_key,
            is_active=True
        )
        
        if file_key:
            excel_file = query.filter_by(file_key=file_key).first()
            if not excel_file:
                raise ValueError(f"Archivo '{file_key}' no encontrado")
        else:
            excel_file = query.order_by(ExcelFiles.created_at.desc()).first()
            if not excel_file:
                raise ValueError(f"No hay archivos activos para '{self.client_key}'")
        
        return excel_file
    
    def _log_operation(self, op_type: OperationType, excel_file_id: int = None,
                      template_id: int = None, section_id: int = None,
                      sheet_name: str = None, marker_text: str = None,
                      marker_found: bool = None, marker_position: str = None,
                      rows_affected: int = None, cells_affected: int = None,
                      input_data: dict = None, output_data: dict = None,
                      status: RenderStatus = RenderStatus.success,
                      error_message: str = None, error_code: str = None,
                      duration_ms: int = None):
        """Registra una operación en la base de datos."""
        try:
            log = OperationLogs(
                operation_id=str(uuid.uuid4()),
                correlation_id=self.correlation_id,
                client_key=self.client_key,
                template_id=template_id,
                excel_file_id=excel_file_id,
                operation_type=op_type,
                section_id=section_id,
                sheet_name=sheet_name,
                marker_text=marker_text,
                marker_found=marker_found,
                marker_position=marker_position,
                rows_affected=rows_affected,
                cells_affected=cells_affected,
                input_data=input_data,
                output_data=output_data,
                status=status,
                error_message=error_message,
                error_code=error_code,
                duration_ms=duration_ms,
                executed_at=datetime.utcnow()
            )
            self.db.add(log)
            self.db.commit()
        except Exception as e:
            print(f"⚠ Error logging operation: {e}")
            self.db.rollback()
    
    def buscar_marcador(self, file_key: str = None, section_key: str = None) -> Tuple[Optional[int], Optional[int]]:
        """
        Busca un marcador en el Excel.
        
        Args:
            file_key: Clave del archivo (opcional, usa el más reciente si no se especifica)
            section_key: Clave de la sección (opcional, usa la única si solo hay una)
        
        Returns:
            (fila, columna) donde se encontró el marcador, o (None, None)
        """
        start_time = datetime.now()
        excel_file, section, _, _, file_path, drive_id, target_user_id = self._get_file_context(file_key, section_key)
        
        marker = section.marker_text
        sheet_name = section.sheet_name
        
        print(f"🔍 Buscando '{marker}'...")
        
        item_id, _ = self.client._resolve_item_id(file_path, target_user_id=target_user_id, drive_id=drive_id)
        sheets, _ = self.client._resolve_worksheets(item_id=item_id, target_user_id=target_user_id, drive_id=drive_id)
        
        if not sheets:
            raise ValueError("No se encontraron hojas")
        
        if sheet_name:
            sheet = next((s for s in sheets if s.get("name") == sheet_name), None)
            if not sheet:
                raise ValueError(f"Hoja '{sheet_name}' no encontrada")
        else:
            sheet = sheets[0]
        
        ws_id = sheet["id"]
        
        if drive_id:
            base = f"{self.client.graph_url}/drives/{drive_id}/items/{item_id}"
        else:
            base = f"{self.client.graph_url}/users/{target_user_id}/drive/items/{item_id}"
        
        url = f"{base}/workbook/worksheets/{ws_id}/usedRange"
        resp, _ = self.client._request_with_retry("GET", url, expected=(200,), headers=self.client._headers())
        
        data = resp.json()
        values = data.get("values", [])
        row_offset = data.get("rowIndex", 0)
        col_offset = data.get("columnIndex", 0)
        
        for row_idx, row in enumerate(values):
            for col_idx, cell_value in enumerate(row):
                if cell_value and marker in str(cell_value):
                    fila = row_offset + row_idx + 1
                    columna = col_offset + col_idx + 1
                    print(f"   ✓ Encontrado en fila {fila}, columna {columna}")
                    
                    duration = int((datetime.now() - start_time).total_seconds() * 1000)
                    self._log_operation(
                        op_type=OperationType.search_marker,
                        excel_file_id=excel_file.id,
                        section_id=section.id,
                        sheet_name=sheet_name,
                        marker_text=marker,
                        marker_found=True,
                        marker_position=f"{fila},{columna}",
                        status=RenderStatus.success,
                        duration_ms=duration
                    )
                    
                    return (fila, columna)
        
        duration = int((datetime.now() - start_time).total_seconds() * 1000)
        self._log_operation(
            op_type=OperationType.search_marker,
            excel_file_id=excel_file.id,
            section_id=section.id,
            sheet_name=sheet_name,
            marker_text=marker,
            marker_found=False,
            status=RenderStatus.error,
            error_message=f"Marcador '{marker}' no encontrado",
            duration_ms=duration
        )
        
        return (None, None)
    
    def llenar_seccion(self, file_key: str, datos: Dict[str, Any], section_key: str = None):
        """
        Llena una sección simple (key-value).
        
        Args:
            file_key: Clave del archivo a editar
            datos: Diccionario {campo: valor}
            section_key: Clave de la sección (opcional)
        """
        start_time = datetime.now()
        print(f"📝 Llenando sección '{section_key}'...")
        
        excel_file, section, fields, _, file_path, drive_id, target_user_id = self._get_file_context(file_key, section_key)
        
        if not fields:
            raise ValueError(f"No hay campos definidos")
        
        columnas = {field.field_key: field.column_offset for field in fields}
        
        marker_row, marker_col = self.buscar_marcador(file_key, section_key)
        if not marker_row:
            raise ValueError(f"No se encontró '{section.marker_text}'")
        
        item_id, _ = self.client._resolve_item_id(file_path, target_user_id=target_user_id, drive_id=drive_id)
        sheets, _ = self.client._resolve_worksheets(item_id=item_id, target_user_id=target_user_id, drive_id=drive_id)
        
        if section.sheet_name:
            sheet = next((s for s in sheets if s.get("name") == section.sheet_name), None)
        else:
            sheet = sheets[0]
        
        ws_id = sheet["id"]
        ws_name = sheet["name"]
        
        fila_destino = marker_row + section.row_offset
        
        if drive_id:
            base = f"{self.client.graph_url}/drives/{drive_id}/items/{item_id}"
        else:
            base = f"{self.client.graph_url}/users/{target_user_id}/drive/items/{item_id}"
        
        print(f"   Escribiendo {len(datos)} campos...")
        cells_written = 0
        errors = []
        
        for campo, valor in datos.items():
            if campo not in columnas:
                continue
            
            col_offset = columnas[campo]
            col_destino = marker_col + section.column_offset + col_offset
            col_letter = _col_index_to_letters(col_destino)
            
            cell_address = f"{ws_name}!{col_letter}{fila_destino}"
            url = f"{base}/workbook/worksheets/{ws_id}/range(address='{cell_address}')"
            
            try:
                self.client._request_with_retry(
                    "PATCH", url, expected=(200,),
                    headers=self.client._headers(),
                    json={"values": [[valor]]}
                )
                print(f"      ✓ {campo}")
                cells_written += 1
            except Exception as e:
                print(f"      ✗ {campo}: {e}")
                errors.append({"field": campo, "error": str(e)})
        
        duration = int((datetime.now() - start_time).total_seconds() * 1000)
        self._log_operation(
            op_type=OperationType.write_section,
            excel_file_id=excel_file.id,
            section_id=section.id,
            sheet_name=ws_name,
            cells_affected=cells_written,
            input_data={"section_key": section_key, "fields": list(datos.keys())},
            status=RenderStatus.success if not errors else RenderStatus.partial,
            error_message=str(errors) if errors else None,
            duration_ms=duration
        )
    
    def llenar_tabla(self, file_key: str, datos: List[Dict[str, Any]], section_key: str = None):
        """
        Llena una tabla (múltiples filas).
        
        Args:
            file_key: Clave del archivo a editar
            datos: Lista de diccionarios con los datos
            section_key: Clave de la sección (opcional, debe ser is_table=True)
        """
        start_time = datetime.now()
        print(f"📊 Llenando tabla '{section_key}'...")
        
        excel_file, section, fields, _, file_path, drive_id, target_user_id = self._get_file_context(file_key, section_key)
        
        if not section.is_table:
            raise ValueError(f"'{section_key}' no es una tabla")
        
        if not fields:
            raise ValueError("No hay campos definidos")
        
        columnas = {field.field_key: field.column_offset for field in fields}
        
        marker_row, marker_col = self.buscar_marcador(file_key, section_key)
        if not marker_row:
            raise ValueError(f"No se encontró '{section.marker_text}'")
        
        item_id, _ = self.client._resolve_item_id(file_path, target_user_id=target_user_id, drive_id=drive_id)
        sheets, _ = self.client._resolve_worksheets(item_id=item_id, target_user_id=target_user_id, drive_id=drive_id)
        
        if section.sheet_name:
            sheet = next((s for s in sheets if s.get("name") == section.sheet_name), None)
        else:
            sheet = sheets[0]
        
        ws_id = sheet["id"]
        ws_name = sheet["name"]
        
        fila_inicio = marker_row + section.row_offset
        num_filas = len(datos)
        num_columnas = len(columnas)
        
        matriz = [[None] * num_columnas for _ in range(num_filas)]
        
        for row_idx, fila_datos in enumerate(datos):
            for campo, valor in fila_datos.items():
                if campo in columnas:
                    col_offset = columnas[campo]
                    matriz[row_idx][col_offset] = valor
        
        col_inicio_letter = _col_index_to_letters(marker_col + section.column_offset)
        col_fin_letter = _col_index_to_letters(marker_col + section.column_offset + num_columnas - 1)
        fila_fin = fila_inicio + num_filas - 1
        
        range_address = f"{ws_name}!{col_inicio_letter}{fila_inicio}:{col_fin_letter}{fila_fin}"
        
        if drive_id:
            base = f"{self.client.graph_url}/drives/{drive_id}/items/{item_id}"
        else:
            base = f"{self.client.graph_url}/users/{target_user_id}/drive/items/{item_id}"
        
        url = f"{base}/workbook/worksheets/{ws_id}/range(address='{range_address}')"
        
        print(f"   Escribiendo {num_filas} filas...")
        
        try:
            self.client._request_with_retry(
                "PATCH", url, expected=(200,),
                headers=self.client._headers(),
                json={"values": matriz}
            )
            print(f"      ✓ {num_filas} filas escritas")
            
            duration = int((datetime.now() - start_time).total_seconds() * 1000)
            self._log_operation(
                op_type=OperationType.write_table,
                excel_file_id=excel_file.id,
                section_id=section.id,
                sheet_name=ws_name,
                rows_affected=num_filas,
                cells_affected=num_filas * num_columnas,
                input_data={"section_key": section_key, "row_count": num_filas},
                status=RenderStatus.success,
                duration_ms=duration
            )
        except Exception as e:
            duration = int((datetime.now() - start_time).total_seconds() * 1000)
            self._log_operation(
                op_type=OperationType.write_table,
                excel_file_id=excel_file.id,
                section_id=section.id,
                sheet_name=ws_name,
                input_data={"section_key": section_key, "row_count": num_filas},
                status=RenderStatus.error,
                error_message=str(e),
                duration_ms=duration
            )
            print(f"      ✗ Error: {e}")
            raise
        
        if section.merge_ranges and len(section.merge_ranges) > 0:
            try:
                print(f"      Aplicando merges...")
                
                for i in range(num_filas):
                    fila_actual = fila_inicio + i
                    
                    for merge_range in section.merge_ranges:
                        if ":" in merge_range:
                            col_inicio_merge, col_fin_merge = merge_range.split(":")
                            rango_merge = f"{col_inicio_merge}{fila_actual}:{col_fin_merge}{fila_actual}"
                        else:
                            continue
                        
                        merge_url = f"{base}/workbook/worksheets/{ws_id}/range(address='{rango_merge}')/merge"
                        try:
                            self.client._request_with_retry(
                                "POST", merge_url, expected=(200, 204),
                                headers=self.client._headers(),
                                json={"across": True}
                            )
                        except Exception as e_merge:
                            print(f"      ⚠ No se pudo mergear {rango_merge}")
                
                print(f"      ✓ Merges aplicados")
            except Exception as e_merges:
                print(f"      ⚠ Error con merges: {e_merges}")
    
    def insertar_filas(self, file_key: str, datos: List[Dict[str, Any]], section_key: str = None):
        """
        Inserta filas nuevas moviendo las existentes hacia abajo.

        La posición de inserción se determina a partir del marcador definido en la
        sección (como hacen `llenar_seccion` y `llenar_tabla`).

        Args:
            file_key: Clave del archivo a editar
            datos: Lista de diccionarios con los datos
            section_key: Clave de la sección (opcional)
        """
        start_time = datetime.now()
        print(f"➕ Insertando filas en '{section_key}'...")

        excel_file, section, fields, _, file_path, drive_id, target_user_id = self._get_file_context(file_key, section_key)

        if not fields:
            raise ValueError("No hay campos definidos")

        # Obtener marcador y calcular fila de inicio según la definición de la sección
        marker_row, marker_col = self.buscar_marcador(file_key, section_key)
        if not marker_row:
            raise ValueError(f"No se encontró '{section.marker_text}'")

        # La posición donde insertar es marker_row + section.row_offset
        fila_inicio = marker_row + section.row_offset
        print(f"   Calculada fila_inicio desde marcador: {fila_inicio}")
        
        columnas = {field.field_key: field.column_offset for field in fields}
        
        item_id, _ = self.client._resolve_item_id(file_path, target_user_id=target_user_id, drive_id=drive_id)
        sheets, _ = self.client._resolve_worksheets(item_id=item_id, target_user_id=target_user_id, drive_id=drive_id)
        
        if section.sheet_name:
            sheet = next((s for s in sheets if s.get("name") == section.sheet_name), None)
            if not sheet:
                raise ValueError(f"Hoja '{section.sheet_name}' no encontrada")
        else:
            sheet = sheets[0]
        
        ws_id = sheet["id"]
        ws_name = sheet["name"]
        
        if drive_id:
            base = f"{self.client.graph_url}/drives/{drive_id}/items/{item_id}"
        else:
            base = f"{self.client.graph_url}/users/{target_user_id}/drive/items/{item_id}"
        
        num_filas = len(datos)
        num_columnas = len(columnas)
        
        matriz = [[None] * num_columnas for _ in range(num_filas)]
        
        for row_idx, fila_datos in enumerate(datos):
            for campo, valor in fila_datos.items():
                if campo in columnas:
                    col_offset = columnas[campo]
                    matriz[row_idx][col_offset] = valor
        
        columna_inicio = section.column_offset + 1
        col_inicio_letter = _col_index_to_letters(columna_inicio)
        col_fin_letter = _col_index_to_letters(columna_inicio + num_columnas - 1)
        fila_fin = fila_inicio + num_filas - 1
        
        range_simple = f"{col_inicio_letter}{fila_inicio}:{col_fin_letter}{fila_fin}"
        
        print(f"   Insertando {num_filas} filas...")
        
        inserted = False
        try:
            for i in range(num_filas):
                row_range = f"{fila_inicio}:{fila_inicio}"
                insert_url = f"{base}/workbook/worksheets/{ws_id}/range(address='{row_range}')/insert"
                
                self.client._request_with_retry(
                    "POST", insert_url, expected=(200, 201),
                    headers=self.client._headers(),
                    json={"shift": "Down"}
                )
            print(f"      ✓ Filas insertadas")
            inserted = True
        except Exception as e1:
            print(f"      ⚠ Error insertando: {e1}")
        
        if not inserted:
            print(f"      Usando escritura directa...")
            try:
                url = f"{base}/workbook/worksheets/{ws_id}/range(address='{range_simple}')"
                self.client._request_with_retry(
                    "PATCH", url, expected=(200,),
                    headers=self.client._headers(),
                    json={"values": matriz}
                )
                print(f"      ✓ Datos escritos")
                return
            except Exception as e3:
                raise Exception(f"No se pudo insertar: {e3}")
        
        url = f"{base}/workbook/worksheets/{ws_id}/range(address='{range_simple}')"
        
        try:
            self.client._request_with_retry(
                "PATCH", url, expected=(200,),
                headers=self.client._headers(),
                json={"values": matriz}
            )
            print(f"      ✓ Datos escritos")
            
            duration = int((datetime.now() - start_time).total_seconds() * 1000)
            self._log_operation(
                op_type=OperationType.insert_rows,
                excel_file_id=excel_file.id,
                section_id=section.id,
                sheet_name=ws_name,
                rows_affected=num_filas,
                cells_affected=num_filas * num_columnas,
                input_data={"section_key": section_key, "fila_inicio": fila_inicio, "row_count": num_filas},
                status=RenderStatus.success,
                duration_ms=duration
            )
        except Exception as e:
            duration = int((datetime.now() - start_time).total_seconds() * 1000)
            self._log_operation(
                op_type=OperationType.insert_rows,
                excel_file_id=excel_file.id,
                section_id=section.id,
                sheet_name=ws_name,
                input_data={"section_key": section_key, "fila_inicio": fila_inicio, "row_count": num_filas},
                status=RenderStatus.error,
                error_message=str(e),
                duration_ms=duration
            )
            print(f"      ✗ Error: {e}")
            raise
        
        if section.merge_ranges and len(section.merge_ranges) > 0:
            try:
                print(f"      Aplicando merges...")
                
                for i in range(num_filas):
                    fila_actual = fila_inicio + i
                    
                    for merge_range in section.merge_ranges:
                        if ":" in merge_range:
                            col_inicio_merge, col_fin_merge = merge_range.split(":")
                            rango_merge = f"{col_inicio_merge}{fila_actual}:{col_fin_merge}{fila_actual}"
                        else:
                            continue
                        
                        merge_url = f"{base}/workbook/worksheets/{ws_id}/range(address='{rango_merge}')/merge"
                        try:
                            self.client._request_with_retry(
                                "POST", merge_url, expected=(200, 204),
                                headers=self.client._headers(),
                                json={"across": True}
                            )
                        except Exception:
                            pass
                
                print(f"      ✓ Merges aplicados")
            except Exception as e_merges:
                print(f"      ⚠ Error con merges: {e_merges}")
    
    def procesar_excel(self, file_key: str, secciones: Dict[str, Any]):
        """
        Procesa múltiples secciones en un archivo.
        
        Args:
            file_key: Clave del archivo a editar
            secciones: {"section_key": datos}
        """
        print(f"🔥 Procesando Excel '{file_key}'...")
        
        excel_file, _, _, _, _, _, _ = self._get_file_context(file_key)
        
        for section_key, datos in secciones.items():
            print(f"\n📝 Sección: {section_key}")
            
            section = self.db.query(ExcelSections).filter_by(
                client_key=self.client_key,
                template_id=excel_file.template_id,
                section_key=section_key,
                is_active=True
            ).first()
            
            if not section:
                print(f"   ⚠ Sección no encontrada - saltando")
                continue
            
            if section.is_table:
                self.llenar_tabla(file_key, section_key, datos)
            else:
                self.llenar_seccion(file_key, section_key, datos)
        
        print("\n✅ Completado")
    
    def copy_template(self, dest_file_name: str, template_key: str = None, file_key: str = None, context_data: dict = None) -> Tuple[str, str, int]:
        """
        Copia un template sin llenarlo y lo registra en la DB.
        
        Args:
            dest_file_name: Nombre del archivo de destino
            template_key: Clave del template (opcional, usa el único activo si no se especifica)
            file_key: Clave única para el archivo (auto-generada si no se provee)
            context_data: Datos de contexto adicionales (dict)
        
        Returns:
            (item_id, web_url, excel_file_id) del archivo copiado
        """
        start_time = datetime.now()
        
        template = self._get_template(template_key)
        print(f"📋 Copiando template '{template.template_key}'...")
        
        creds = self.db.query(TenantCredentials).filter_by(
            client_key=self.client_key,
            enabled=True
        ).first()
        
        if not creds:
            raise ValueError("Credenciales no encontradas")
        
        storage = self.db.query(StorageTargets).filter_by(
            client_key=self.client_key,
            tenant_id=creds.id
        ).first()
        
        if not storage:
            raise ValueError("Storage no encontrado")
        
        def _join_path(*parts):
            return "/".join(p.strip("/") for p in parts if p)
        
        template_path = _join_path(template.template_folder_path, template.template_file_name)
        dest_path = _join_path(storage.default_dest_folder_path, dest_file_name)
        
        drive_id = None
        target_user_id = None
        
        if storage.location_type.value.upper() == "DRIVE":
            drive_id = storage.location_identifier
        else:
            target_user_id = storage.location_identifier
        
        print(f"   📂 De: {template_path}")
        print(f"   📁 A: {dest_path}")
        
        print(f"   ⬇ Descargando template...")
        template_bytes, _ = self.client.download_file_bytes(
            template_path,
            target_user_id=target_user_id,
            drive_id=drive_id
        )
        
        conflict_behavior = getattr(template, 'default_conflict_behavior', 'rename') or 'rename'
        
        print(f"   ⬆ Copiando...")
        result, _ = self.client.upload_file_bytes(
            template_bytes,
            dest_path,
            conflict_behavior=conflict_behavior,
            target_user_id=target_user_id,
            drive_id=drive_id
        )
        
        item_id = result.get("id")
        web_url = result.get("webUrl")
        
        print(f"   ✅ Copiado")
        print(f"      ID: {item_id}")
        print(f"      URL: {web_url}")
        
        # Registrar archivo en la base de datos
        generated_file_key = file_key or f"{template.template_key}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        new_file = ExcelFiles(
            client_key=self.client_key,
            template_id=template.id,
            storage_target_id=storage.id,
            file_key=generated_file_key,
            file_folder_path=storage.default_dest_folder_path,
            file_name=dest_file_name,
            item_id=item_id,
            web_url=web_url,
            context_data=context_data,
            is_active=True
        )
        
        try:
            self.db.add(new_file)
            self.db.commit()
            self.db.refresh(new_file)
            
            print(f"      📝 Registrado como: {generated_file_key}")
            
            duration = int((datetime.now() - start_time).total_seconds() * 1000)
            self._log_operation(
                op_type=OperationType.copy_template,
                template_id=template.id,
                excel_file_id=new_file.id,
                input_data={"template_key": template.template_key, "dest_file_name": dest_file_name},
                output_data={"item_id": item_id, "web_url": web_url, "file_key": generated_file_key},
                status=RenderStatus.success,
                duration_ms=duration
            )
            
            return item_id, web_url, new_file.id
        except Exception as e:
            self.db.rollback()
            print(f"      ⚠ Error registrando en DB: {e}")
            
            duration = int((datetime.now() - start_time).total_seconds() * 1000)
            self._log_operation(
                op_type=OperationType.copy_template,
                template_id=template.id,
                input_data={"template_key": template.template_key, "dest_file_name": dest_file_name},
                output_data={"item_id": item_id, "web_url": web_url},
                status=RenderStatus.partial,
                error_message=f"Archivo copiado pero no registrado en DB: {e}",
                duration_ms=duration
            )
            
            return item_id, web_url, None
