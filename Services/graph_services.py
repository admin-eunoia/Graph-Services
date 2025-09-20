from typing import Optional, Any
import logging
from Auth.Microsoft_Graph_Auth import MicrosoftGraphAuthenticator
import requests
from excel_render import fill_cells_in_memory, EXCEL_MIME
from io import BytesIO
from typing import Optional, Dict, Any


class GraphServices:
    def __init__(self, user_id: Optional[str] = None):
        self.auth = MicrosoftGraphAuthenticator()
        self.user_id = user_id or self.auth.user_id
        self.graph_url = "https://graph.microsoft.com/v1.0"

    # -------------------------- Helpers internos (Graph) -------------------------- #
    def _join_path(self, folder_path: str, file_name: str) -> str:
        """Une carpeta + archivo al formato que Graph espera en root:/...:/content."""
        return f"{folder_path.strip('/')}/{file_name}"

    def _download_file_bytes(self, full_path: str) -> bytes:
        """
        Descarga el contenido binario de un archivo en OneDrive usando ruta relativa.
        full_path: 'Carpeta/Subcarpeta/Archivo.xlsx'
        """
        headers = self.auth.get_headers()
        url = f"{self.graph_url}/users/{self.user_id}/drive/root:/{full_path}:/content"
        resp = requests.get(url, headers=headers, timeout=60)
        if resp.status_code != 200:
            raise Exception(f"❌ Error al descargar '{full_path}': {resp.status_code} - {resp.text}")
        return resp.content

    def _upload_bytes_fail(self, full_dest_path: str, file_bytes: bytes) -> Dict[str, Any]:
        """
        Sube bytes a OneDrive con conflictBehavior=fail (no sobrescribe si existe).
        Retorna el driveItem (JSON).
        """
        headers = self.auth.get_headers().copy()
        headers["Content-Type"] = EXCEL_MIME
        url = (
            f"{self.graph_url}/users/{self.user_id}/drive/root:/{full_dest_path}"
            f":/content?@microsoft.graph.conflictBehavior=fail"
        )
        resp = requests.put(url, headers=headers, data=file_bytes, timeout=60)
        if not (200 <= resp.status_code < 300):
            raise Exception(f"❌ Error al subir '{full_dest_path}': {resp.status_code} - {resp.text}")
        return resp.json()

    # -------------------------- Método público (orquestador) ---------------------- #
    def fill_excel_from_template(
        self,
        *,
        template_folder_path: str,
        template_file_name: str,
        dest_folder_path: str,
        dest_file_name: str,
        data: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        1) Descarga el template (.xlsx) desde OneDrive (RAM)
        2) Rellena celdas en RAM con openpyxl
        3) Sube el archivo final a OneDrive con conflictBehavior=fail

        :param template_folder_path: Carpeta del template (p.ej. "Templates")
        :param template_file_name:   Archivo de template (p.ej. "Factura.xlsx")
        :param dest_folder_path:     Carpeta destino (p.ej. "Clientes/Acme")
        :param dest_file_name:       Archivo destino (p.ej. "Factura_123.xlsx")
        :param data:                 {"A1": 10, "Hoja1!B3": "ACME", ...}
        :return:                     driveItem JSON (incluye webUrl)
        """
        # 1) Descargar template
        template_path = self._join_path(template_folder_path, template_file_name)
        template_bytes = self._download_file_bytes(template_path)

        # 2) Rellenar en memoria
        out_buf: BytesIO = fill_cells_in_memory(template_bytes, data)

        # 3) Subir a destino con fail (no sobrescribe si existe)
        dest_path = self._join_path(dest_folder_path, dest_file_name)
        drive_item = self._upload_bytes_fail(dest_path, out_buf.getvalue())
        return drive_item

        
    def copy_excel(self, name) -> None:
        """
        Copia un archivo Excel en OneDrive.

        :param template_name: Nombre del archivo origen dentro de la carpeta (p.ej. "Plantilla.xlsx")
        :param copy_template_name: Nombre del archivo destino (p.ej. "MiCopia.xlsx")
        :param folder_path: Ruta relativa de la carpeta en OneDrive (p.ej. "Documentos/Proyectos")
        """
        headers = self.auth.get_headers()

        template_name = get_template_name()
        copy_template_name = get_copy_template_name()
        folder_path = get_folder_path()

        ruta_archivo_origen = f"{folder_path}/{template_name}"
        folder_id = self.auth.get_folder_id(folder_path)
        source_file_id = self.auth.get_file_id(ruta_archivo_origen)
        
        url = f"{self.graph_url}/users/{self.user_id}/drive/items/{source_file_id}/copy"
        
        body = {
            "name": copy_template_name,
            "parentReference": {"id": folder_id},
        }
        
        response = requests.post(url, headers=headers, json=body)
        if response.status_code == 202:
            print(f"✅ Copia iniciada: '{template_name}' ➜ '{copy_template_name}'")
        else:
            raise Exception(
                f"❌ Error al copiar archivo: {response.status_code} - {response.text}"
            )
            
    def fill_excel(self, file_path: str, worksheet_name: str, data: dict[str, Any]) -> None:
        """
        Escribe pares celda:valor en una hoja de Excel existente.

        :param file_path: Ruta relativa al archivo en OneDrive (p.ej. "Documentos/MiCopia.xlsx")
        :param worksheet_name: Nombre de la hoja (p.ej. "Hoja1")
        :param data: Diccionario con pares celda: valor, p.ej. {"A1": "Juan", "B2": 30}
        """
        file_id = self.auth.get_file_id(file_path)
        headers = self.auth.get_headers()

        for celda, valor in data.items():
            url = f"{self.graph_url}/users/{self.user_id}/drive/items/{file_id}/workbook/worksheets/{worksheet_name}/range(address='{celda}')"
            body = {"values": [[valor]]}

            response = requests.patch(url, headers=headers, json=body)

            if response.status_code != 200:
                raise Exception(f"❌ Error escribiendo en celda {celda}: {response.status_code} - {response.text}")

    # -------------------------- Internal helpers -------------------------- #
    def _request(self, method: str, url: str, expected: int | tuple[int, ...], **kwargs) -> requests.Response:
        """Internal wrapper around requests to centralize error handling.

        :param method: HTTP method (GET/POST/PATCH/DELETE)
        :param url: Full URL
        :param expected: int or tuple of acceptable status codes
        :param kwargs: passed to requests.request
        :raises Exception: on unexpected status code
        :return: Response object
        """
        headers = kwargs.pop("headers", None) or self.auth.get_headers()
        resp = requests.request(method, url, headers=headers, **kwargs)
        ok_codes: tuple[int, ...] = (expected,) if isinstance(expected, int) else expected
        if resp.status_code not in ok_codes:
            raise Exception(f"❌ Error {method} {url}: {resp.status_code} - {resp.text}")
        return resp

    # -------------------------- File level operations --------------------- #
    def create_excel(self, file_name: str, folder_path: str = "") -> str:
        """Crea un archivo .xlsx vacío (drive item) de la forma más simple posible.

        Usa la operación estándar de OneDrive para crear un archivo enviando un POST a
        la colección `children` con `file: {}`. El archivo queda con tamaño 0 bytes.
        NOTA: Un .xlsx de 0 bytes no es un workbook válido para abrir en Excel; si
        necesitas un workbook funcional deberías subir contenido válido o partir de
        una plantilla. Este método cumple con la simplicidad solicitada (similar a copy).
        """
        if not file_name.lower().endswith('.xlsx'):
            raise ValueError("El nombre debe terminar en .xlsx")

        # Determinar endpoint children de la carpeta (o raíz si no se pasa folder_path)
        if folder_path:
            # Obtener ID de la carpeta para crear el archivo dentro
            folder_id = self.auth.get_folder_id(folder_path)
            url = f"{self.graph_url}/users/{self.user_id}/drive/items/{folder_id}/children"
        else:
            url = f"{self.graph_url}/users/{self.user_id}/drive/root/children"

        body = {
            "name": file_name,
            "file": {},
            "@microsoft.graph.conflictBehavior": "fail"  # o replace / rename según preferencia
        }
        headers = self.auth.get_headers()
        resp = requests.post(url, headers=headers, json=body)
        if resp.status_code not in (200, 201):
            raise Exception(f"❌ Error creando archivo: {resp.status_code} - {resp.text}")
        item_id = resp.json()["id"]
        print(f"✅ Archivo Excel creado (vacío): {file_name} (ID: {item_id})")
        return item_id

    def delete_excel(self, file_path: str) -> None:
        """Elimina un archivo Excel existente dado su ruta relativa."""
        file_id = self.auth.get_file_id(file_path)
        url = f"{self.graph_url}/users/{self.user_id}/drive/items/{file_id}"
        resp = self._request("DELETE", url, expected=(204,))
        if resp.status_code == 204:
            print(f"🗑️  Excel eliminado: {file_path}")

    def list_excels(self, folder_path: str = "") -> list[dict[str, Any]]:
        """Lista archivos .xlsx dentro de una carpeta.

        :return: Lista de objetos simplificados con name y id
        """
        path_segment = f"/{folder_path.strip('/')}" if folder_path else ""
        url = f"{self.graph_url}/users/{self.user_id}/drive/root:{path_segment}:/children"
        resp = self._request("GET", url, expected=200)
        items = resp.json().get('value', [])
        excels = [
            {"name": it["name"], "id": it["id"], "size": it.get("size"), "lastModifiedDateTime": it.get("lastModifiedDateTime")}
            for it in items if it.get('name', '').lower().endswith('.xlsx')
        ]
        return excels

    # -------------------------- Worksheet operations ---------------------- #
    def add_worksheet(self, file_path: str, worksheet_name: str) -> dict[str, Any]:
        """Agrega una nueva hoja a un workbook existente."""
        file_id = self.auth.get_file_id(file_path)
        url = f"{self.graph_url}/users/{self.user_id}/drive/items/{file_id}/workbook/worksheets/add"
        body = {"name": worksheet_name}
        resp = self._request("POST", url, expected=201, json=body)
        data = resp.json()
        print(f"➕ Hoja creada: {worksheet_name}")
        return data

    def delete_worksheet(self, file_path: str, worksheet_name: str) -> None:
        """Elimina una hoja por nombre.

        Nota: Excel siempre requiere al menos una hoja; Graph fallará si intentas borrar la última.
        """
        file_id = self.auth.get_file_id(file_path)
        # Listar hojas para encontrar ID
        list_url = f"{self.graph_url}/users/{self.user_id}/drive/items/{file_id}/workbook/worksheets"
        resp = self._request("GET", list_url, expected=200)
        worksheets = resp.json().get('value', [])
        target = next((w for w in worksheets if w.get('name') == worksheet_name), None)
        if not target:
            raise Exception(f"Hoja '{worksheet_name}' no encontrada")
        ws_id = target['id']
        del_url = f"{self.graph_url}/users/{self.user_id}/drive/items/{file_id}/workbook/worksheets/{ws_id}"
        self._request("DELETE", del_url, expected=204)
        print(f"🗑️  Hoja eliminada: {worksheet_name}")

    # -------------------------- Cell / Range operations ------------------- #
    def read_cells(self, file_path: str, worksheet_name: str, cells: list[str]) -> dict[str, Any]:
        """Lee múltiples celdas individuales usando una sola llamada batch.

        :return: dict direccion -> valor
        """
        file_id = self.auth.get_file_id(file_path)
        batch_url = f"{self.graph_url}/$batch"
        # Construir subpeticiones
        requests_payload = []
        for idx, cell in enumerate(cells, start=1):
            range_path = f"/users/{self.user_id}/drive/items/{file_id}/workbook/worksheets/{worksheet_name}/range(address='{cell}')"
            requests_payload.append({
                "id": str(idx),
                "method": "GET",
                "url": range_path
            })
        body = {"requests": requests_payload}
        headers = self.auth.get_headers().copy()
        resp = requests.post(batch_url, headers=headers, json=body)
        if resp.status_code != 200:
            raise Exception(f"❌ Error batch read: {resp.status_code} - {resp.text}")
        results = resp.json().get('responses', [])
        values: dict[str, Any] = {}
        for r in results:
            cid = int(r['id']) - 1
            cell_addr = cells[cid]
            if 200 <= r.get('status', 0) < 300:
                rng = r.get('body', {})
                vals = rng.get('values') or [[None]]
                values[cell_addr] = vals[0][0] if vals and vals[0] else None
            else:
                values[cell_addr] = None
        return values

    def write_range(self, file_path: str, worksheet_name: str, start_cell: str, values_2d: list[list[Any]]) -> None:
        """Escribe un rango rectangular (lista 2D) comenzando en start_cell.

        Calcula automáticamente la dirección final para el rango.
        """
        if not values_2d or not values_2d[0]:
            raise ValueError("values_2d no puede estar vacío")
        rows = len(values_2d)
        cols = len(values_2d[0])

        def col_to_num(col: str) -> int:
            n = 0
            for c in col:
                if not c.isalpha():
                    break
                n = n * 26 + (ord(c.upper()) - 64)
            return n

        def num_to_col(num: int) -> str:
            s = ""
            while num:
                num, r = divmod(num - 1, 26)
                s = chr(r + 65) + s
            return s

        import re
        m = re.match(r"([A-Za-z]+)(\d+)", start_cell)
        if not m:
            raise ValueError("start_cell inválido")
        start_col_letters, start_row_str = m.groups()
        start_row = int(start_row_str)
        start_col_num = col_to_num(start_col_letters)
        end_col_num = start_col_num + cols - 1
        end_row = start_row + rows - 1
        end_col_letters = num_to_col(end_col_num)
        range_address = f"{start_col_letters}{start_row}:{end_col_letters}{end_row}"

        file_id = self.auth.get_file_id(file_path)
        url = f"{self.graph_url}/users/{self.user_id}/drive/items/{file_id}/workbook/worksheets/{worksheet_name}/range(address='{range_address}')"
        body = {"values": values_2d}
        resp = self._request("PATCH", url, expected=200, json=body)
        if resp.status_code == 200:
            print(f"✍️  Rango {range_address} escrito correctamente")

    # -------------------------- Table operations -------------------------- #
    def create_table(self, file_path: str, worksheet_name: str, range_address: str, table_name: str, has_headers: bool = True) -> dict[str, Any]:
        """Crea una tabla a partir de un rango.

        :param range_address: Ej: "A1:D10"
        :param table_name: Nombre lógico de la tabla
        :param has_headers: Indica si la primera fila del rango son encabezados
        """
        file_id = self.auth.get_file_id(file_path)
        url = f"{self.graph_url}/users/{self.user_id}/drive/items/{file_id}/workbook/tables/add"
        body = {
            "address": f"{worksheet_name}!{range_address}",
            "hasHeaders": has_headers
        }
        resp = self._request("POST", url, expected=201, json=body)
        table = resp.json()
        # Renombrar si es necesario
        if table_name:
            t_id = table.get('id')
            if t_id:
                rename_url = f"{self.graph_url}/users/{self.user_id}/drive/items/{file_id}/workbook/tables/{t_id}"
                self._request("PATCH", rename_url, expected=200, json={"name": table_name})
                table['name'] = table_name
        print(f"📊 Tabla creada: {table.get('name')} en {range_address}")
        return table

    def add_table_rows(self, file_path: str, table_name: str, rows: list[list[Any]]) -> None:
        """Agrega filas a una tabla existente por su nombre."""
        if not rows:
            return
        file_id = self.auth.get_file_id(file_path)
        # Encontrar tabla por nombre
        list_url = f"{self.graph_url}/users/{self.user_id}/drive/items/{file_id}/workbook/tables"
        resp = self._request("GET", list_url, expected=200)
        tables = resp.json().get('value', [])
        target = next((t for t in tables if t.get('name') == table_name), None)
        if not target:
            raise Exception(f"Tabla '{table_name}' no encontrada")
        t_id = target['id']
        rows_url = f"{self.graph_url}/users/{self.user_id}/drive/items/{file_id}/workbook/tables/{t_id}/rows/add"
        body = {"values": rows}
        self._request("POST", rows_url, expected=200, json=body)
        print(f"{len(rows)} filas agregadas a la tabla {table_name}")
