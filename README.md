# Graph API – Autenticación y Servicios de Excel

Este proyecto expone dos clases principales:

- `MicrosoftGraphAuthenticator` en `Microsoft_Graph_Auth.py`: gestiona la autenticación (MSAL) y ayuda a resolver IDs de archivos/carpetas.
- `GraphServices` en `graph_services.py`: ofrece servicios de `copy_excel` y `fill_excel` sobre Microsoft Graph.

## Requisitos

- Python 3.9+
- Tener una app registrada en Azure AD con permisos de Microsoft Graph adecuados.
- Variables de entorno configuradas.

## Instalación

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Copia `.env.example` a `.env` y completa tus credenciales:

```bash
cp .env.example .env
```

Edita `.env` con tus valores reales.

## Uso rápido

Ejecuta el ejemplo de uso:

```bash
python example_usage.py
```

O bien, usa las clases directamente:

```python
from graph_services import GraphServices

services = GraphServices()  # usa el user_id por defecto del autenticador

# 1) Copiar un Excel dentro de una carpeta
services.copy_excel(
    original_name="Plantilla.xlsx",
    copy_name="MiCopia.xlsx",
    folder_path="Documentos/Proyectos",
)

# 2) Escribir datos en celdas de un Excel existente
services.fill_excel(
    file_path="Documentos/Proyectos/MiCopia.xlsx",
    worksheet_name="Hoja1",
    data={
        "A1": "Juan",
        "A2": 30,
        "A3": "M",
        "A4": "Ingeniero",
    },
)

# 3) Crear un Excel vacío en una carpeta
services.create_excel("NuevoArchivo.xlsx", folder_path="Documentos/Proyectos")

# 4) Listar Excels en una carpeta
print(services.list_excels("Documentos/Proyectos"))

# 5) Añadir y eliminar hojas
services.add_worksheet("Documentos/Proyectos/NuevoArchivo.xlsx", "Datos")
services.delete_worksheet("Documentos/Proyectos/NuevoArchivo.xlsx", "Datos")

# 6) Escribir un rango (matriz 2D) comenzando en A1
services.write_range(
    file_path="Documentos/Proyectos/MiCopia.xlsx",
    worksheet_name="Hoja1",
    start_cell="B5",
    values_2d=[["Producto", "Precio"], ["A", 10], ["B", 20]],
)

# 7) Leer celdas múltiples con batch
vals = services.read_cells(
    file_path="Documentos/Proyectos/MiCopia.xlsx",
    worksheet_name="Hoja1",
    cells=["A1", "B2", "B5"],
)
print(vals)

# 8) Crear tabla y agregar filas
services.create_table(
    file_path="Documentos/Proyectos/MiCopia.xlsx",
    worksheet_name="Hoja1",
    range_address="B5:C7",  # Debe cubrir encabezados + filas iniciales
    table_name="TablaProductos",
    has_headers=True,
)
services.add_table_rows(
    file_path="Documentos/Proyectos/MiCopia.xlsx",
    table_name="TablaProductos",
    rows=[["C", 30], ["D", 40]],
)

# 9) Eliminar un Excel
services.delete_excel("Documentos/Proyectos/NuevoArchivo.xlsx")
```

## Notas

- La operación de copia (`/copy`) devuelve `202 Accepted`; Microsoft Graph realiza la copia en segundo plano. Si necesitas confirmar finalización, deberás consultar el estado del trabajo usando el header `Location` que retorna ese endpoint.
- Asegúrate de que el `user_id` tenga acceso al archivo/carpeta objetivo en OneDrive.
- La creación de un Excel vacío genera un archivo mínimo válido (ZIP estructurado) para poder subirlo directamente vía `PUT /content`.
- Operaciones batch para lectura de múltiples celdas reducen el número de llamadas individuales.
- No puedes eliminar la última hoja de un workbook; Graph devolverá error.
- Para tablas: el `range_address` debe incluir encabezados si `has_headers=True`.

## Variables de entorno esperadas

- `MICROSOFT_CLIENT_ID`
- `MICROSOFT_CLIENT_SECRET`
- `MICROSOFT_TENANT_ID`
