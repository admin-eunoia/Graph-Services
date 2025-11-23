# Excel Live Writer - Documentación

Sistema de escritura en Excel EN VIVO usando Microsoft Graph API. Permite editar archivos Excel sin descargarlos, **incluso si están abiertos por otros usuarios**.

---

## 📋 Tabla de Contenidos

1. [Introducción](#introducción)
2. [Requisitos](#requisitos)
3. [Funciones Disponibles](#funciones-disponibles)
4. [Ejemplos de Uso](#ejemplos-de-uso)
5. [Casos de Uso Comunes](#casos-de-uso-comunes)

---

## 🎯 Introducción

`excel_live_writer.py` proporciona funciones para editar archivos Excel almacenados en OneDrive de manera directa usando la API de Microsoft Graph, sin necesidad de descargar/subir archivos.

### Ventajas:

- ✅ **Edición en tiempo real** - No requiere descargar/subir archivos
- ✅ **Sin bloqueos** - Funciona incluso si alguien tiene el archivo abierto
- ✅ **Basado en marcadores** - Encuentra secciones automáticamente por texto
- ✅ **Manejo de merges** - Aplica merges de celdas automáticamente
- ✅ **Múltiples secciones** - Procesa varias secciones en una sola llamada

---

## 📦 Requisitos

```python
from Services.excel_live_writer import (
    buscar_marcador_live,
    llenar_seccion_live,
    llenar_tabla_live,
    insertar_filas_live,
    procesar_excel_live
)
from Services.graph_services import GraphServices
from Auth.Microsoft_Graph_Auth import MicrosoftGraphAuthenticator
```

### Autenticación:

```python
auth = MicrosoftGraphAuthenticator(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
token = auth.get_access_token()
client = GraphServices(token)
```

---

## 🔧 Funciones Disponibles

### 1. `buscar_marcador_live()`

Busca un marcador (texto) en el Excel y devuelve su posición.

#### Parámetros:

| Parámetro        | Tipo            | Requerido | Descripción                                               |
| ---------------- | --------------- | --------- | --------------------------------------------------------- |
| `client`         | `GraphServices` | ✅        | Cliente autenticado de GraphServices                      |
| `file_path`      | `str`           | ✅        | Ruta del archivo en OneDrive (ej: "Carpeta/archivo.xlsx") |
| `marker`         | `str`           | ✅        | Texto a buscar en el Excel                                |
| `target_user_id` | `str`           | ❌        | Email del usuario (si no se usa `drive_id`)               |
| `drive_id`       | `str`           | ❌        | ID del drive (alternativa a `target_user_id`)             |
| `sheet_name`     | `str`           | ❌        | Nombre de la hoja (None = primera hoja)                   |

#### Retorna:

`Tuple[Optional[int], Optional[int]]` - `(fila, columna)` donde se encontró el marcador, o `(None, None)` si no se encuentra.

#### Ejemplo:

```python
fila, columna = buscar_marcador_live(
    client=client,
    file_path="Proyectos/Contrato.xlsx",
    marker="DATOS DEL CLIENTE",
    target_user_id="usuario@empresa.com"
)

if fila:
    print(f"Marcador encontrado en fila {fila}, columna {columna}")
```

---

### 2. `llenar_seccion_live()`

Llena una sección simple (clave-valor) en el Excel. Ideal para formularios donde cada campo está en una fila diferente.

#### Parámetros:

| Parámetro        | Tipo             | Requerido | Descripción                                     |
| ---------------- | ---------------- | --------- | ----------------------------------------------- |
| `client`         | `GraphServices`  | ✅        | Cliente autenticado de GraphServices            |
| `file_path`      | `str`            | ✅        | Ruta del archivo en OneDrive                    |
| `marker`         | `str`            | ✅        | Marcador que identifica la sección              |
| `datos`          | `Dict[str, Any]` | ✅        | Diccionario con los datos `{campo: valor}`      |
| `columnas`       | `Dict[str, int]` | ✅        | Mapeo de campo a offset de columna `{campo: 0}` |
| `target_user_id` | `str`            | ❌        | Email del usuario                               |
| `drive_id`       | `str`            | ❌        | ID del drive                                    |
| `sheet_name`     | `str`            | ❌        | Nombre de la hoja                               |

#### Comportamiento:

- Busca el `marker` en el Excel
- Escribe los datos en la **fila siguiente** al marcador
- Cada campo se escribe en la columna especificada por el offset

#### Ejemplo:

```python
llenar_seccion_live(
    client=client,
    file_path="Proyectos/Contrato.xlsx",
    marker="DATOS DEL CLIENTE",
    datos={
        "nombre": "ACME Corporation",
        "rfc": "ACM123456ABC",
        "telefono": "555-1234"
    },
    columnas={
        "nombre": 0,    # Columna A (relativa al marcador)
        "rfc": 1,       # Columna B
        "telefono": 2   # Columna C
    },
    target_user_id="usuario@empresa.com"
)
```

**Resultado en Excel:**

```
| DATOS DEL CLIENTE |           |              |
|-------------------|-----------|--------------|
| ACME Corporation  | ACM123456ABC | 555-1234  |
```

---

### 3. `llenar_tabla_live()`

Llena una tabla con múltiples filas de datos. Ideal para listas, tablas de seguimiento, etc.

#### Parámetros:

| Parámetro          | Tipo                   | Requerido | Descripción                                      |
| ------------------ | ---------------------- | --------- | ------------------------------------------------ |
| `client`           | `GraphServices`        | ✅        | Cliente autenticado de GraphServices             |
| `file_path`        | `str`                  | ✅        | Ruta del archivo en OneDrive                     |
| `marker`           | `str`                  | ✅        | Marcador que identifica la tabla                 |
| `datos`            | `List[Dict[str, Any]]` | ✅        | Lista de diccionarios con los datos de cada fila |
| `columnas`         | `Dict[str, int]`       | ✅        | Mapeo de campo a offset de columna               |
| `target_user_id`   | `str`                  | ❌        | Email del usuario                                |
| `drive_id`         | `str`                  | ❌        | ID del drive                                     |
| `sheet_name`       | `str`                  | ❌        | Nombre de la hoja                                |
| `merges_a_aplicar` | `List[str]`            | ❌        | Lista de rangos a mergear (ej: `["A:C"]`)        |

#### Comportamiento:

- Busca el `marker` en el Excel
- Escribe los datos comenzando en **marcador + 2 filas** (asume una fila de encabezado)
- Escribe todas las filas en una sola operación
- Opcionalmente aplica merges a cada fila

#### Ejemplo:

```python
llenar_tabla_live(
    client=client,
    file_path="Proyectos/Contrato.xlsx",
    marker="SEGUIMIENTO",
    datos=[
        {"fecha": "2025-01-15", "medio": "Email", "comentarios": "Contacto inicial"},
        {"fecha": "2025-01-20", "medio": "Teléfono", "comentarios": "Seguimiento"},
        {"fecha": "2025-01-25", "medio": "Reunión", "comentarios": "Cierre de contrato"}
    ],
    columnas={
        "fecha": 0,
        "medio": 1,
        "comentarios": 2
    },
    target_user_id="usuario@empresa.com",
    merges_a_aplicar=["A:C"]  # Opcional: merge columnas A-C en cada fila
)
```

**Resultado en Excel:**

```
| SEGUIMIENTO |                    |                        |
|-------------|--------------------|-----------------------|
| FECHA       | MEDIO              | COMENTARIOS           |
| 2025-01-15  | Email              | Contacto inicial      |
| 2025-01-20  | Teléfono           | Seguimiento           |
| 2025-01-25  | Reunión            | Cierre de contrato    |
```

---

### 4. `insertar_filas_live()`

Inserta nuevas filas en una posición específica, **moviendo las filas existentes hacia abajo**. Ideal para agregar datos a tablas existentes sin sobrescribir.

#### Parámetros:

| Parámetro          | Tipo                   | Requerido | Descripción                                      |
| ------------------ | ---------------------- | --------- | ------------------------------------------------ |
| `client`           | `GraphServices`        | ✅        | Cliente autenticado de GraphServices             |
| `file_path`        | `str`                  | ✅        | Ruta del archivo en OneDrive                     |
| `fila_inicio`      | `int`                  | ✅        | Número de fila donde insertar (1-indexed)        |
| `datos`            | `List[Dict[str, Any]]` | ✅        | Lista de diccionarios con los datos              |
| `columnas`         | `Dict[str, int]`       | ✅        | Mapeo de campo a offset de columna               |
| `target_user_id`   | `str`                  | ❌        | Email del usuario                                |
| `drive_id`         | `str`                  | ❌        | ID del drive                                     |
| `sheet_name`       | `str`                  | ❌        | Nombre de la hoja                                |
| `columna_inicio`   | `int`                  | ❌        | Columna inicial (default=1 para A)               |
| `merges_a_aplicar` | `List[str]`            | ❌        | Lista de rangos a mergear (ej: `["A:C", "D:F"]`) |

#### Comportamiento:

1. **Inserta filas vacías** en la posición especificada
2. Las filas existentes se mueven hacia abajo
3. **Llena las nuevas filas** con los datos proporcionados
4. **Aplica merges** si se especificaron

#### Ejemplo:

```python
insertar_filas_live(
    client=client,
    file_path="Proyectos/Contrato.xlsx",
    fila_inicio=25,  # Insertar en la fila 25
    datos=[
        {"fecha": "2025-12-01", "medio": "WhatsApp", "comentarios": "Nueva entrada"},
        {"fecha": "2025-12-02", "medio": "Video", "comentarios": "Segunda entrada"},
        {"fecha": "2025-12-03", "medio": "Presencial", "comentarios": "Tercera entrada"}
    ],
    columnas={
        "fecha": 0,
        "medio": 1,
        "comentarios": 2
    },
    target_user_id="usuario@empresa.com",
    columna_inicio=1,  # Comenzar en columna A
    merges_a_aplicar=["A:C"]  # Mergear A-C en cada fila
)
```

**Resultado:**

- Se insertan 3 filas nuevas en la fila 25
- Las filas 25, 26, 27... existentes se mueven a 28, 29, 30...
- Las nuevas filas se llenan con los datos
- Las columnas A-C se mergean en cada fila nueva

---

### 5. `procesar_excel_live()`

Función todo-en-uno para procesar **múltiples secciones** en un solo archivo Excel. Ideal para llenar formularios completos o reportes.

#### Parámetros:

| Parámetro        | Tipo              | Requerido | Descripción                                 |
| ---------------- | ----------------- | --------- | ------------------------------------------- |
| `client`         | `GraphServices`   | ✅        | Cliente autenticado de GraphServices        |
| `file_path`      | `str`             | ✅        | Ruta del archivo en OneDrive                |
| `secciones`      | `Dict[str, Any]`  | ✅        | Datos por sección `{nombre_seccion: datos}` |
| `configuracion`  | `Dict[str, Dict]` | ✅        | Configuración de cada sección               |
| `target_user_id` | `str`             | ❌        | Email del usuario                           |
| `drive_id`       | `str`             | ❌        | ID del drive                                |

#### Estructura de `configuracion`:

Cada sección debe tener:

- `marker`: Texto marcador en el Excel
- `es_tabla`: `True` para tablas, `False` para secciones simples
- `columnas`: Mapeo de campo a offset de columna
- `sheet_name` (opcional): Nombre de la hoja

#### Ejemplo Completo:

```python
procesar_excel_live(
    client=client,
    file_path="Proyectos/Contrato.xlsx",
    secciones={
        "cliente": {
            "nombre": "ACME Corporation",
            "rfc": "ACM123456ABC",
            "telefono": "555-1234"
        },
        "evento": {
            "fecha": "2025-12-15",
            "lugar": "Centro de Convenciones",
            "pax": 200
        },
        "seguimiento": [
            {"fecha": "2025-01-15", "medio": "Email", "comentarios": "Inicial"},
            {"fecha": "2025-01-20", "medio": "Teléfono", "comentarios": "Seguimiento"}
        ]
    },
    configuracion={
        "cliente": {
            "marker": "DATOS DEL CLIENTE",
            "es_tabla": False,
            "columnas": {"nombre": 0, "rfc": 1, "telefono": 2}
        },
        "evento": {
            "marker": "DATOS DEL EVENTO",
            "es_tabla": False,
            "columnas": {"fecha": 0, "lugar": 1, "pax": 2}
        },
        "seguimiento": {
            "marker": "SEGUIMIENTO",
            "es_tabla": True,
            "columnas": {"fecha": 0, "medio": 1, "comentarios": 2}
        }
    },
    target_user_id="usuario@empresa.com"
)
```

---

## 💡 Casos de Uso Comunes

### Caso 1: Llenar Formulario Simple

```python
# Llenar sección de cliente
llenar_seccion_live(
    client=client,
    file_path="Contratos/Nuevo_Contrato.xlsx",
    marker="INFORMACIÓN DEL CLIENTE",
    datos={
        "nombre_completo": "Juan Pérez García",
        "email": "juan@ejemplo.com",
        "empresa": "Tech Solutions SA"
    },
    columnas={"nombre_completo": 0, "email": 1, "empresa": 2},
    target_user_id="admin@empresa.com"
)
```

### Caso 2: Agregar Registros a Tabla de Seguimiento

```python
# Insertar nuevas entradas en tabla existente
insertar_filas_live(
    client=client,
    file_path="Proyectos/Seguimiento_2025.xlsx",
    fila_inicio=30,  # Después de los registros existentes
    datos=[
        {"fecha": "2025-11-23", "actividad": "Llamada", "notas": "Cliente satisfecho"}
    ],
    columnas={"fecha": 0, "actividad": 1, "notas": 2},
    target_user_id="admin@empresa.com",
    merges_a_aplicar=["A:C"]
)
```

### Caso 3: Llenar Reporte Completo

```python
# Procesar múltiples secciones de un reporte
procesar_excel_live(
    client=client,
    file_path="Reportes/Reporte_Mensual.xlsx",
    secciones={
        "encabezado": {
            "mes": "Noviembre",
            "año": 2025,
            "departamento": "Ventas"
        },
        "metricas": [
            {"kpi": "Ventas Totales", "valor": 150000, "meta": 120000},
            {"kpi": "Nuevos Clientes", "valor": 45, "meta": 40},
            {"kpi": "Satisfacción", "valor": 95, "meta": 90}
        ]
    },
    configuracion={
        "encabezado": {
            "marker": "INFORMACIÓN DEL REPORTE",
            "es_tabla": False,
            "columnas": {"mes": 0, "año": 1, "departamento": 2}
        },
        "metricas": {
            "marker": "MÉTRICAS",
            "es_tabla": True,
            "columnas": {"kpi": 0, "valor": 1, "meta": 2}
        }
    },
    target_user_id="admin@empresa.com"
)
```

---

## ⚙️ Configuración de Merges

Los merges se especifican usando notación de columnas:

```python
merges_a_aplicar=[
    "A:C",    # Merge columnas A, B, C
    "D:F",    # Merge columnas D, E, F
    "G:J"     # Merge columnas G, H, I, J
]
```

Esto es útil para:

- Celdas de comentarios largos
- Campos de descripción
- Mantener el formato visual de templates

---

## 🔍 Notas Importantes

### Índices de Fila y Columna:

- **Filas**: 1-indexed (la fila 1 es la primera)
- **Columnas**: Los offsets son 0-indexed relativos al marcador

### Estructura Esperada:

```
| MARCADOR      |              |                |
|---------------|--------------|----------------|
| dato1         | dato2        | dato3          |  <- Se escribe aquí (marcador + 1)
```

Para tablas:

```
| MARCADOR      |              |                |
|---------------|--------------|----------------|
| Header 1      | Header 2     | Header 3       |  <- marcador + 1
| dato fila 1   | dato fila 1  | dato fila 1    |  <- Se escribe aquí (marcador + 2)
| dato fila 2   | dato fila 2  | dato fila 2    |
```

### Manejo de Errores:

- Si no se encuentra el marcador, se lanza `ValueError`
- Si hay errores de API, se imprime el error pero no detiene el proceso
- Los merges son opcionales y no críticos (fallan silenciosamente si hay problemas)

---

## 📝 Testing

Ejecuta los tests para verificar que todo funciona:

```bash
python test_live_writer.py
```

Tests disponibles:

1. ✅ Editar una celda
2. ✅ Editar sección completa
3. ✅ Llenar tabla
4. ✅ Procesar todo
5. ✅ Insertar filas con merges

---

## 🚀 Ventajas vs Descarga/Subida

| Característica  | Live Writer  | Download/Upload       |
| --------------- | ------------ | --------------------- |
| Velocidad       | ⚡ Rápido    | 🐌 Lento              |
| Archivo abierto | ✅ Funciona  | ❌ Falla (423 Locked) |
| Ancho de banda  | 📉 Mínimo    | 📈 Alto               |
| Complejidad     | 🎯 Simple    | 🔧 Complejo           |
| Escalabilidad   | ✅ Excelente | ❌ Limitada           |

---

## 📞 Soporte

Para reportar problemas o sugerencias, contacta al equipo de desarrollo.
