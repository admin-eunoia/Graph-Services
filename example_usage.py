from Services.graph_services import GraphServices

services = GraphServices()
"""Ejemplos de uso ampliados de GraphServices.

ADVERTENCIA: Estos llamados realizan operaciones reales sobre OneDrive.
Descomenta según necesites probar cada funcionalidad.
"""

# 1. Copiar Excel existente
# services.copy_excel("Cliente; dan.xlsx", "MiCopia.xlsx", "Prueba WAMAN")

# 2. Escribir celdas individuales
services.fill_excel("Prueba WAMAN/MiCopia.xlsx", "Hoja1", {"A1": "Rodrigo", "B2": 28})

# 3. Crear Excel vacío
# services.create_excel("NuevoDemo.xlsx", folder_path="Prueba WAMAN")

# 4. Listar Excels en carpeta
# print(services.list_excels("Prueba WAMAN"))

# 5. Añadir hoja
# services.add_worksheet("Prueba WAMAN/NuevoDemo.xlsx", "Datos")

# 6. Escribir rango 2D
# services.write_range("Prueba WAMAN/MiCopia.xlsx", "Hoja1", "C3", [["Col1", "Col2"], [1, 2], [3, 4]])

# 7. Leer varias celdas
# print(services.read_cells("Prueba WAMAN/MiCopia.xlsx", "Hoja1", ["A1", "B2", "C3"]))

# 8. Crear tabla (suponiendo datos existentes en C3:D5)
# services.create_table("Prueba WAMAN/MiCopia.xlsx", "Hoja1", "C3:D5", "TablaEjemplo", has_headers=True)

# 9. Agregar filas a la tabla
# services.add_table_rows("Prueba WAMAN/MiCopia.xlsx", "TablaEjemplo", [["Nueva", 99]])

# 10. Eliminar hoja
# services.delete_worksheet("Prueba WAMAN/NuevoDemo.xlsx", "Datos")

# 11. Eliminar Excel
# services.delete_excel("Prueba WAMAN/NuevoDemo.xlsx")

