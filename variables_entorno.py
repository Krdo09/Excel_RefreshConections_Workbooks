from metodos_programa import *

# Ruta directorio de archivos para actulizar
directorio_rutas_actualizar = 'S:/Z. TIC/Julian Agudelo/Archivos en construcción/Python/Actualizacion Automatica Libros Excel/Archivos Actualizar.xlsx'

# Cargar información de archivos que se deben actualizar
archivos_diarios = pd.read_excel(directorio_rutas_actualizar, sheet_name='Archivos Diarios')

# Ordenar archivos diarios en orden de prioridad
archivos_diarios.sort_values(by='Priorida Actu.', ascending=True, inplace=True)

print(archivos_diarios)