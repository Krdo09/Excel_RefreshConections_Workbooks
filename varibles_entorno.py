import pandas as pd
# Ruta directorio de archivos para actulizar
directorio_rutas_actualizar = 'S:/Z. TIC/Julian Agudelo/Archivos en construcción/Python/Actualizacion Automatica Libros Excel/Archivos Actualizar.xlsx'

# Cargar parametros necesarios
frecuencias_df = pd.read_excel(directorio_rutas_actualizar, 
                            sheet_name='PdeC',
                            usecols='A:B',
                            nrows=8)

dias_df = pd.read_excel(directorio_rutas_actualizar, 
                            sheet_name='PdeC',
                            usecols='D:E',
                            nrows=8)
