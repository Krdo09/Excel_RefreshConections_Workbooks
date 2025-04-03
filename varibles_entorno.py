import pandas as pd
# Ruta directorio de archivos para actulizar
directorio_rutas_actualizar = 'S:/Z. TIC/Julian Agudelo/Archivos en construcción/Python/Actualizacion Automatica Libros Excel/Archivos Actualizar.xlsx'

# Cargar información de archivos que se deben actualizar
nombres_directorios = ['Archivos Diarios', 'Archivos x Frecuencia']
tablas_rutas = {}
for nombre_hoja in nombres_directorios:
    excel_tabla = pd.read_excel(directorio_rutas_actualizar, sheet_name=nombre_hoja)
    tablas_rutas[nombre_hoja] = excel_tabla   

# Filtrar archivos para actualizar de 'Archivos_Frecuencia'
archivos_frecuencia_df = tablas_rutas[nombres_directorios[1]]
archivos_frecuencia_df = archivos_frecuencia_df[archivos_frecuencia_df['Prox. Actu.'] == datetime.now().strftime("%Y/%m/%d")]
# Sobre ecribir df
tablas_rutas['Archivos x Frecuencia'] = archivos_frecuencia_df

# Cargar parametros necesarios
frecuencias_df = pd.read_excel(directorio_rutas_actualizar, 
                            sheet_name='PdeC',
                            usecols='A:B',
                            nrows=8)

dias_df = pd.read_excel(directorio_rutas_actualizar, 
                            sheet_name='PdeC',
                            usecols='D:E',
                            nrows=8)
