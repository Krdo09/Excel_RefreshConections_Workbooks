from varibles_entorno import *
from metodos_programa import *

if __name__ == '__main__':
    
    # Configuracion de logs
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        filename='HistoriaEjecucion.log',
        filemode='a'
        )
    # Se eliminan los logs de tipo info para 'numexpr', mensajes default
    logging.getLogger("numexpr").setLevel(logging.ERROR)
    
    # Bloque principal de ejecucion
    try:
        # Crear trazabilidad de archivo de rutas
        trazabilidad_archivo(ruta_libro=Path(directorio_rutas_actualizar), acronimo='00TIC')
        
        # Cargar información de archivos que se deben actualizar
        nombres_hojas = ['Archivos Diarios', 'Archivos x Frecuencia']
        tablas_rutas = {}
        for nombre_hoja in nombres_hojas:
            excel_tabla = pd.read_excel(directorio_rutas_actualizar, sheet_name=nombre_hoja)
            tablas_rutas[nombre_hoja] = excel_tabla   

        # Filtrar archivos para actualizar de 'Archivos_Frecuencia'
        archivos_frecuencia_df = tablas_rutas[nombres_hojas[1]]
        archivos_frecuencia_df = archivos_frecuencia_df[archivos_frecuencia_df['Prox. Actu.'] == datetime.now().strftime("%Y/%m/%d")]
        # Sobre ecribir df
        tablas_rutas['Archivos x Frecuencia'] = archivos_frecuencia_df


        # Mensaje de inicio de ejecucion principal
        logging.info(
            f'Inicio actualizacion para tipo archivo: "{archivos_nombre}"'
            )
        
        # Condiciones de para determinar que archivos segun su llave se van a actualizar
        if archivos_nombre in set(tipos_WorkBooks.keys()):
            # Abrir rutas del archivo
            with open(
                tipos_WorkBooks[archivos_nombre], 
                mode='r', 
                encoding='utf-8'
                ) as rutas_archivos:

                # Bloque principal para la actualización de los libros  
                for ruta_txt in rutas_archivos:
                    # Limpiar ruta de carcteres especiales
                    ruta_txt = ruta_txt.replace('\n', '')
                    # Convertir en objeto Path (administrador de rutas)
                    obj_path = Path(ruta_txt)
                    # Crear trazabilidad de archivo
                    trazabilidad_archivo(obj_path, acronimo="00TIC")
                    # Actualizar libros
                    actualizar_libros(obj_path)

        else:
            # Mensaje, no cumplimiento condicón en bloque if
            logging.warning(
                f'El tipo de actualizacion "{archivos_nombre}" no se encuentra en las opciones'
                )

    except Exception as error:
        # Mensaje de error
        logging.error(
            f'No fue posible actualizar los archivos, revisar excepcion:\n{error}'
            )

    else:
        # Mensaje de ejecución exitosa
        logging.info(
            'El programa se ejecuto correctamente'
            )

