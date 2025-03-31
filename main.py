from rutas_entorno import *
from metodos_programa import *

if __name__ == '__main__':
    
    # Obtener argumento ingresado en ejecutable .bat
    if len(sys.argv) > 1:
        archivos_nombre = sys.argv[1]

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
                    trazabilidad_archivo(obj_path, acronimo="TIC_Automatico")
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

