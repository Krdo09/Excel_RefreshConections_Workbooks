from metodos_programa import *

if __name__ == '__main__':
    
    # Ingresar periodo de actualizacion
    tipo_semestre = input('Ingresar periodo de actualizacion: ')
    # Variable para control de flujo programa
    tipos_actualizacion = {
        'mensual': {'mensual': 'TIC_UPDATE_MENSUAL'},
        'bimestral': {'bimestral': 'TIC_UPDATE_BIMESTRAL'},
        'trimestral': {'trimestral': 'TIC_UPDATE_TRIMESTRAL'},
        'cuatrimestral': {'cuatrimestral':'TIC_UPDATE_CUATRIMESTRAL'},
        'semestral': {'semestral':'TIC_UPDATE_SEMESTRAL'},
        'anual': {'anual': 'TIC_UPDATE_ANUAL'},
        'pruebas': {'pruebas': 'TIC_UPDATE'}
    }

    # Configuracion de logs
    logging.basicConfig(
        level=logging.INFO,
        format=' %(asctime)s - %(levelname)s - %(message)s',
        filename='HistoriaEjecucion.log',
        filemode='a')
    # Se eliminan los logs de tipo info para 'numexpr', mensajes default
    logging.getLogger("numexpr").setLevel(logging.ERROR)

    # Bloque principal de ejecucion
    try:
        # Mensaje de inicio de ejecucion principal
        logging.info(
            f'Inicio actualizacion para tipo archivo: "{tipo_semestre}"'
            )
        
        # Condiciones de para determinar que archivos segun su fecha se van a actualizar
        if  list(tipos_actualizacion[tipo_semestre].keys())[0] in list(tipos_actualizacion[tipo_semestre].keys()):
            # Abrir rutas del archivo
            with open(
                f'rutas_archivos_actualizar/rutas_adaptadas/{tipo_semestre}.txt', mode='r', encoding='utf-8'
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
                f'El tipo de actualizacion "{tipo_semestre}" no se encuentra en las opciones'
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

