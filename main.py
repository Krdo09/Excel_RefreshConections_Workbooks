from metodos_programa import *
import pandas as pd

if __name__ == '__main__':
    
    # Ingresar periodo de actualizacion
    tipo_semestre = input('Ingresar periodo de actualizacion: ')
    # Variable para control de flujo
    tipos_actualizacion = {
        'mensual': 'mensual',
        'bimestral': 'bimestral',
        'trimestral': 'trimestral',
        'cuatrimestral': 'cuatrimestral',
        'semestral': 'semestral',
        'anual': 'anual'
    }


    # Configuracion de logs
    logging.basicConfig(
        level=logging.INFO,
        format=' %(asctime)s - %(levelname)s - %(message)s',
        filename='HistoriaEjecucion.log',
        filemode='a')
    # Se eliminan los logs de tipo info para 'numexpr', mensajes default
    logging.getLogger("numexpr").setLevel(logging.ERROR)


    # Iniciar aplicacion excel
    excel_app = win32com.client.Dispatch('Excel.Application')
    # Desactivar interfaz de la aplicacion
    excel_app.Visible = False
    # Desactivar mensajes de confirmación de guardado
    excel_app.DisplayAlerts = False

    # Bloque principal de ejecucion
    try:
        # Mensaje de inicio de ejecucion principal
        logging.info(
            f'Inicio actualizacion archivos {tipos_actualizacion[tipo_semestre]}es'
        )
        # Condiciones de para determinar que archivos segun su fecha se van a actualizar
        if  tipos_actualizacion[tipo_semestre] in list(tipos_actualizacion.values()):
            # Abrir rutas del archivo
            with open(
                f'rutas_archivos_actualizar/rutas_adaptadas/{tipos_actualizacion[tipo_semestre]}.txt', mode='r', encoding='utf-8') as rutas_archivos:
                # Actualización de rutas
                for ruta in rutas_archivos:
                    print(ruta)

        else:
            # Mensaje, no cumplimiento condicón en bloque if
            logging.warning(
                f'El tipo de actualizacion "{tipos_actualizacion[tipo_semestre]}" no se encuentra en las opciones'
            )

    except Exception as error:
        # Mensaje de error
        logging.error(
            f'No fue posible actualizar los archivos, revisar excepcion:\n{error}')
        # Desactivar app excel
        excel_app.Quit()

    else:
        # Desactivar app excel al salir de bloque condicional
        logging.info(
            'El programa se ejecuto correctamente'
        )
        # Desactivar app excel
        excel_app.Quit()
