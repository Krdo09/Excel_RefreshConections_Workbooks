from metodos_programa import *
import pandas as pd

if __name__ == '__main__':
    # Ingresar periodo de actualizacion
    tipo_actualizacion = input('Ingresar periodo de actualizacion: ')

    # Configuracion de logs
    logging.basicConfig(
        level=logging.INFO,
        format=' %(asctime)s - %(levelname)s - %(message)s',
        filename='HistoriaEjecucion.log',
        filemode='a')
    # Se eliminan los logs de tipo info para 'numexpr'
    logging.getLogger("numexpr").setLevel(logging.ERROR)

    # Iniciar aplicacion excel
    excel_app = win32com.client.Dispatch('Excel.Application')
    # Desactivar interfaz de la aplicacion
    excel_app.Visible = False
    # Desactivar mensajes de confirmación de guardado
    excel_app.DisplayAlerts = False

    try:
        # Mensaje de inicio de ejecucion principal
        logging.info(
            f'Inicio actualizacion archivos {tipo_actualizacion}es'
        )
        # Condiciones de actualizacion
        if tipo_actualizacion == 'mensual':
            pass

        elif tipo_actualizacion == 'bimestral':
            pass

        elif tipo_actualizacion == 'trimestral':
            pass

        elif tipo_actualizacion == 'cuatrimestral':
            pass

        elif tipo_actualizacion == 'semestral':
            pass

        elif tipo_actualizacion == 'anual':
            pass

        else:
            # Mensaje, no cumplimiento condicón en bloque if
            logging.warning(
                f'El tipo de actualizacion "{tipo_actualizacion}" no se encuentra en las opciones'
            )

    except Exception as error:
        # Mensaje de error
        logging.error(
            f'No fue posible actualizar los archivos, revisar excepción:\n{error}')
        # Desactivar app excel
        excel_app.Quit()

    else:
        # Desactivar app excel al salir de bloque condicional
        logging.info(
            'El programa se ejecuto correctamente'
        )
        # Desactivar app excel
        excel_app.Quit()
