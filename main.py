from metodos_programa import *
import pandas as pd

if __name__ == '__main__':
    # Ingresar periodo de actualizacion
    tipo_actualizacion = input('Ingresar periodo de actualizacion') 

    # Iniciar aplicacion excel
    excel_app = win32com.client.Dispatch('Excel_Application')
    # Desactivar visualización grafica de la aplicacion
    excel_app.Visible = False
    # Desactivar mensajes de confirmación de guardado
    excel_app.DisplayAlerts = False

    try:
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
            pass

        # Desactivar app excel al final de la ejecucion
        excel_app.Quit()

    except Exception as error:
        # Desactivar app excel
        excel_app.Quit()
