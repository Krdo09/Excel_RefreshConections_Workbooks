from datetime import datetime
from pathlib import Path
import logging
import shutil
import sys
import time 
import win32com.client as win32com


def trazabilidad_archivo(ruta_libro: Path, acronimo: str) -> None:
    """
    Copia y pega un archivo seleccionado en la ruta indicada,
    agregando la fecha en que se realizó dicha copia y el acronimo
    de persona que ejecutó el script; generando una trazabilidad de los 
    cambios realizados en dicho archivo.

    El formato de nombre que se le asigna a la copia es:
    "nombre_original_archivo - aaaa-mm-dd - acronimo_persona"

    Parameters:
    ruta_libro: Ruta del archivo .xlsx que se desea generar trazabilidad | type str
    acronimo: Acronimo del encargada de ejecutar el script | type str

    Returns:
    Sin retorno de datos | type None

    Raises:
    FileNotFoundError: No se encontro la ruta, el archivo o la ruta es incorrecta
    """
    try:
        # Se recopila fecha para la trazabilidad del archivo
        fecha_actual = datetime.now().strftime("%Y-%m-%d")

        # Nuevo nombre para la copia del archivo
        trazabilidad_archivo = (
            f"{ruta_libro.stem} - {fecha_actual} - {acronimo}{ruta_libro.suffix}"
            )

        # Se obtiene la ruta padre donde está alojado el archivo
        ruta_padre = ruta_libro.parent

        # Nueva ruta para el archivo copiado
        ruta_archivo_copia = ruta_padre / trazabilidad_archivo

        # Se genera la trazabilidad del archivo original
        shutil.copy(ruta_libro.as_posix(), ruta_archivo_copia)

        # Mensaje para creación de trazabilidad correcta
        logging.info(
            f'Se creo exitosamente trazabilidad para el archivo "{ruta_libro.stem}"'
            )

    except Exception as error:
        # Mensaje error
        logging.error(
            f'Verificar la estructura o ruta del archivo "{ruta_libro.stem}", excepcion:\n{error}'
            )


def actualizar_libros(ruta_libro: Path) -> None:
    """
    Con la ruta la administrada se abre el respectivo archivo de excel
    y se actualizan todas las consultas asociadas a dicho documento.

    Parameters:
    excel_app: Clase de win32com con aplicacion Excel ejecutada | type CDispatch
    ruta_libro: Ruta del archivo .xlsx que se desea transformar | type str

    Returns:
    None

    Raises:
    FileNotFoundError: No se encontro la ruta o es incorrecta
    """
    try:
        print(ruta_libro.stem)
        # Notificar inicio de actualización del archivo
        logging.info(
            f'Inicio actualizacion del archivo "{ruta_libro.stem}"'
        )
        # Abrir aplicación Excel
        excel_ejecutado = win32com.Dispatch('Excel.Application')
        # Activar o desactivar pestaña de excel
        excel_ejecutado.Visible = False
        # Activar o desactivar pestañas emergentes
        excel_ejecutado.DisplayAlerts = False 

        # Abrir el archivo de excel
        libro_actualizar = excel_ejecutado.Workbooks.Open(ruta_libro.as_posix())

        # Ejecutar comando para actualizar consultas
        consulta = 1
        for query in libro_actualizar.Queries:
            query.Refresh()
            print(consulta)
            # Dejar reposar el algoritmo por 10 segundos
            time.sleep(10)
            consulta += 1

        # Esperar a que se terminen de ejecutar las rutinas antes de cerrar archivo
        excel_ejecutado.CalculateUntilAsyncQueriesDone()
 
        # Guardar archivo actualizado
        libro_actualizar.Save()

        # Cerrar archivo
        libro_actualizar.Close()

        # Cerrar aplicación Excel
        excel_ejecutado.Quit()

        # Mensaje para actualización correcta
        logging.info(
            f'Actualizacion completa para el archivo: "{ruta_libro.stem}"'
            )

    except Exception as error:
        # Mensaje error      
        logging.error(
            f'Verificar el archivo "{ruta_libro.stem}", excepcion:\n{error}'
            )   

        # Cerrar archivo sin guardar
        libro_actualizar.Close()

        # Cerrar aplicación
        excel_ejecutado.Quit()
