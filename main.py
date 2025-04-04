from variables_entorno import *

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
        
        # Actulizar rutas en los directorios
        indice = 0
        for directorio in tablas_rutas.values:
            # Mensaje inicio actualización tipos de archivos
            logging.info(
                f'Inicio actualizacion para: "{nombres_directorios[indice]}"'
            )
            
            # Recorrer cada registro (archivo) en el directorio
            for _, fila in directorio.itertuples():
                # Mensaje inicio actualizacion archivo
                logging.info(
                    f'Inicio actualizacion para tipo archivo: "{fila[0]}"'
                    )
                
                # Cargar ruta del archivo
                ruta_archivo = Path(fila[1])
                # Crear trazabilidad del archivo
                trazabilidad_archivo(ruta_archivo, acronimo='00TIC')
                # Actualizar archivos
                actualizar_libros(ruta_archivo)
                
                #Moficar con fechas actualizadas
                nuevas_fechas = proxima_fecha_actualizacion(
                    frecuencias=frecuencias_df,
                    dias=dias_df,
                    fecha_actualizacion=fila[4],
                    frecuencia=fila[3]
                )

            indice += 1
            
            # Mensaje finzalizacion actualizacion tipos de archivos
            logging.info(
                f"""Actualización Finalizada para: {nombres_directorios[indice]}\n
                ----------------------------------------------------------------------------------------------------------------"""
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

