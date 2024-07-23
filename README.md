# Backup Automation Script

## Descripción

Este proyecto se encarga de gestionar y procesar archivos de respaldo, generando bitácoras semanales en formato Excel y verificando el estado de varios perfiles mediante la lectura de archivos de registro.

## Requisitos

- Python 3.x
- Bibliotecas: `openpyxl`, `dateutil`, `logging`, `configparser`

Puedes instalar las bibliotecas necesarias usando pip:

pip install openpyxl python-dateutil

## Estructura del Proyecto

El proyecto contiene los siguientes archivos y directorios:

- `Bitacora_APP.py`: Script principal para la generación de la bitácora.
- `variables.ini`: Archivo de configuración para variables.
- `config.ini`: Archivo de configuración para rutas y posiciones de columnas.
- `Bitacora_APP.xlsx`: Plantilla base para la bitácora en formato Excel.
- `archivo_log.txt`: Archivo de registro para errores.

## Funcionalidades

- **Carga de Configuración**: Lee archivos de configuración para rutas y posiciones de columnas.
- **Generación de Carpeta**: Crea una copia del archivo base de Excel para el año actual si no existe.
- **Marcado de Archivos**: Verifica la existencia de archivos y marca las posiciones correspondientes en la bitácora.
- **Copia de Filas**: Copia el estilo de una fila a otra en el archivo de Excel.
- **Procesamiento de Perfiles**: Lee y analiza archivos de registro para determinar el estado de los perfiles y actualiza la bitácora en consecuencia.

## Uso

- **Configuración Inicial**:
  - Asegúrate de tener los archivos `variables.ini` y `config.ini` configurados correctamente.
  - Coloca la plantilla `Bitacora_APP.xlsx` en la carpeta especificada en la configuración.

- **Ejecución del Script**:
  - Ejecuta el script principal `Bitacora_APP.py` para iniciar el proceso de generación de la bitácora.

    ```bash
    python Bitacora_APP.py
    ```

- **Revisión de Logs**:
  - Verifica el archivo `archivo_log.txt` para errores y detalles del procesamiento.

## Funciones Principales

- **`cargar_configuracion()`**: Carga la configuración desde los archivos `.ini`.
- **`marcar(r, subcarpeta_, archivo1, archivo2, posicion_, incremento)`**: Marca la presencia de archivos en las rutas especificadas.
- **`verificar_ruta(ruta)`**: Verifica la existencia y permisos de una ruta.
- **`generacionCarpeta()`**: Genera la carpeta de bitácora para el año actual.
- **`convFecha(f)`**: Convierte una fecha en el formato adecuado.
- **`calcular_posicion_columna(fecha)`**: Calcula la posición de columna en función de la fecha.
- **`Bitacora_APP()`**: Función principal para el procesamiento de la bitácora.
- **`copiar_fila(origen_row, origen_col, destino_row, destino_col, sh)`**: Copia el estilo de una fila a otra en el archivo de Excel.
- **`detectar_formato_fecha(fecha)`**: Detecta y convierte el formato de una fecha.
- **`leer_ultimas_lineas(nombre_archivo, num_lineas)`**: Lee las últimas líneas de un archivo de log.
- **`limpiar_cadena(fecha_str)`**: Limpia una cadena de fecha de caracteres no deseados.
- **`convertir_formato(fecha_str)`**: Convierte una cadena de fecha en un objeto `datetime`.
- **`marcar_perfiles()`**: Marca el estado de los perfiles basado en archivos de log.

