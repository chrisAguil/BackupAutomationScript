# Bitácora APP

## Descripción

Este script de Python gestiona la generación de una bitácora de respaldo para diversos sistemas. Realiza tareas como copiar archivos de respaldo, actualizar registros en un archivo de Excel y procesar logs. Utiliza librerías como `openpyxl` para manipular archivos de Excel y `configparser` para leer archivos de configuración.

## Requisitos

- Python 3.x
- Librerías:
  - `python-dateutil`
  - `openpyxl`
  - `configparser`
  - `shutil`
  - `logging`
  - `re`

## Estructura del Proyecto

El proyecto consta de un script principal `main.py` y dos archivos de configuración: `variables.ini` y `config.ini`. La estructura del directorio es la siguiente:

## Funciones Principales

### cargar_configuracion()
**Qué:** Carga la configuración desde archivos INI y establece variables globales necesarias para el funcionamiento del script.  
**Cómo:** Lee los archivos `variables.ini` y `config.ini` para configurar rutas y otros parámetros.  
**Dónde:** Se utiliza al inicio del script para inicializar la configuración.  
**Por Qué:** Es esencial para definir rutas y configuraciones que el script necesita para funcionar correctamente.  

**Parámetros:**  
Ninguno.

---

### marcar(r, subcarpeta_, archivo1, archivo2, posicion_, incremento)
**Qué:** Marca archivos presentes en las rutas especificadas.  
**Cómo:** Verifica la existencia de archivos en subcarpetas específicas y actualiza una lista de contadores según la posición.  
**Dónde:** Utilizado para controlar la presencia de archivos en diferentes rutas.  
**Por Qué:** Permite el seguimiento de archivos y asegura que se encuentren en las ubicaciones correctas.  

**Parámetros:**  
- `r` (str): Ruta base para buscar los archivos.  
- `subcarpeta_` (str): Subcarpeta dentro de la ruta base.  
- `archivo1` (str): Nombre del primer archivo a buscar.  
- `archivo2` (str): Nombre del segundo archivo a buscar.  
- `posicion_` (int): Índice en la lista de contadores.  
- `incremento` (int): Valor a incrementar en el contador si el archivo existe.  

---

### verificar_ruta(ruta)
**Qué:** Verifica si una ruta existe y si tiene permisos de lectura y escritura.  
**Cómo:** Realiza comprobaciones para asegurar que la ruta es accesible y que los permisos son adecuados.  
**Dónde:** Se usa antes de realizar operaciones de archivo para evitar errores relacionados con rutas no válidas.  
**Por Qué:** Asegura que las rutas necesarias están disponibles y accesibles para evitar fallos en el script.  

**Parámetros:**  
- `ruta` (str): Ruta a verificar.  

---

### generacionCarpeta()
**Qué:** Genera un archivo de bitácora para el año actual si no existe.  
**Cómo:** Copia el archivo base de bitácora a una nueva ubicación con el nombre del año actual.  
**Dónde:** Utilizado para crear un archivo de bitácora semanalmente.  
**Por Qué:** Asegura que existe un archivo de bitácora para el año en curso, facilitando el registro de datos.  

**Parámetros:**  
Ninguno.

---

### convFecha(f)
**Qué:** Convierte una fecha a un formato específico.  
**Cómo:** Calcula y formatea el mes y el día de una fecha y su día anterior.  
**Dónde:** Utilizado para operaciones relacionadas con fechas dentro del script.  
**Por Qué:** Facilita el manejo de fechas en formatos necesarios para otras operaciones.  

**Parámetros:**  
- `f` (datetime.date): Fecha que se desea convertir.  

---

### calcular_posicion_columna(fecha)
**Qué:** Calcula la posición de columna en base a una fecha.  
**Cómo:** Calcula la diferencia en semanas entre la fecha proporcionada y una fecha de referencia para determinar la posición de columna.  
**Dónde:** Utilizado para ubicar datos en una hoja de cálculo.  
**Por Qué:** Permite ubicar correctamente los datos en función de la fecha.  

**Parámetros:**  
- `fecha` (datetime.date): Fecha para calcular la posición de columna.  

---

### Bitacora_APP()
**Qué:** Ejecuta el proceso principal para actualizar la bitácora semanalmente.  
**Cómo:** Realiza varias operaciones como copiar columnas, marcar archivos y actualizar una hoja de cálculo con los datos correspondientes.  
**Dónde:** Función principal que se ejecuta periódicamente para actualizar la bitácora.  
**Por Qué:** Mantiene actualizada la bitácora con la información relevante semanalmente.  

**Parámetros:**  
Ninguno.

---

### copiar_fila(origen_row, origen_col, detino_row, detino_col, sh)
**Qué:** Copia el estilo de una celda a otra en una hoja de cálculo.  
**Cómo:** Copia el estilo de la celda en la fila y columna de origen a la celda en la fila y columna de destino.  
**Dónde:** Utilizado para mantener el formato consistente en hojas de cálculo.  
**Por Qué:** Asegura que el formato de las celdas se copie correctamente a nuevas ubicaciones.  

**Parámetros:**  
- `origen_row` (int): Fila de origen de la celda.  
- `origen_col` (int): Columna de origen de la celda.  
- `detino_row` (int): Fila de destino de la celda.  
- `detino_col` (int): Columna de destino de la celda.  
- `sh` (openpyxl.worksheet.worksheet.Worksheet): Hoja de cálculo en la que se realiza la operación.  

---

### detectar_formato_fecha(fecha)
**Qué:** Detecta y formatea una cadena de fecha en un formato específico.  
**Cómo:** Utiliza el módulo `dateutil` para analizar y convertir el formato de la fecha.  
**Dónde:** Utilizado para manejar fechas en diferentes formatos.  
**Por Qué:** Permite la conversión y análisis de fechas en formatos variados.  

**Parámetros:**  
- `fecha` (str): Fecha en formato de cadena a analizar.  

---

### leer_ultimas_lineas(nombre_archivo, num_lineas)
**Qué:** Lee las últimas líneas de un archivo de log y extrae información relevante.  
**Cómo:** Lee el archivo de log, extrae las últimas líneas y procesa la fecha para su análisis.  
**Dónde:** Utilizado para analizar archivos de log para verificar errores o información de perfil.  
**Por Qué:** Facilita la obtención y análisis de la información más reciente de los archivos de log.  

**Parámetros:**  
- `nombre_archivo` (str): Nombre del archivo de log a leer.  
- `num_lineas` (int): Número de líneas a leer desde el final del archivo.  

---

### limpiar_cadena(fecha_str)
**Qué:** Limpia una cadena de fecha eliminando caracteres no deseados.  
**Cómo:** Utiliza expresiones regulares para eliminar caracteres no alfanuméricos.  
**Dónde:** Utilizado para preparar cadenas de fecha para su análisis.  
**Por Qué:** Asegura que las fechas sean procesadas correctamente sin caracteres no deseados.  

**Parámetros:**  
- `fecha_str` (str): Cadena de fecha que se desea limpiar.  

---

### convertir_formato(fecha_str)
**Qué:** Convierte una cadena de fecha a un formato de fecha de Python.  
**Cómo:** Utiliza el formato de fecha específico para analizar y extraer mes, día y año.  
**Dónde:** Utilizado para convertir fechas en formatos legibles por Python.  
**Por Qué:** Facilita la conversión de fechas en diferentes formatos para su procesamiento en Python.  

**Parámetros:**  
- `fecha_str` (str): Cadena de fecha en formato específico a convertir.  

---

### marcar_perfiles()
**Qué:** Marca el estado de perfiles basándose en archivos de log.  
**Cómo:** Lee los archivos de log, verifica errores y actualiza la hoja de cálculo con los resultados.  
**Dónde:** Utilizado para registrar y marcar errores en perfiles.  
**Por Qué:** Permite un seguimiento preciso del estado de perfiles y errores asociados.  

**Parámetros:**  
Ninguno.

## Flujo General

1. **Inicio:** El script comienza cargando la configuración y estableciendo parámetros.
2. **Generación de Carpeta:** Se crea una nueva bitácora para el año actual si no existe.
3. **Procesamiento de Archivos:** Se verifican y marcan archivos según su existencia.
4. **Actualización de Excel:** Se actualiza el archivo de Excel con la información de los archivos y logs procesados.
5. **Procesamiento de Logs:** Se leen y procesan los logs para marcar perfiles y registrar errores.

## Resumen de Explicación del Código

El script `main.py` está diseñado para automatizar el proceso de gestión de una bitácora de respaldo. Utiliza archivos de configuración para definir rutas y parámetros, verifica la existencia de archivos en esas rutas, y actualiza un archivo de Excel con los resultados. Además, procesa logs para marcar perfiles y errores. El flujo del script asegura que los archivos y registros se manejen de manera eficiente y se mantenga un registro actualizado de los respaldos realizados.



## Dependencias
Para ejecutar este script, es necesario tener Python 3 instalado.

### Instalación de Python

**Para Linux:**

1. Para instalar Python 3:
    ```bash
    sudo dnf update
    sudo dnf install python3
    ```

2. Para actualizar Python 3 a la última versión:
    ```bash
    sudo dnf update
    sudo dnf upgrade python3
    ```
**Si estás utilizando una distribución basada en Debian (como Ubuntu o Debian), reemplaza `dnf` por `apt`. Debería funcionar sin problemas.**

**Para macOS:**

1. Python 2 viene preinstalado en macOS. Para instalar Python 3 utilizando Homebrew:
    ```bash
    brew install python@3
    ```

2. Para actualizar Python 3 a la última versión:
    ```bash
    brew upgrade python@3
    ```

**Para Windows:**

1. Descarga el instalador de Python desde el sitio web oficial [Python.org](https://www.python.org/downloads/) y ejecútalo. Asegúrate de marcar la casilla "Add Python to PATH" durante la instalación para poder acceder a Python desde la línea de comandos.

2. Para actualizar Python en Windows, descarga el instalador de la versión más reciente desde el sitio web oficial y ejecútalo, seleccionando "Modificar" o "Actualizar" para actualizar tu instalación existente.

### Creación del entorno virtual

**Para Linux y macOS:**

1. Crear un nuevo entorno virtual en el directorio 'venv':
    ```bash
    python3 -m venv venv
    ```

2. Activar el entorno virtual:
    ```bash
    source venv/bin/activate
    ```

   Ahora estás dentro del entorno virtual 'venv'.

3. Para desactivar el entorno virtual, simplemente ejecuta:
    ```bash
    deactivate
    ```

**Para Windows:**

1. Crear un nuevo entorno virtual en el directorio 'venv':
    ```bash
    python -m venv venv
    ```

2. Activar el entorno virtual:
    ```bash
    venv\Scripts\activate
    ```

   Ahora estás dentro del entorno virtual 'venv'.

3. Para desactivar el entorno virtual en Windows, ejecuta el mismo comando que para Linux:
    ```bash
    deactivate
    ```

Una vez creado el entorno, puedes proceder a instalar las librerías necesarias. Este programa requiere la instalación de varios módulos:
- **openpyxl**
- **configparser**
- **python-dateutil**

Para instalar estos módulos, ejecuta el siguiente comando en la terminal desde el directorio actual del proyecto:

```bash
pip install -r requirements.txt
```



## Estructura del proyecto
BITACORA_BD_V2/
* /
  * mainv2.py              # Script principal para ejecutar el llenado de la bitácora
  * config.ini             # Archivo de configuración principal
  * configMarcado.conf     # Archivo de variables adicionales
  * marcarAltex.jar        # Archivo Java que se ejecuta durante el proceso
  * Marcado_altex.lobo     # Archivo de retorno del ejecutable jar
  * README.md              # Documentación del proyecto
    
* Bitacora_de_respaldos_BD/
  * Bitacora_APP.xlsx      # Archivo Excel base de la bitácora
  * Bitacora_APP_2024.xlsx # Archivo Excel de la bitácora para el año 2024  
