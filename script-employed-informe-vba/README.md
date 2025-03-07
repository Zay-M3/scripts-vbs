# script-employed-informe-vba

Este directorio contiene scripts en VBA diseñados para automatizar la búsqueda y generación de informes de empleados en archivos de Excel. A continuación se describen los scripts incluidos y su funcionamiento.

## Scripts

### 1. `form-search-employee.bas`

Este script contiene un formulario de búsqueda de empleados que permite al usuario ingresar un nombre y un rango de fechas para buscar y generar un informe de empleados en un archivo de Excel.

#### Funcionamiento:
- El usuario debe completar el formulario con el nombre del empleado y las fechas de inicio y fin.
- Si alguno de los campos está vacío, se muestra un mensaje de advertencia solicitando que se complete el formulario.
- El script convierte las fechas ingresadas en formato de fecha y llama a la función `copiar_pegar` con el nombre del empleado y las fechas.
- Una vez completada la búsqueda, el formulario se cierra automáticamente.

#### Inicialización del Formulario:
- Al inicializar el formulario, se llenan los cuadros de selección de día y mes con los valores correspondientes.
- Los cuadros de selección de año se llenan con los años 2023 y 2024.

### 2. `copy-paste.bas`

Este script copia datos de empleados desde una hoja de origen a una nueva hoja en un archivo de Excel existente, aplica formato y añade fórmulas específicas para el informe.

#### Funcionamiento:
- Abre un archivo de Excel existente o crea uno nuevo si no existe.
- Copia los datos de empleados desde una hoja de origen a una nueva hoja en el archivo de Excel.
- Aplica formato a la nueva hoja, ajustando el ancho de columnas y filas.
- Inserta nuevas filas y columnas según sea necesario y añade fórmulas específicas para el informe.
- Aplica bordes a la tabla y muestra un mensaje de confirmación una vez que el informe ha sido generado exitosamente.

## Licencia

Este proyecto está licenciado bajo la Licencia MIT. Consulta el archivo [LICENSE](../LICENSE) para obtener más detalles.