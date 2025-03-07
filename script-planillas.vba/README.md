# script-planillas.vba

Este directorio contiene scripts en VBA diseñados para automatizar la creación y el formateo de archivos de Excel. A continuación se describen los scripts incluidos y su funcionamiento.

## Scripts

### 1. `create-excel.bas`

Este script crea un nuevo archivo de Excel con un nombre especificado en la celda `E24` de la hoja `Hoja1` del libro actual.

#### Funcionamiento:
- Obtiene el nombre del archivo de la celda `E24`.
- Verifica si el nombre está vacío y muestra un mensaje de advertencia si es así.
- Crea un nuevo libro de Excel y lo guarda en la carpeta `Documents` del usuario con el nombre especificado.
- Muestra un mensaje de confirmación una vez que el archivo ha sido creado exitosamente.

### 2. `script-planillas-format.bas`

Este script copia datos de una hoja de origen a una nueva hoja en un archivo de Excel existente, aplica formato y añade fórmulas.

#### Funcionamiento:
- Abre un archivo de Excel existente especificado en la celda `E24` de la hoja `Hoja1`.
- Copia un rango de datos de la hoja `Hoja1` a una nueva hoja en el archivo existente.
- Renombra la nueva hoja y ajusta el ancho de columnas y filas.
- Inserta nuevas filas y columnas, y añade fórmulas.
- Aplica bordes a la tabla y muestra un mensaje de confirmación una vez que los datos han sido copiados y modificados exitosamente.

## Licencia

Este proyecto está licenciado bajo la Licencia MIT. Consulta el archivo [LICENSE](../LICENSE) para obtener más detalles.