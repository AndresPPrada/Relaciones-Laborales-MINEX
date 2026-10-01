# MINEX | Gestión de Relaciones Laborales

## Versión actual: V26 — 01/10/2026

Sistema web estático para la gestión de casos de Relaciones Laborales. GitHub es únicamente el medio de publicación del proyecto; MINEX mantiene su operación local mediante IndexedDB.

## V26 — Corrección estructural de descargas
- Se reconstruye la base sobre la versión V25.8, restaurando las funciones internas que habían quedado incompletas.
- Se elimina la dependencia de ExcelJS/CDN para la descarga de casos.
- Se incorpora **JSZip local** dentro del proyecto para generar el archivo directamente en el navegador.
- La plantilla oficial utilizada es exactamente el archivo `Libro3.xlsx` suministrado, conservando la hoja `VOLUMETRIAS`, formatos, fórmulas, tabla `Tabla1`, columnas ocultas y estructura del libro.
- La descarga de casos modifica únicamente el XML de datos necesario dentro del XLSX; el resto del libro se conserva sin reconstruirlo desde cero.
- Se actualiza automáticamente el rango de `Tabla1` al número real de casos exportados.
- Las fórmulas de columnas calculadas se conservan y se ajustan a cada fila.
- Se fuerza el recálculo al abrir el archivo en Excel.
- Se incorpora una copia embebida de la plantilla como respaldo si GitHub Pages no logra cargar el archivo desde `assets`.
- La detección y revisión de duplicados se conserva.
- Se mantiene un único `README.md`; no se crean READMEs por versión.

## Funcionalidades conservadas
- Dashboard.
- Casos: creación, consulta, edición, cierre y seguimiento.
- Seguimientos y pendientes.
- Informes e indicadores.
- Base de empleados.
- Importación y exportación.
- Filtros y consultas.
- Detección de posibles duplicados.
- Revisión de duplicados antes de incorporar registros al maestro.
- Detalle de la solicitud.
- IndexedDB como almacenamiento operativo.
- Scroll y filtros de texto mejorados.

## Plantilla oficial de Casos
Archivo fuente: `Libro3.xlsx`.

Archivo incluido en el proyecto: `assets/FORMATO_CARGA_CASOS_RELACIONES_LABORALES.xlsx`.

La exportación no crea una hoja plana nueva: parte de la plantilla corporativa y conserva su estructura.

## Regla crítica: duplicados
La detección de posibles duplicados y la cola de **Revisión de duplicados** son funcionalidades críticas y no deben eliminarse ni simplificarse.

## Principio de desarrollo
Cada nueva versión parte de la anterior y conserva los avances existentes. Este `README.md` es el único README oficial y se actualiza en cada edición.

## Historial

### V26
- Corrección de la arquitectura de descarga de casos.
- Exportación XLSX mediante edición directa del paquete XLSX con JSZip.
- Uso exacto de `Libro3.xlsx` como plantilla oficial.
- Respaldo embebido de la plantilla.
- Restauración de funciones internas de la aplicación que habían quedado incompletas en V25.8.

### V25.8
- Corrección de carga de plantilla para GitHub Pages.
- Separación conceptual entre GitHub (publicación) y MINEX (aplicación/datos).

### V25.7
- Plantilla oficial integrada en el proyecto.

### V25.6
- Descarga de casos con estructura corporativa de `VOLUMETRIAS`.

### V25.5
- Ajustes de exportación para evitar descargas planas.

### V25.4
- Uso de plantilla Excel corporativa.

### V25.3
- Consolidación de documentación en un único README.

### V25.2
- Mejora de filtros de texto.

### V25.1
- Corrección del desplazamiento vertical.

### V25.0
- Preparación del proyecto para GitHub Pages.

### V24
- Base funcional anterior: casos, seguimientos, empleados, informes, importación/exportación y revisión de duplicados.
