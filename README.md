# MINEX | Gestión de Relaciones Laborales

## Versión actual: V27.27 — 10/10/2026

## V27.27 — Rendimiento del menú
- Se reutilizan métricas, agrupaciones de informes, detección de duplicados y categorías de seguimientos hasta que cambian los datos.
- Las búsquedas globales reutilizan el texto normalizado de cada registro.
- Empleados carga 100 filas por vez y ofrece un botón para mostrar más.
- Al abrir Casos, Seguimientos o Empleados se inicia con una cantidad de filas acotada.
- La sincronización con Supabase se mantiene sin cambios.

## V27.26 — Menú más ágil
- Se reutilizan resultados de indicadores, seguimientos, reportes y detección de duplicados mientras los datos no cambien.
- Las búsquedas reutilizan índices de texto por registro para evitar normalizar todos los campos en cada actualización.
- Las métricas se calculan en un recorrido de los casos y se conservan para las pantallas que las comparten.
- Se actualiza la versión de caché de la aplicación.

## V27.25 — Navegación de seguimientos más ágil
- Las secciones reutilizan los datos cargados y no vuelven a descargarlos al navegar.
- Seguimientos muestra inicialmente hasta 50 registros por categoría y permite cargar más.
- La búsqueda de empleados actualiza solo los resultados y conserva el foco en el campo.
- Se renueva la versión de caché para que los navegadores reciban los archivos publicados más recientes.

## V27.24 — Centros de trabajo unificados y mejor fluidez
- Se agrupan variantes claras de centros de trabajo para evitar opciones duplicadas en los filtros.
- Se normalizan variantes de Dinastía, Edificio Calle 13, Gran Estación, Alto Viento, Lajas, Patio Centro, Patio La Lejía, La Unión y Esmeralda.
- El campo Centro de trabajo permite seleccionar una sugerencia existente o escribir un centro nuevo.
- Se conservan separadas las opciones que pueden representar ubicaciones distintas, como Patio Centro, Patio Centro 1 y Patio Centro 2.
- La tabla de Casos carga 100 filas por vez y permite mostrar más cuando se necesite.
- La búsqueda de Casos espera brevemente a que termine la escritura antes de filtrar y reconstruir la lista.
- Se reutilizan opciones y cálculos de días hábiles ya procesados mientras no cambien sus datos.

## V27.22 — Fecha límite legal y priorización simple
- Se agrega una fecha límite legal independiente, resaltada en negrilla en el formulario y en la lista de Casos.
- La fecha límite legal no modifica días acumulados, promedios ni otros contadores.
- Se agregan prioridad y responsable asignado para organizar los casos.
- La lista permite filtrar por prioridad y responsable y ordenar primero los casos urgentes.
- Las nuevas propiedades se conservan al cargar y exportar los informes de detalle.
- El cálculo de días de gestión sigue siendo independiente de los vencimientos legales.

## V27.21 — Días hábiles de gestión en Casos
- Los días acumulados se cuentan desde el día siguiente a la radicación, de lunes a viernes, excluyendo festivos nacionales de Colombia.
- Los casos abiertos se cuentan hasta hoy y los cerrados hasta su fecha de cierre.
- El cálculo se aplica al abrir Casos, en el formulario, en los informes y en la exportación oficial.
- La cifra es un indicador interno de gestión; los vencimientos legales deben controlarse según la norma aplicable a cada trámite.

## V26.4 — Corrección de desbordamiento de filtros
- Se corrige el ancho mínimo de los campos de fecha nativos.
- Los grupos de filtros usan columnas flexibles con `minmax(0, 1fr)`.
- En resoluciones intermedias los filtros pasan a dos columnas para conservar legibilidad.
- Se evita el desbordamiento horizontal de filtros sin ocultar información.

## V26.3 — Experiencia responsive y filtros de Casos
- Se optimiza la pestaña **Casos** con grupos de filtros visualmente diferenciados.
- Se agregan indicadores de filtros activos y chips para retirar criterios individualmente.
- Los rangos de fechas muestran claramente los campos Desde y Hasta.
- Se mejora la búsqueda global, el contador de resultados y la jerarquía visual.
- Se incorpora una adaptación responsive global para navegación, paneles, formularios, tablas y acciones en móviles y tabletas.
- Se mantiene la operación en GitHub Pages sin introducir backend ni dependencia de autenticación ficticia.

Sistema web estático para la gestión de casos de Relaciones Laborales. GitHub es únicamente el medio de publicación del proyecto; MINEX mantiene su operación local mediante IndexedDB.

## V26.1 — Corrección estructural de descargas
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

## V26.2 — Descarga de casos en el formato oficial (corrección)
- La descarga ya no depende de rutas que no existían en el repositorio (`./jszip.min.js` y `./data/case_template_base64.js`). `index.html` solo carga `app.js`; JSZip y la plantilla se cargan desde `app.js` con respaldos (`jszip.min.js` → `jszip_min.js` → CDN).
- Nueva plantilla embebida `case_template_base64.js` (hoja `VOLUMETRIAS`, `Tabla1`, estilos, anchos y columnas T:U ocultas, tomados de `CARGA0610.xlsx`). Contiene solo encabezados y una fila modelo.
- Las fórmulas se escriben con las fórmulas de columna calculada de `Tabla1` (antes se copiaba la fila 2, cuyas columnas VISIBLE_DASH y MATERNIDAD_DASH tenían fórmulas incorrectas). Además se guardan valores calculados para que el archivo se vea bien en cualquier visor.
- Se elimina `calcChain.xml` al exportar (quedaba desactualizado y Excel podía marcar el archivo como dañado).
- Se ajusta el alto de fila al texto, se limpian caracteres inválidos en XML y se aceptan fechas ISO o DD/MM/AAAA.
- «Exportar», «Seleccionados» y «Duplicados en revisión» descargan ahora en el mismo formato oficial.
- La detección y revisión de duplicados no se modificó.

## Carga incremental de datos
- Las nuevas cargas de casos y empleados comparan cada fila con la base existente.
- Las coincidencias se actualizan conservando su identificador y los registros nuevos se agregan.
- La información existente que no aparezca en el archivo no se elimina.
- Los duplicados dentro del mismo archivo continúan en revisión antes de incorporarse.

## Trabajo multiusuario con Supabase
- Los casos y empleados se almacenan en Supabase cuando el usuario inicia sesión.
- La bandeja de duplicados compartida usa `public.duplicate_reviews`.
- Los cambios remotos de casos, empleados y duplicados se actualizan mediante Supabase Realtime.
- Ejecuta una vez [`supabase-schema.sql`](./supabase-schema.sql) en Supabase > SQL Editor antes de usar la bandeja de duplicados compartida.
- La migración inicial compara los datos locales con los remotos; no borra información ni detiene la integración porque Supabase ya tenga registros.

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

### V26.2
- Corrección de la descarga de casos en el formato oficial (ver arriba).

### V26.1
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


## V26.1 — corrección de arranque
- Se restauró la función `openDB()` de IndexedDB que faltaba en V26.
- Se conserva el almacenamiento local con los almacenes `cases`, `employees`, `volumetries`, `config` y `duplicateReviews`.
- No se modifica la lógica de revisión de duplicados ni la exportación XLSX.
