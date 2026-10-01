# MINEX V11

## Publicación correcta en GitHub Pages

Sube **el contenido de este ZIP**, no la carpeta contenedora. `index.html` debe quedar en la raíz del repositorio.

La versión incluye `data_bootstrap.js`, por lo que la aplicación puede iniciar incluso si GitHub Pages no sirve correctamente la carpeta `data/`; aun así se recomienda publicar también `data/`.

Estructura:
- index.html
- app.js
- styles.css
- data_bootstrap.js
- data/
- assets/ (si aplica)

Después de publicar, usa Ctrl+F5.


## V12
- Los datos existentes se conservan al recargar y actualizar la aplicación mediante IndexedDB.
- Se incorporan acciones separadas para eliminar casos, empleados, volumetrías o toda la información.
- Restaurar datos iniciales es una acción independiente y explícita.
