# MINEX V15 — Corrección de importación de empleados

- La confirmación de importación ya no queda bloqueada por registros con error.
- Se habilita la confirmación si existe al menos un registro válido.
- Duplicados y registros con error se omiten y se informa al usuario antes de reemplazar la base.
- La carga de empleados reemplaza completamente la base anterior con los registros válidos de la carga más reciente.
- Se reforzó el reconocimiento de encabezados con espacios, tildes, barras y variantes comunes.
