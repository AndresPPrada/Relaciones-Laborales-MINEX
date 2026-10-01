# MINEX Gestión de Relaciones Laborales — V10

Cambios principales:
- La base de empleados utiliza exactamente las 18 columnas definidas por MINEX.
- Las columnas de empleado se normalizan (incluida DIRECCION sin espacios invisibles).
- La importación de empleados es de **reemplazo total**: cada archivo nuevo elimina la base anterior y deja únicamente la carga más reciente.
- Se detectan documentos duplicados dentro del archivo antes de confirmar.
- Los campos de fecha de empleados se normalizan.
- La carga de empleados muestra una advertencia explícita y solicita confirmación final.
- Seguimientos: filtros Vencidos, Vencen hoy, Esta semana (lunes a viernes) y Todos.
- Se conserva la arquitectura local IndexedDB y la importación/exportación existente.

## Columnas oficiales de empleados
EMPRESA CONTRATO | TIPO DOCUMENTO | DOCUMENTO | NOMBRE DEL EMPLEADO | CARGO | NIVEL EN LA ESTRUCTURA | JEFE INMEDIATO | SUB AREA | AREA / UNIDAD ORGANIZACIONAL | DIRECCION | GERENCIA | CLASIFICACION COSTO / GASTO | UBICACION | CLASIFICACION GENERAL | FECHA ANTIGÜEDAD | ULTIMA FECHA INGRESO | TELEFONO | CORREO
