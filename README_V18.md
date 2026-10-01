# MINEX V18 — Fechas y estado automáticos en casos

En el formulario de casos, dentro de la sección de fechas y estado:
- Solo se pueden editar **Fecha de seguimiento** y **Fecha de cierre**.
- **Mes cierre** se calcula automáticamente a partir de la fecha de cierre.
- **Estado** se calcula automáticamente: con fecha de cierre = CERRADO; sin fecha de cierre = EN SEGUIMIENTO.
- **Días acumulados** se calcula automáticamente desde la fecha del caso hasta la fecha de cierre o, si sigue abierto, hasta hoy. El cálculo es inclusivo.
- Al poner o quitar la fecha de cierre, los campos calculados se actualizan inmediatamente.
