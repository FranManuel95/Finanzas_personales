# Estructura del libro

Tras ejecutar `crearDashboard()` tendrás estas pestañas:

## `Dashboard`
Resumen del **mes activo** (definido en `Config`). Lee de las demás
pestañas con `SUMIFS`. Bloques:

- Ingresos del mes (FM, Lucía, total).
- Bote común (aportación fija de cada uno + total).
- Gastos compartidos (fijos, variables, total, diferencia frente al bote).
- Gastos individuales (FM y Lucía).
- Balance personal del mes (cuánto le queda a cada uno tras ingresos −
  aportación al bote − gastos propios).
- Ahorro (objetivo, real, % cumplimiento, acumulado del año).

## `Config`
Tabla clave-valor con los parámetros que tocas a mano:

| Parámetro                                | Ejemplo            |
|------------------------------------------|--------------------|
| Mes activo                               | `2026-05`          |
| Objetivo ahorro mensual conjunto (€)     | `400`              |
| Telegram chat IDs autorizados            | `123456, 987654`   |
| Telegram bot token                       | `1234:AAH…`        |

> Las aportaciones al bote común NO son fijas: se registran como
> movimientos desde el bot (botón 🏦) y se acumulan en la pestaña
> `Aportaciones_Bote`.

## `Categorias`
Tres columnas: *Compartido fijo*, *Compartido variable*, *Individual*.
Editables. Sirven de validación y como botones en el bot.

## `Ingresos`
`Fecha | Persona (FM/Lucía) | Concepto | Importe (€) | Recurrente (Sí/No)`

## `Aportaciones_Bote`
`Fecha | Persona (FM/Lucía) | Concepto | Importe (€)`

Cada vez que uno de los dos mete dinero al bote común, lo registra desde
el bot. El Dashboard suma las aportaciones del mes activo por persona.

## `Gastos_Compartidos_Fijos`
`Fecha | Concepto | Categoría | Importe (€) | Día de cargo`

## `Gastos_Compartidos_Variables`
`Fecha | Concepto | Categoría | Importe (€) | Pagado por (Bote común/FM/Lucía)`

> *"Pagado por"* permite registrar cuándo, excepcionalmente, ha pagado
> uno de los dos un gasto del bote común (para reembolsar después).

## `Gastos_FM` y `Gastos_Lucia`
`Fecha | Concepto | Categoría | Importe (€)`

## `Ahorro`
Una fila por mes del año en curso. Las columnas `Real` y `Acumulado año`
son **fórmulas** que se calculan solas leyendo de las pestañas de
ingresos y gastos.

## Datos = la verdad
El bot solo añade filas; nunca borra ni modifica. Si te equivocas en un
registro, edítalo directamente en la pestaña correspondiente.
