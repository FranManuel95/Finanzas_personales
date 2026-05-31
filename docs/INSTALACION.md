# Instalación paso a paso

Tiempo estimado: 15 minutos.

## 1. Crea la hoja de cálculo

1. Entra en <https://sheets.google.com> y crea una hoja en blanco.
2. Nómbrala como quieras (ej. "Finanzas FM & Lucía").
3. Compártela con la cuenta de Google de Lucía para que también pueda editar.

## 2. Pega los scripts de Apps Script

1. En la hoja: menú **Extensiones → Apps Script**.
2. Borra el `Code.gs` que viene por defecto.
3. Crea estos tres archivos (botón `+` → Script) con el contenido del repo:
   - `CrearDashboard.gs` → contenido de [`setup/CrearDashboard.gs`](../setup/CrearDashboard.gs)
   - `Config.gs` → contenido de [`bot/Config.gs`](../bot/Config.gs)
   - `Bot.gs` → contenido de [`bot/Bot.gs`](../bot/Bot.gs)
4. Guarda con **Ctrl/Cmd + S**.

## 3. Genera las pestañas del dashboard

1. En el editor, selecciona la función `crearDashboard` en el desplegable
   de la barra superior.
2. Pulsa **Ejecutar**. La primera vez te pedirá autorizar permisos:
   acepta con tu cuenta de Google (es tu propio script, no se publica).
3. Vuelve al Sheet → verás todas las pestañas creadas y la pestaña
   **Dashboard** activa.
4. Ajusta los valores de la pestaña **Config**:
   - **Mes activo**: déjalo como está (mes actual en formato `AAAA-MM`).
   - **Objetivo ahorro mensual conjunto**: meta de ahorro al mes entre los dos.
   - Telegram bot token y chat IDs los rellenamos en los pasos siguientes.

   Las aportaciones al bote común NO se ponen aquí: se registran como
   movimientos desde el bot (botón 🏦) cada vez que uno aporta dinero.

## 4. Crea el bot de Telegram

1. Abre Telegram y busca **@BotFather**.
2. Envíale `/newbot`. Te pedirá un nombre y un username (debe acabar en `bot`).
3. Te devolverá un **token** con esta pinta: `1234567890:AAH...`.
4. Pega ese token en la pestaña **Config** del Sheet, fila
   *Telegram bot token*.

## 5. Despliega el Apps Script como Web App

1. En el editor de Apps Script: **Implementar → Nueva implementación**.
2. Tipo: **Aplicación web**.
3. Configuración:
   - **Ejecutar como**: yo mismo (`tu_email@gmail.com`).
   - **Quién tiene acceso**: **Cualquier usuario** (necesario para que
     Telegram pueda llamar al webhook; el bot internamente filtra por
     chat ID autorizado).
4. Pulsa **Implementar** y autoriza si pide.
5. Copia la **URL del Web App** (acaba en `/exec`).

## 6. Registra el webhook

1. Abre el archivo `Bot.gs` en el editor.
2. Busca la función `registrarWebhook()` al final del archivo.
3. Pega la URL del Web App dentro de la constante `URL_WEBAPP`.
4. Guarda y ejecuta la función `registrarWebhook`.
5. En los logs (**Ver → Registro**) debería aparecer
   `{"ok":true,"result":true,"description":"Webhook was set"}`.

## 7. Autoriza tu chat (y el de Lucía)

1. En Telegram, abre tu bot y envíale `/id`.
2. El bot te responderá con tu **chat ID** numérico.
3. Repítelo con Lucía desde su Telegram.
4. Vuelve a la pestaña **Config** y pega ambos IDs separados por coma
   en *Telegram chat IDs autorizados* (ej. `123456789, 987654321`).

> Si dejas ese campo vacío, el bot acepta a cualquiera que conozca su
> nombre — útil para probar, mala idea para uso real.

## 8. ¡Listo!

Envía `/nuevo` al bot. Te aparecerán los botones para registrar un
movimiento. Envía `/resumen` para ver el balance del mes en pantalla.

## Comandos del bot

| Comando      | Para qué sirve                                |
|--------------|-----------------------------------------------|
| `/nuevo`     | Empieza un nuevo registro (menú principal)    |
| `/resumen`   | Resumen del mes activo                        |
| `/objetivo`  | Cambia el objetivo de ahorro mensual          |
| `/cancelar`  | Aborta el flujo actual                        |
| `/id`        | Muestra tu chat ID                            |
| `/menu`      | Igual que `/nuevo`                            |
| `/start`     | Igual que `/nuevo`                            |

## Mantenimiento

- Cambiar el **mes activo**: edita la celda en la pestaña Config.
- Añadir/quitar **categorías**: edita la pestaña Categorias. El bot las
  lee en caliente, no necesita reinicio.
- Si despliegas una versión nueva del script: usa **Implementar →
  Gestionar implementaciones → Editar (lápiz) → Nueva versión**. Así
  conservas la misma URL y no hace falta volver a registrar el webhook.
