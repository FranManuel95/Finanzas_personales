# Finanzas personales — FM y Lucía

Dashboard de finanzas personales en Google Sheets con bot de Telegram integrado,
todo sobre Google Apps Script (sin servidor, sin LLM, gratis).

## Qué incluye

- **Dashboard** con ingresos, gastos compartidos (fijos y variables), gastos
  individuales de cada uno, balance del mes y seguimiento de ahorro
  (objetivo mensual + acumulado anual).
- **Bote común fijo**: cada uno aporta una cantidad mensual acordada y de
  ahí salen los gastos compartidos.
- **Bot de Telegram** con flujo guiado por botones (sin escribir comandos
  ni depender de IA). Permite registrar ingresos y gastos desde el móvil
  en pocos toques.

## Cómo se monta

Sigue [docs/INSTALACION.md](docs/INSTALACION.md) paso a paso. Resumen:

1. Crear una hoja de cálculo en blanco en Google Sheets.
2. Abrir el editor de Apps Script y pegar los archivos `.gs` de este repo
   (o configura el despliegue automático con clasp — ver
   [docs/CLASP_DEPLOY.md](docs/CLASP_DEPLOY.md)).
3. Ejecutar `crearDashboard()` una vez → te genera todas las pestañas.
4. Crear el bot con `@BotFather` en Telegram, pegar el token en `Config.gs`.
5. Desplegar el Apps Script como Web App y registrar el webhook.

## Estructura del repo

```
apps-script/        ← código que se sube a Apps Script (clasp rootDir)
  appsscript.json   ← manifest del proyecto
  Bot.gs            ← bot de Telegram (entrypoint doPost)
  Config.gs         ← config + lectura de Ajustes
  Notificaciones.gs ← triggers programados
  CrearDashboard.gs ← genera el libro entero (función crearDashboard)
.clasp.json         ← apunta al scriptId del proyecto
.github/workflows/  ← despliegue automático a Apps Script en cada push
docs/               ← guías
```

## Estructura del Sheet

Ver [docs/ESTRUCTURA.md](docs/ESTRUCTURA.md).
