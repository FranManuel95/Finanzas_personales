# Despliegue automático con clasp

A partir de ahora, **cada push a `main` o a una rama `claude/**` actualiza el
código en tu proyecto de Apps Script automáticamente**. No más copiar-pegar
los 4 `.gs` a mano.

## Cómo está montado

- Los 4 archivos `.gs` viven en `apps-script/` junto con el manifest
  `appsscript.json`.
- `.clasp.json` (raíz del repo) apunta a tu `scriptId`.
- El workflow `.github/workflows/clasp-push.yml` se ejecuta en cada push
  que toque `apps-script/`. Hace `clasp push --force` con tus credenciales
  OAuth guardadas como secret de GitHub.

## Setup inicial (una sola vez)

Son 3 pasos. Te llevan ~5 minutos.

### 1. Pon tu `scriptId` en `.clasp.json`

Abre el editor de Apps Script de tu proyecto. La URL es así:

```
https://script.google.com/u/0/home/projects/AKfycb...XYZ/edit
```

El tramo entre `projects/` y `/edit` es tu **scriptId**. Cópialo y pégalo
en `.clasp.json`:

```json
{ "scriptId": "AKfycb...XYZ", "rootDir": "apps-script" }
```

Commit + push.

### 2. Genera tus credenciales OAuth de clasp

En **tu máquina** (Mac/Windows/Linux con Node 18+):

```bash
npm install -g @google/clasp
clasp login
```

Se abre el navegador, autorizas con tu cuenta Google (la dueña del Sheet
y del Apps Script), y clasp guarda un archivo en `~/.clasprc.json`.

> Si no quieres instalar nada local, dímelo y lo hacemos en esta sesión
> con `clasp login --no-localhost` — pero el archivo se queda en un
> contenedor efímero, así que hay que moverlo a GitHub Secrets enseguida.

### 3. Guarda `~/.clasprc.json` como secret de GitHub

1. Abre el archivo en un editor de texto y copia su contenido completo
   (es un JSON corto).
2. En GitHub: `Settings` → `Secrets and variables` → `Actions` → `New
   repository secret`.
3. Nombre: `CLASPRC_JSON`. Valor: pega el JSON. Guarda.

Hecho. El siguiente push a `apps-script/**` dispara el workflow y sube
el código.

## Flujo a partir de ahora

```
Yo (Claude) modifico los .gs en el repo
       │
       └─► git push
              │
              └─► GitHub Action ejecuta `clasp push`
                     │
                     └─► Tu Apps Script queda actualizado
```

Lo único que sigues haciendo tú a mano (porque el bot necesita una nueva
versión publicada para que Telegram lo vea):

1. **Re-desplegar la Web App** cuando cambie el bot: editor → Implementar
   → Gestionar implementaciones → ✏ Editar → Nueva versión → Implementar.
2. **Registrar el webhook** la primera vez: en el editor, ejecuta
   `registrarWebhook` una vez (después no hace falta más).
3. **Ejecutar `crearDashboard`** la primera vez para generar el libro.

Estos 3 también pueden automatizarse con `clasp run` si habilitas la
Apps Script API en el proyecto Google Cloud — dímelo si quieres ese paso.

## Troubleshooting

- **El workflow falla con "Could not read API credentials"**: el secret
  `CLASPRC_JSON` no existe o está vacío. Repite el paso 3.
- **El workflow falla con "Project file (.clasp.json) not found"**: el
  `scriptId` en `.clasp.json` no está bien. Revisa el paso 1.
- **`clasp push` borra mi código del editor**: imposible — el workflow
  usa `--force` que sobreescribe el editor con el contenido del repo. Si
  tienes cambios sin commitear en el editor, hazlos en el repo primero.
