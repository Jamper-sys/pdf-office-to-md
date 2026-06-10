# pdf-office-to-md

App de escritorio en Python (Windows) para convertir **PDF y archivos de
Office** (`.docx`, `.xlsx`, `.pptx`, etc.) a **Markdown** (`.md`),
construida sobre [MarkItDown](https://github.com/microsoft/markitdown).

## Características

- Interfaz gráfica con ventana propia (Tkinter — sin dependencias gráficas
  externas).
- Arrastrar y soltar archivos (requiere `tkinterdnd2`).
- Integración con el menú contextual de Windows: **clic derecho → Convertir
  a Markdown**.
- Selección de archivos sueltos o de carpetas enteras (recursivo opcional).
- Vista previa del Markdown generado dentro de la app.
- Barra de progreso con contador, botón de cancelación, ventana
  redimensionable.
- Post-procesado opcional del Markdown: une líneas partidas de PDFs, quita
  números de página sueltos, colapsa saltos en exceso.
- Front-matter YAML automático con metadatos (`title`, `source`,
  `converted_at`, número de páginas en PDFs).
- Aviso si un PDF parece escaneado y necesita OCR.
- Recuerda tus opciones entre sesiones (`%APPDATA%\convertir_a_md\config.json`).
- Sin sobrescribir: si ya existe `archivo.md`, genera `archivo (1).md`,
  `archivo (2).md`, …
- Auto-actualización: comprueba `version.json` en este repo al arrancar
  y avisa si hay versión nueva.
- También incluye una versión **CLI** (`pdf_office_to_md.py`).

## Formatos soportados

PDF · Word (`.docx`, `.doc`) · Excel (`.xlsx`, `.xls`, `.csv`) · PowerPoint
(`.pptx`, `.ppt`) · HTML · EPUB · TXT · RTF · JSON · XML · ZIP ·
Imágenes (`.png`, `.jpg` — con OCR) · Audio (`.mp3`, `.wav`, `.m4a` — con
transcripción).

## Instalación

Requiere Python 3.10 o superior en Windows.

```powershell
pip install -r requirements.txt
```

## Uso

### App de escritorio

Doble clic sobre `convertir_a_md.pyw`.

### Menú contextual de Windows

Ejecuta `registry\instalar_menu_contextual.ps1` (clic derecho → "Ejecutar
con PowerShell"). A partir de ahí, clic derecho sobre cualquier archivo
soportado → **Convertir a Markdown**. (En Windows 11, "Mostrar más
opciones" primero si no aparece.)

> El instalador es **auto-localizable**: deriva la ruta del `.pyw` de su
> propia carpeta, así que debe estar junto a `convertir_a_md.pyw`. Si mueves
> la carpeta, vuelve a ejecutarlo y las rutas del registro se corrigen solas.
> (Se sustituyó el antiguo `.reg`, que tenía rutas fijas y se rompía al mover
> la carpeta.)

Para quitarlo: `registry\desinstalar_menu_contextual.reg`.

### Accesos directos (Escritorio y Menú Inicio)

Ejecuta `registry\crear_accesos_directos.ps1` (clic derecho → "Ejecutar con
PowerShell"). Crea un acceso directo "Convertir a Markdown" en el Escritorio
y en el Menú Inicio. El script se auto-ubica: apunta al `convertir_a_md.pyw`
que esté en su misma carpeta, así que colócalo junto a la app.

### CLI

```powershell
python pdf_office_to_md.py "ruta\a\archivo.pdf"
python pdf_office_to_md.py "ruta\a\carpeta" -r -o "ruta\salida"
python pdf_office_to_md.py archivo.docx --beside
```

Por defecto guarda los `.md` en tu carpeta de Descargas
(`~\Downloads`).

## Atajo opcional en PowerShell

Añade a tu `$PROFILE`:

```powershell
function pdf2md {
    python "C:\Users\<tu_usuario>\Repos\pdf-office-to-md\pdf_office_to_md.py" @args
}
```

## Publicar una nueva versión

Hay un script que automatiza todo el ciclo (bump de versión, sincronización
a OneDrive, commit, push y publicación del release con el `.pyw` adjunto).
El script vive en la carpeta de instalación
(`...\Claude\Convertir a Markdown\release.ps1`) pero opera sobre este repo
(la ruta del repo está fijada dentro del propio script):

```powershell
cd "C:\Users\javie\OneDrive - Notley Green Primary School\Claude\Convertir a Markdown"
.\release.ps1 -Version "1.3.0" -Notes "Soporte para .odt. Corregido bug X."
```

El script:
1. Actualiza `__version__` en `convertir_a_md.pyw`.
2. Reescribe `version.json` con la nueva versión y las notas.
3. Copia el `.pyw` a tu instalación en OneDrive.
4. Hace `git add`, `commit` y `push`.
5. Crea un release `vX.Y.Z` en GitHub con el `.pyw` como descarga.

La app instalada detecta el nuevo `version.json` en cuanto se reabre.

## CI

Cada push y cada PR a `main` ejecuta `.github/workflows/ci.yml` en GitHub
Actions:

- `py_compile` sobre `convertir_a_md.pyw` y `pdf_office_to_md.py` (chequeo
  de sintaxis).
- Validación del `version.json` (semver y URL).
- Aviso si `__version__` y `version.json` no coinciden.

## Licencia

MIT — ver [LICENSE](LICENSE).
