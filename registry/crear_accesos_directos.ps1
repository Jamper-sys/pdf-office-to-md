# crear_accesos_directos.ps1
# Crea accesos directos a "Convertir a Markdown" en el Escritorio
# y en el Menu Inicio del usuario actual.
#
# Uso: clic derecho sobre este archivo -> "Ejecutar con PowerShell".
# Si Windows lo bloquea, abre PowerShell y ejecuta:
#     powershell -ExecutionPolicy Bypass -File "$PWD\crear_accesos_directos.ps1"

$ErrorActionPreference = "Stop"

# La app vive EN LA MISMA CARPETA que este script (auto-ubicacion).
$scriptDir = $PSScriptRoot
$appPath   = Join-Path $scriptDir "convertir_a_md.pyw"

# Localizar pythonw.exe (sin consola). Primero el del PATH; si no, la
# ruta tipica de instalacion de Python para el usuario.
$pythonw = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source
if (-not $pythonw) {
    $pythonw = "C:\Users\javie\AppData\Local\Programs\Python\Python312\pythonw.exe"
}

# Icono: usa el de pythonw.exe (es un icono valido). Si quieres otro,
# cambia esta linea a una ruta .ico tuya, o a "shell32.dll,1" para
# un icono generico de documento.
$iconPath = $pythonw

if (-not (Test-Path $appPath)) {
    Write-Error "No encuentro la app en: $appPath"
    exit 1
}
if (-not (Test-Path $pythonw)) {
    Write-Error "No encuentro pythonw.exe en: $pythonw"
    exit 1
}

$shell = New-Object -ComObject WScript.Shell

function Crear-Shortcut($ruta) {
    $sc = $shell.CreateShortcut($ruta)
    $sc.TargetPath       = $pythonw
    $sc.Arguments        = '"' + $appPath + '"'
    $sc.WorkingDirectory = $scriptDir
    $sc.IconLocation     = $iconPath
    $sc.Description      = "Convierte PDF y archivos de Office a Markdown"
    $sc.Save()
    Write-Host "Creado: $ruta"
}

# Escritorio
$desktop = [Environment]::GetFolderPath("Desktop")
Crear-Shortcut (Join-Path $desktop "Convertir a Markdown.lnk")

# Menu Inicio (carpeta Programs del usuario)
$startMenu = [Environment]::GetFolderPath("Programs")
Crear-Shortcut (Join-Path $startMenu "Convertir a Markdown.lnk")

Write-Host ""
Write-Host "Listo. Busca 'Convertir a Markdown' en el menu Inicio o en el Escritorio."
