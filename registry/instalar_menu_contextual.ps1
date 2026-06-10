# instalar_menu_contextual.ps1
# Anade (o repara) la entrada "Convertir a Markdown" en el menu contextual
# del Explorador para PDF y archivos de Office.
#
# AUTO-LOCALIZABLE: usa la ruta de ESTE script para encontrar
# convertir_a_md.pyw (debe estar en la misma carpeta). Si mueves la carpeta,
# vuelve a ejecutar este archivo y las rutas se corrigen solas.
#
# Uso: clic derecho -> "Ejecutar con PowerShell".
# Si Windows lo bloquea:
#   powershell -ExecutionPolicy Bypass -File "instalar_menu_contextual.ps1"
#
# Se instala SOLO para el usuario actual (HKCU); no requiere admin.

$ErrorActionPreference = "Stop"

$scriptDir = $PSScriptRoot
$appPath   = Join-Path $scriptDir "convertir_a_md.pyw"

# Localizar pythonw.exe (sin consola)
$pythonw = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source
if (-not $pythonw) {
    $pythonw = "C:\Users\javie\AppData\Local\Programs\Python\Python312\pythonw.exe"
}

if (-not (Test-Path $appPath)) { Write-Error "No encuentro la app en: $appPath"; exit 1 }
if (-not (Test-Path $pythonw)) { Write-Error "No encuentro pythonw.exe en: $pythonw"; exit 1 }

# Extensiones soportadas para el menu contextual
$exts = @(".pdf",".docx",".doc",".xlsx",".xls",".pptx",".ppt",
          ".csv",".epub",".html",".txt",".rtf")

# El comando que ejecuta Windows al pulsar la entrada
$command = '"{0}" "{1}" "%1"' -f $pythonw, $appPath

foreach ($ext in $exts) {
    $base = "HKCU:\Software\Classes\SystemFileAssociations\$ext\Shell\ConvertirAMd"
    $cmd  = "$base\command"
    New-Item -Path $base -Force | Out-Null
    New-ItemProperty -Path $base -Name "(default)" -Value "Convertir a Markdown" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $base -Name "Icon" -Value $pythonw -PropertyType String -Force | Out-Null
    # Position=Top: coloca la entrada arriba del menu clasico, antes de 'Enviar a'.
    New-ItemProperty -Path $base -Name "Position" -Value "Top" -PropertyType String -Force | Out-Null
    New-Item -Path $cmd -Force | Out-Null
    New-ItemProperty -Path $cmd -Name "(default)" -Value $command -PropertyType String -Force | Out-Null
    Write-Host "OK  $ext"
}

Write-Host ""
Write-Host "Instalado. Apunta a:"
Write-Host "  $appPath"
Write-Host "Clic derecho sobre un archivo soportado -> 'Convertir a Markdown'."
Write-Host "(En Windows 11, pulsa antes 'Mostrar mas opciones'.)"
