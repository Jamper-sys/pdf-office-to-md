<#
.SYNOPSIS
    Publica una nueva versión de pdf-office-to-md.

.DESCRIPTION
    En un solo comando:
      1. Actualiza __version__ en convertir_a_md.pyw.
      2. Reescribe version.json con la nueva versión y las notas.
      3. Copia el .pyw a la carpeta de OneDrive (instalación local de uso diario).
      4. Hace commit y push del cambio a la rama main.
      5. Crea un release v<Version> en GitHub con el .pyw adjunto.

.PARAMETER Version
    Número de versión semántica (ej. "1.3.0"). Obligatorio.

.PARAMETER Notes
    Notas de la versión. Una línea o varias separadas por `n. Obligatorio.

.EXAMPLE
    .\release.ps1 -Version "1.3.0" -Notes "Soporte para .odt. Corregido bug en post-procesado."
#>

param(
    [Parameter(Mandatory=$true)][string]$Version,
    [Parameter(Mandatory=$true)][string]$Notes
)

$ErrorActionPreference = "Stop"

# --------------------------------------------------------------------- rutas

$repo     = $PSScriptRoot
$pyw      = Join-Path $repo "convertir_a_md.pyw"
$verjson  = Join-Path $repo "version.json"
$onedrive = "C:\Users\javie\OneDrive - Notley Green Primary School\Claude\Convertir a Markdown\convertir_a_md.pyw"

# --------------------------------------------------------------- localizar gh

$gh = (Get-Command gh.exe -ErrorAction SilentlyContinue).Source
if (-not $gh) { $gh = "C:\Program Files\GitHub CLI\gh.exe" }
if (-not $gh -or -not (Test-Path $gh)) {
    throw "No se encontro gh CLI. Instalalo con: winget install GitHub.cli"
}

# --------------------------------------------------------------- comprobaciones

if (-not (Test-Path $pyw))     { throw "No existe $pyw" }
if (-not (Test-Path $verjson)) { throw "No existe $verjson" }

Push-Location $repo
try {
    if ((git status --porcelain) -ne $null) {
        Write-Warning "Hay cambios sin commitear en el repo. Resuelvelos antes de release."
        git status
        throw "Working tree no limpio."
    }
} finally {
    Pop-Location
}

# --------------------------------------------------- escribir sin BOM (UTF-8)

function Write-Utf8NoBom([string]$path, [string]$text) {
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($path, $text, $enc)
}

# ----------------------------------- 1) Actualizar __version__ en el .pyw

Write-Host "1/5  Actualizando __version__ -> $Version"
$content = [System.IO.File]::ReadAllText($pyw, [System.Text.UTF8Encoding]::new($false))
$new = [System.Text.RegularExpressions.Regex]::Replace(
    $content,
    '(?m)^__version__\s*=\s*"[^"]*"',
    '__version__ = "' + $Version + '"'
)
if ($new -eq $content) { throw "No se encontro la linea __version__ en $pyw" }
Write-Utf8NoBom $pyw $new

# ------------------------------------- 2) Reescribir version.json

Write-Host "2/5  Actualizando version.json"
$payload = [ordered]@{
    version = $Version
    url     = "https://github.com/Jamper-sys/pdf-office-to-md/releases/latest"
    notes   = $Notes
}
$json = $payload | ConvertTo-Json -Depth 3
Write-Utf8NoBom $verjson ($json + "`n")

# --------------------------------------------- 3) Copiar a OneDrive

Write-Host "3/5  Sincronizando .pyw a OneDrive"
Copy-Item $pyw $onedrive -Force

# ----------------------------------------- 4) Commit y push

Write-Host "4/5  Commit y push"
Push-Location $repo
try {
    git add convertir_a_md.pyw version.json
    git commit -m "v${Version}: ${Notes}"
    if ($LASTEXITCODE -ne 0) { throw "git commit fallo" }
    git push
    if ($LASTEXITCODE -ne 0) { throw "git push fallo" }
} finally {
    Pop-Location
}

# ----------------------------------------- 5) Crear release en GitHub

Write-Host "5/5  Creando release en GitHub"
$tag = "v$Version"
Push-Location $repo
try {
    & $gh release create $tag $pyw --title $tag --notes $Notes
    if ($LASTEXITCODE -ne 0) { throw "gh release create fallo" }
} finally {
    Pop-Location
}

Write-Host ""
Write-Host "OK  Release $tag publicado correctamente."
Write-Host "    https://github.com/Jamper-sys/pdf-office-to-md/releases/tag/$tag"
