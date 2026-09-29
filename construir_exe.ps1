$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$proyecto = $PSScriptRoot
$aplicacion = Join-Path $proyecto 'app.py'
$transformador = Join-Path $proyecto 'transformador.py'

if (-not (Test-Path -LiteralPath $aplicacion -PathType Leaf) -or
    -not (Test-Path -LiteralPath $transformador -PathType Leaf)) {
    throw 'Se necesitan app.py y transformador.py en esta carpeta.'
}

$python = Get-Command python -ErrorAction SilentlyContinue
if ($python) {
    $ejecutable = $python.Source
    $argumentosPython = @()
} else {
    $lanzador = Get-Command py -ErrorAction SilentlyContinue
    if (-not $lanzador) { throw 'Se necesita Python 3 para crear el ejecutable.' }
    $ejecutable = $lanzador.Source
    $argumentosPython = @('-3')
}

& $ejecutable @argumentosPython -c 'import PyInstaller, openpyxl'
if ($LASTEXITCODE -ne 0) {
    throw 'PyInstaller y openpyxl deben estar disponibles en el Python usado para compilar.'
}

& $ejecutable @argumentosPython -m PyInstaller `
    --noconfirm `
    --clean `
    --onedir `
    --windowed `
    --noupx `
    --hidden-import transformador `
    --name TransformarEstudiantes `
    --distpath (Join-Path $proyecto 'dist') `
    --workpath (Join-Path $proyecto 'build') `
    --specpath $proyecto `
    $aplicacion

if ($LASTEXITCODE -ne 0) { throw 'PyInstaller no pudo crear el ejecutable.' }
Write-Host "Ejecutable creado en: $(Join-Path $proyecto 'dist\TransformarEstudiantes\TransformarEstudiantes.exe')"
