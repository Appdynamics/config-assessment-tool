[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]] $CatArgs
)

$ErrorActionPreference = 'Stop'
$RootDir = $PSScriptRoot
$VenvDir = if ($env:CAT_VENV_DIR) { $env:CAT_VENV_DIR } else { Join-Path $RootDir '.venv' }
$Python = Join-Path $VenvDir 'Scripts/python.exe'
$Pipenv = Join-Path $VenvDir 'Scripts/pipenv.exe'

function Invoke-Native {
    param(
        [Parameter(Mandatory = $true)][string] $Command,
        [Parameter(Mandatory = $true)][string[]] $Arguments
    )

    & $Command @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code ${LASTEXITCODE}: $Command $($Arguments -join ' ')"
    }
}

Set-Location $RootDir

$PythonVersion = $null
if (Test-Path -LiteralPath $Python) {
    $PythonVersion = (& $Python -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")').Trim()
    if ($LASTEXITCODE -ne 0) { $PythonVersion = $null }
}

if ($PythonVersion -ne '3.12') {
    $VenvParent = Split-Path -Parent $VenvDir
    if (-not (Test-Path -LiteralPath $VenvParent)) {
        New-Item -ItemType Directory -Force -Path $VenvParent | Out-Null
    }

    $Uv = Get-Command uv -ErrorAction SilentlyContinue
    if ($Uv) {
        Invoke-Native -Command $Uv.Source -Arguments @('venv', '--clear', '--python', '3.12', $VenvDir)
    } else {
        $PyLauncher = Get-Command py -ErrorAction SilentlyContinue
        if (-not $PyLauncher) {
            throw 'Python 3.12 is required. Install uv or the Python 3.12 Windows launcher, then rerun this script.'
        }
        Invoke-Native -Command $PyLauncher.Source -Arguments @('-3.12', '-m', 'venv', '--clear', $VenvDir)
    }
}

if (-not (Test-Path -LiteralPath $Python)) {
    throw "Virtual environment Python was not created at $Python"
}

if (-not (Test-Path -LiteralPath $Pipenv)) {
    $Uv = Get-Command uv -ErrorAction SilentlyContinue
    if ($Uv) {
        Invoke-Native -Command $Uv.Source -Arguments @('pip', 'install', '--python', $Python, 'pipenv')
    } else {
        Invoke-Native -Command $Python -Arguments @('-m', 'pip', 'install', 'pipenv')
    }
}

$env:PIPENV_PYTHON = $Python
$env:PIPENV_VENV_IN_PROJECT = '1'
$env:PIPENV_IGNORE_VIRTUALENVS = '1'
Invoke-Native -Command $Pipenv -Arguments @('install', '--deploy', '--system')

@('input/jobs', 'input/thresholds', 'logs', 'output') | ForEach-Object {
    New-Item -ItemType Directory -Force -Path (Join-Path $RootDir $_) | Out-Null
}

$BackendDir = Join-Path $RootDir 'backend'
$env:PYTHONPATH = if ($env:PYTHONPATH) {
    "$RootDir;$BackendDir;$($env:PYTHONPATH)"
} else {
    "$RootDir;$BackendDir"
}

if ($CatArgs.Count -eq 0 -or $CatArgs[0] -in @('--ui', '--run')) {
    $Port = if ($env:CAT_PORT) { $env:CAT_PORT } else { '8501' }
    $CatHost = if ($env:CAT_HOST) { $env:CAT_HOST } else { '127.0.0.1' }
    Write-Host "Starting CAT UI at http://${CatHost}:$Port"
    $StreamlitArgs = @('-m', 'streamlit', 'run', (Join-Path $RootDir 'frontend/frontend.py'), '--server.address', $CatHost, '--server.port', $Port, '--server.headless', 'true')
    Invoke-Native -Command $Python -Arguments $StreamlitArgs
} else {
    $CliArgs = @((Join-Path $RootDir 'bin/bundle_main.py')) + $CatArgs
    Invoke-Native -Command $Python -Arguments $CliArgs
}
