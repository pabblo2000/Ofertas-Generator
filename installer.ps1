# installer.ps1

$OutputEncoding = [System.Text.Encoding]::UTF8
Write-Host "Iniciando instalacion..."

if (-not (Test-Path ".venv")) {
    Write-Host "Entorno virtual no encontrado. Creando .venv..."
    python -m venv .venv
    if (-not (Test-Path ".venv")) {
        Write-Host "ERROR: No se pudo crear el entorno virtual."
        exit 1
    }
} else {
    Write-Host ".venv ya existe. Saltando creacion."
}

Write-Host "Activando el entorno virtual..."
& .\.venv\Scripts\Activate.ps1

Write-Host "Actualizando pip..."
python -m pip install --upgrade pip

if (-not (Test-Path "requirements.txt")) {
    Write-Host "ERROR: requirements.txt no encontrado."
    exit 1
}

Write-Host "Instalando dependencias desde requirements.txt..."
python -m pip install -r requirements.txt

Write-Host "*******************************************************"
Write-Host "**** APLICACION INSTALADA ****"
Write-Host "*******************************************************"
Write-Host "**** EJECUTE run_app.vbs PARA INICIAR LA APLICACION****"
Write-Host "*******************************************************"
Write-Host "**** PUEDE CERRAR ESTA VENTANA ****"
Write-Host "*******************************************************"
Pause
