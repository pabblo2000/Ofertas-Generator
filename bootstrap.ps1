# bootstrap.ps1

# Función para verificar si un comando existe
function Command-Exists {
    param ([string]$cmd)
    $ErrorActionPreference = "SilentlyContinue"
    $null -ne (Get-Command $cmd -ErrorAction SilentlyContinue)
}

Write-Host "Verificando si Python está instalado..."
if (-not (Command-Exists "python")) {
    Write-Host "Python no está instalado. Descargando e instalando Python..."
    # Ajusta la URL a la versión de Python que necesites (este ejemplo usa Python 3.9.7 64-bit)
    $installerUrl = "https://www.python.org/ftp/python/3.9.7/python-3.9.7-amd64.exe"
    $installerPath = "$env:TEMP\python-installer.exe"
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
    Write-Host "Instalando Python..."
    # Ejecuta el instalador en modo silencioso. Se instalará para todos los usuarios y se agregará al PATH.
    Start-Process -FilePath $installerPath -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait
    Remove-Item $installerPath
    # Verificar nuevamente
    if (-not (Command-Exists "python")) {
        Write-Error "La instalación de Python falló. Por favor, instálalo manualmente."
        exit 1
    }
}

$pythonVersion = python --version 2>&1
Write-Host "Python instalado: $pythonVersion"

# Actualizar pip e instalar dependencias
Write-Host "Instalando dependencias..."
python -m pip install --upgrade pip
if (-Not (Test-Path "requirements.txt")) {
    Write-Error "No se encontró requirements.txt. Asegúrate de tenerlo en la carpeta del proyecto."
    exit 1
}
python -m pip install -r requirements.txt

# Ejecutar la aplicación Streamlit
Write-Host "Ejecutando la aplicación Streamlit..."
python -m streamlit run app.py
