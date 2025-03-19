$OutputEncoding = [System.Text.Encoding]::UTF8
Write-Host "Iniciando instalacion..."

# Comprobar si Python esta instalado
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "Python no se encuentra instalado. Se procedera a descargar e instalar Python 3.11.9..."
    $pythonInstallerUrl = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
    $installerPath = "$env:TEMP\python-3.11.9-amd64.exe"
    Invoke-WebRequest -Uri $pythonInstallerUrl -OutFile $installerPath
    Write-Host "Instalando Python 3.11.9..."
    Start-Process -FilePath $installerPath -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait
    Remove-Item $installerPath
    if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
        Write-Host "ERROR: No se pudo instalar Python 3.11.9. Por favor, instale Python manualmente."
        exit 1
    } else {
        Write-Host "Python 3.11.9 instalado correctamente."
    }
} else {
    Write-Host "Python ya esta instalado."
}

# Verificar si existe el entorno virtual (.venv) y crearlo si no existe
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
Write-Host "**** EJECUTE run_app.vbs PARA INICIAR LA APLICACION ****"
Write-Host "*******************************************************"
Write-Host @"
__     __           __  __                _____ _                  _______ _     _       _______    _     
\ \   / /          |  \/  |              / ____| |                |__   __| |   (_)     |__   __|  | |    
 \ \_/ /__  _   _  | \  / | __ _ _   _  | |    | | ___  ___  ___     | |  | |__  _ ___     | | __ _| |__  
  \   / _ \| | | | | |\/| |/ _` | | | | | |    | |/ _ \/ __|/ _ \    | |  | '_ \| / __|    | |/ _` | '_ \ 
   | | (_) | |_| | | |  | | (_| | |_| | | |____| | (_) \__ \  __/    | |  | | | | \__ \    | | (_| | |_) |
   |_|\___/ \__,_| |_|  |_|\__,_|\__, |  \_____|_|\___/|___/\___|    |_|  |_| |_|_|___/    |_|\__,_|_.__/ 
                                  __/ |                                                                   
                                 |___/                                                                    

"@
exit 0
