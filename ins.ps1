$desiredVersion = "3.11.9"

# Intentar obtener la versión de Python 3.11 usando el lanzador
try {
    $versionOutput = & py -3.11 --version 2>&1
    $installedVersion = ($versionOutput -replace 'Python ', '').Trim()
} catch {
    $installedVersion = ""
}

if ($installedVersion -ne $desiredVersion) {
    Write-Host "La versión instalada de Python ($installedVersion) no coincide con la requerida ($desiredVersion) o no se encontró."
    Write-Host "Procediendo a descargar e instalar Python $desiredVersion..."
    
    $pythonInstallerUrl = "https://www.python.org/ftp/python/$desiredVersion/python-$desiredVersion-amd64.exe"
    $installerPath = "$env:TEMP\python-$desiredVersion-amd64.exe"
    Invoke-WebRequest -Uri $pythonInstallerUrl -OutFile $installerPath
    Write-Host "Instalando Python $desiredVersion..."
    Start-Process -FilePath $installerPath -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait
    Remove-Item $installerPath
    
    # Revalidar la instalación usando el lanzador
    try {
        $versionOutput = & py -3.11 --version 2>&1
        $installedVersion = ($versionOutput -replace 'Python ', '').Trim()
    } catch {
        $installedVersion = ""
    }
    
    if ($installedVersion -ne $desiredVersion) {
        Write-Host "ERROR: No se pudo instalar Python $desiredVersion. Por favor, instale Python manualmente."
        exit 1
    } else {
        Write-Host "Python $desiredVersion instalado correctamente."
    }
} else {
    Write-Host "Python $desiredVersion ya está instalado."
}

# Verificar si existe el entorno virtual (.venv) y crearlo si no existe, usando Python 3.11
if (-not (Test-Path ".venv")) {
    Write-Host "Entorno virtual no encontrado. Creando .venv..."
    py -3.11 -m venv .venv
    if (-not (Test-Path ".venv")) {
        Write-Host "ERROR: No se pudo crear el entorno virtual."
        exit 1
    }
} else {
    Write-Host ".venv ya existe. Saltando creación."
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
  \   / _ \| | | | | |\/| |/ _\ | | | | | |    | |/ _ \/ __|/ _ \    | |  | '_ \| / __|    | |/ _` | '_ \ 
   | | (_) | |_| | | |  | | (_| | |_| | | |____| | (_) \__ \  __/    | |  | | | | \__ \    | | (_| | |_) |
   |_|\___/ \__,_| |_|  |_|\__,_|\__, |  \_____|_|\___/|___/\___|    |_|  |_| |_|_|___/    |_|\__,_|_.__/ 
                                  __/ |                                                                   
                                 |___/                                                                    

"@
exit 0
