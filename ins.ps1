$desiredVersion = "3.11.9"

# Intentar obtener la version de Python 3.11 usando el lanzador
try {
    $versionOutput = & py -3.11 --version 2>&1
    $installedVersion = ($versionOutput -replace 'Python ', '').Trim()
} catch {
    $installedVersion = ""
}

if ($installedVersion -ne $desiredVersion) {
    Write-Host "La version instalada de Python ($installedVersion) no coincide con la requerida ($desiredVersion) o no se encontro."
    Write-Host "Procediendo a instalar Python $desiredVersion..."
    
    # Definir la ruta del instalador local (se asume que esta en la raiz del proyecto)
    $localInstallerPath = Join-Path -Path (Get-Location) -ChildPath "python-$desiredVersion-amd64.exe"
    
    if (Test-Path $localInstallerPath) {
        Write-Host "Instalador local encontrado. Ejecutando $localInstallerPath..., puede tomar unos minutos."
        Start-Process -FilePath $localInstallerPath -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait
    } else {
        Write-Host "Instalador local no encontrado. Procediendo a descargar..., puede tomar unos minutos."
        $pythonInstallerUrl = "https://www.python.org/ftp/python/$desiredVersion/python-$desiredVersion-amd64.exe"
        $installerPath = "$env:TEMP\python-$desiredVersion-amd64.exe"
        Invoke-WebRequest -Uri $pythonInstallerUrl -OutFile $installerPath
        Start-Process -FilePath $installerPath -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait
        Remove-Item $installerPath
    }
    
    # Revalidar la instalacion usando el lanzador
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
    Write-Host "Python $desiredVersion ya esta instalado."
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

Write-Host "Instalando dependencias desde requirements.txt..., puede tomar unos minutos."
$requirements = Get-Content "requirements.txt" | Where-Object { $_.Trim() -and -not ($_.Trim().StartsWith("#")) }
$total = $requirements.Count
$i = 0

foreach ($package in $requirements) {
    $i++
    Write-Progress -Activity "Instalando dependencias" -Status "Instalando $package ($i de $total)" -PercentComplete (($i / $total) * 100)
    python -m pip install $package
}


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
