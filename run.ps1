$OutputEncoding = [System.Text.Encoding]::UTF8

# Verificar si existe el entorno virtual (.venv)
if (-not (Test-Path ".venv")) {
    Write-Host "Entorno virtual no encontrado. Ejecutando installer.bat..."
    Start-Process -Wait -FilePath ".\installer.bat"
}

# Si config.py no existe, solicitar datos y crearlo
if (-not (Test-Path "config.py")) {
    Write-Host "config.py no encontrado. Por favor, ingrese los siguientes datos para la configuracion."
    $nombre = Read-Host "Ingrese su nombre"
    $configContent = @"
    
correo_proveedor = ""
default_template = r".\plantilla.docx"
output_folder = r""
nombre = "$nombre"
selected_docs = ["Word", "PDF"]
enable_advanced_date_fields = False
enable_custom_fields = False
enable_description = True
enable_alcance = True

"@

    Set-Content -Path "config.py" -Value $configContent -Encoding UTF8
    Write-Host "Archivo config.py creado con exito."
    Write-Host @"
    
  _______                                                               _ _ 
 |__   __|                                                             |_| |
    | |_   _ _ __   ___   _   _  ___  _   _ _ __    ___ _ __ ___   __ _ _| |
    | | | | | '_ \ / _ \ | | | |/ _ \| | | | '__|  / _ \ '_ \ _ \ / _\ | | |
    | | |_| | |_) |  __/ | |_| | (_) | |_| | |    |  __/ | | | | | (_| | | |
    |_|\__, | .__/ \___|  \__, |\___/ \__,_|_|     \___|_| |_| |_|\__,_|_|_|
        __/ | |            __/ |                                            
       |___/|_|           |___/                                             

"@


}


Write-Host "Activando el entorno virtual..."
& .\.venv\Scripts\Activate.ps1

Write-Host "Ejecutando la aplicacion Streamlit..."
python -m streamlit run app.py

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
