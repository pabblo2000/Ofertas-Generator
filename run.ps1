$OutputEncoding = [System.Text.Encoding]::UTF8

# Si config.py no existe, solicitar datos y crearlo
if (-not (Test-Path "config.py")) {
    Write-Host "config.py no encontrado. Por favor, ingrese los siguientes datos para la configuracion."
    $modo_guardado = "Mediante descarga"  
    $nombre = Read-Host "Ingrese su nombre"
    $configContent = @"
    
correo_proveedor = "correo"
modo_guardado = "$modo_guardado"
default_template = r".\plantilla.docx"
output_folder = r""
nombre = "$nombre"
"@
    Set-Content -Path "config.py" -Value $configContent -Encoding UTF8
    Write-Host "Archivo config.py creado con exito."
    Write-Host @"

    _____ _                  _   _     _       _        _     
    / ____| |                | | | |   (_)     | |      | |    
   | |    | | ___  ___  ___  | |_| |__  _ ___  | |_ __ _| |__  
   | |    | |/ _ \/ __|/ _ \ | __| '_ \| / __| | __/ _` | '_ \ 
   | |____| | (_) \__ \  __/ | |_| | | | \__ \ | || (_| | |_) |
    \_____|_|\___/|___/\___|  \__|_| |_|_|___/  \__\__,_|_.__/ 
                                                                                                                                                                                                                                                                                                                                                         
"@

}

# Verificar si existe el entorno virtual (.venv)
if (-not (Test-Path ".venv")) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("No se encontro el entorno virtual. Por favor, ejecute installer.bat primero.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}

Write-Host "Activando el entorno virtual..."
& .\.venv\Scripts\Activate.ps1

Write-Host "Ejecutando la aplicacion Streamlit..."
python -m streamlit run app.py
