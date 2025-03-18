$OutputEncoding = [System.Text.Encoding]::UTF8

# Si config.py no existe, solicitar datos y crearlo
if (-not (Test-Path "config.py")) {
    Write-Host "config.py no encontrado. Por favor, ingrese los siguientes datos para la configuracian."
    $correo = Read-Host "Ingrese el correo proveedor"
    $modo_guardado = "Mediante descarga"  # Valor por defecto; se puede cambiar luego
    $default_template = Read-Host "Ingrese la ruta de la plantilla por defecto (ej: C:\ruta\plantilla.docx)"
    $output_folder = Read-Host "Ingrese la ruta de salida (opcional)"
    $nombre = Read-Host "Ingrese su nombre"
    $configContent = @"
correo_proveedor = "$correo"
modo_guardado = "$modo_guardado"
default_template = r"$default_template"
output_folder = r"$output_folder"
nombre = "$nombre"
"@
    Set-Content -Path "config.py" -Value $configContent -Encoding UTF8
    Write-Host "Archivo config.py creado con exito. Por favor, reinicie la aplicacian."
    Pause
    exit 0
}

# Verificar si existe el entorno virtual (.venv)
if (-not (Test-Path ".venv")) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("No se encontra el entorno virtual. Por favor, ejecute installer.bat primero.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}

Write-Host "Activando el entorno virtual..."
& .\.venv\Scripts\Activate.ps1

Write-Host "Ejecutando la aplicacian Streamlit..."
python -m streamlit run app.py
