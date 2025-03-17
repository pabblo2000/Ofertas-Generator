# run_app.ps1

$OutputEncoding = [System.Text.Encoding]::UTF8

# Verificar si existe el entorno virtual (.venv)
if (-not (Test-Path ".venv")) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("No se encontro el entorno virtual. Por favor, ejecute 'installer.bat' primero.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}

# Activar el entorno virtual
& .\.venv\Scripts\Activate.ps1

# Ejecutar la aplicación Streamlit
python -m streamlit run app.py
