#Script que activa la politica de ejecución de scripts e instala librerias de python necesarias para
# el script python

Set-ExecutionPolicy RemoteSigned -Scope Process -Force

pip install python-docx -q
Write-Host "Instalando libreria python-docx..."
pip install openpyxl -q
Write-Host "Instalando libreria openpyxl..."