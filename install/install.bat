@echo off
echo Atualizando pip...
"C:\Python27\ArcGIS10.8\python.exe" -m pip install --upgrade pip
if %errorlevel% neq 0 (
    echo Falha ao atualizar pip.
    exit /b %errorlevel%
)

echo Instalando pacotes...
"C:\Python27\ArcGIS10.8\python.exe" -m pip install PyYAML==5.4.1 psycopg2==2.8.6 openpyxl==2.6.4 lxml==4.6.1 python-docx==0.8.11
if %errorlevel% neq 0 (
    echo Falha ao instalar pacotes.
    exit /b %errorlevel%
)

echo Instalacao concluida com sucesso!