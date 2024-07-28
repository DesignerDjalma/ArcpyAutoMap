@echo off

rem Mostra a mensagem inicial
call :showMsg "Atualizacao de Python" "O Python sera atualizado e modulos/bibliotecas Python serao adicionados para correto funcionamento da ferramenta. O processo na primeira atualizacao pode demorar alguns minutos. Aguarde enquanto os pacotes sao atualizados e adicionados."

rem Atualizando pip
echo Atualizando pip...
"C:\Python27\ArcGIS10.8\python.exe" -m pip install --upgrade pip
if %errorlevel% neq 0 (
    call :showMsg "Erro" "Falha ao atualizar pip."
    exit /b %errorlevel%
)

rem Instalando pacotes
echo Instalando pacotes...
"C:\Python27\ArcGIS10.8\python.exe" -m pip install PyYAML==5.4.1 psycopg2==2.8.6 openpyxl==2.6.4 lxml==4.6.1 python-docx==0.8.11
if %errorlevel% neq 0 (
    call :showMsg "Erro" "Falha ao adicionar pacotes."
    exit /b %errorlevel%
)

rem Mostra a mensagem de conclusão
call :showMsg "Sucesso" "Pacotes adicionados com sucesso ao Python 2.7!"

exit /b

:showMsg
setlocal
set "title=%~1"
set "msg=%~2"
mshta "javascript:var sh=new ActiveXObject('WScript.Shell'); sh.Popup('%msg%', 10, '%title%', 64);close();"
endlocal
exit /b
