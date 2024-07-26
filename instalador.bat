echo off    

chcp 65001
cls
                                         
echo:
echo "                                                                             .:~!!777??777!~^.                 .
echo    ███╗   ███╗ █████╗ ██████╗ ███████╗    ██████╗ ██╗   ██╗                 .7YYYY555555555555Y?^^^               .
echo    ████╗ ████║██╔══██╗██╔══██╗██╔════╝    ██╔══██╗╚██╗ ██╔╝██╗              75Y:  ~55555555555555~              .
echo    ██╔████╔██║███████║██║  ██║█████╗      ██████╔╝ ╚████╔╝ ╚═╝              ?557^^^^?555555555555557              .
echo    ██║╚██╔╝██║██╔══██║██║  ██║██╔══╝      ██╔══██╗  ╚██╔╝  ██╗              !JJYJYJJJJJY5555555557              .
echo    ██║ ╚═╝ ██║██║  ██║██████╔╝███████╗    ██████╔╝   ██║   ╚═╝       :^^^^^^^^^^^^^~~~~~~~~~~~Y5555555557 .:....       .
echo    ╚═╝     ╚═╝╚═╝  ╚═╝╚═════╝ ╚══════╝    ╚═════╝    ╚═╝          .!JY5555555555555555555555555557 ^^^^^^^^^^^^^^^:     .
echo:                                                                 .J5555555555555555555555555555557 ^^^^^^^^^^^^^^^^^^    .
echo        ██████╗      ██╗ █████╗ ██╗     ███╗   ███╗ █████╗        75555555555555555555555555555555^ ^^^^^^^^^^^^^^^^^^^^:   .
echo        ██╔══██╗     ██║██╔══██╗██║     ████╗ ████║██╔══██╗      .Y5555555555555555555555555555Y?^ :^^^^^^^^^^^^^^^^^^^^^^   .
echo        ██║  ██║     ██║███████║██║     ██╔████╔██║███████║      :55555555555Y?!~^^~^~~~~~~^^^:...^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^   .
echo        ██║  ██║██   ██║██╔══██║██║     ██║╚██╔╝██║██╔══██║      .Y555555555?. .::^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^   .
echo        ██████╔╝╚█████╔╝██║  ██║███████╗██║ ╚═╝ ██║██║  ██║       J55555555Y..^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^:   .
echo        ╚═════╝  ╚════╝ ╚═╝  ╚═╝╚══════╝╚═╝     ╚═╝╚═╝  ╚═╝       ^55555555? :^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^.   .
echo:                                                                  ^J555555J :^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^:     .
echo            ███████╗██╗██╗     ██╗  ██╗ ██████╗                     .^7????! :^^^^^^^^^^^^^^^^^^^^^^:::::::::::::::::.       .
echo            ██╔════╝██║██║     ██║  ██║██╔═══██╗                             :^^^^^^^^^^^^^^^^^^^^::::::::::.              .
echo            █████╗  ██║██║     ███████║██║   ██║                             :^^^^^^^^^^^^^^^^^^^^^^^^^^^^^::^^^^^^^^:              .
echo            ██╔══╝  ██║██║     ██╔══██║██║   ██║                             :^^^^^^^^^^^^^^^^^^^^^^^^^^^.  .^^^^^^:              .
echo            ██║     ██║███████╗██║  ██║╚██████╔╝                              :^^^^^^^^^^^^^^^^^^^^^^^^^^^::^^^^^^:.              .
echo            ╚═╝     ╚═╝╚══════╝╚═╝  ╚═╝ ╚═════╝                                 .::::^^^^^^^^^^^^^^:::..                .

timeout /t 4 
cls

echo:
echo:
echo:

echo %TAB%     *** INTALADOR DE PACOTES PYTHON PARA O ARCMAP 10.x ***...
echo %TAB%     .
echo %TAB%     # ELE VERIFICARÁ SE OS PACOTES JÁ ESTÃO INSTALADOS
echo %TAB%     # CASO VOCÊ JÁ TENHA INSTALADO, ELE FINALIZARÁ SEM PROBLEMAS
echo %TAB%     # CASO VOCÊ NÃO TENHA INSTALADO, OS PACOTES SERÃO BAIXADOS E INSTALADOS
echo %TAB%     .
echo %TAB%     # VOCÊ ESTÁ PRESTES A INSTALAR AS EXTENSÕES PARA O ARCGIS...#

echo:
echo:
echo:


echo Caso nenhuma tecla seja apertada o programa continuará a instalação normalmente
timeout /t 15


start "" "C:\Python27\ArcGIS10.8\python.exe" -m pip install lxml==4.6.1 requests==2.25.0 openpyxl==2.6.4 python-docx==0.8.11 
@REM start "" "C:\Python27\ArcGIS10.8\python.exe" -m pip install backports.functools-lru-cache==1.6.4 certifi==2020.12.5 chardet==3.0.4 cycler==0.10.0 et-xmlfile==1.0.1 functools32==3.2.3.post2 future==0.18.2 idna==2.10 jdcal==1.4.1 kiwisolver==1.1.0 lxml==4.9.2 matplotlib==2.2.5 mpmath==1.2.1 nose==1.3.7 numpy==1.16.6 openpyxl==2.6.4 pandas==0.24.2 pyparsing==2.4.7 python-dateutil==2.8.2 python-docx==0.8.11 pytz==2021.1 requests==2.25.0 scipy==1.2.3 setuptools-scm==5.0.2 six==1.16.0 toml==0.10.2 urllib3==1.26.6 xlrd==1.2.0 xlwt==1.3.0

chcp 936

cls
chcp 65001

echo:
echo:
echo:

echo %TAB%     ASSIM QUE A OUTRA JANELA SE FECHAR
echo %TAB%      INSTALAÇÃO CONLUÍDA COM SUCESSO!

echo:
echo:
echo:

timeout /t 999