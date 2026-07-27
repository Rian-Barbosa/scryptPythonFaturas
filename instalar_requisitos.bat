@echo off
echo ====================================================
echo   Instalador de Requisitos para Automacao de Faturas
echo ====================================================
echo.

rem Verifica se o Python esta configurado e funcional
where python >nul 2>&1
if %errorlevel% neq 0 (
    goto :INSTALL_PYTHON
)

python --version >nul 2>&1
if %errorlevel% neq 0 (
    goto :INSTALL_PYTHON
)

goto :INSTALL_LIBS


:INSTALL_PYTHON
echo [AVISO] Python nao foi encontrado no sistema.
echo Baixando o instalador oficial do Python 3.11.9...
echo Aguarde um instante...
echo.

rem Baixa o instalador usando o curl (nativo no Windows 10/11)
curl -L -o python_installer.exe https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe

if not exist python_installer.exe (
    echo [ERRO] Falha ao baixar o instalador do Python.
    echo Por favor, acesse o site oficial (python.org) e instale manualmente.
    echo ATENCAO: Lembre-se de marcar a opcao "Add python.exe to PATH" na instalacao.
    echo.
    pause
    exit
)

echo [INFO] Instalador baixado!
echo.
echo ==========================================================
echo   ATENCAO: Na janela que vai abrir, certifique-se de:
echo   [x] MARCAR a caixinha "Add python.exe to PATH" (embaixo)
echo   [x] Clicar em "Install Now"
echo ==========================================================
echo.
pause

echo Iniciando o instalador...
start /wait python_installer.exe

rem Limpa o instalador baixado
del python_installer.exe

echo.
echo [INFO] Instalacao concluida!
echo IMPORTANTISSIMO: Como o Python acabou de ser instalado, voce precisa:
echo 1. Fechar esta janela do CMD (clique no X).
echo 2. Abrir novamente o "instalar_requisitos.bat" para instalar as bibliotecas.
echo.
pause
exit


:INSTALL_LIBS
echo [INFO] Python detectado com sucesso!
echo Instalando as bibliotecas necessarias (pymupdf, openpyxl)...
echo.
python -m pip install --upgrade pip
python -m pip install pymupdf openpyxl

echo.
echo ====================================================
echo   Tudo pronto! As dependencias foram instaladas.
echo   Agora voce ja pode rodar o script clicando em
echo   "rodar_automacao.bat".
echo ====================================================
echo.
pause
