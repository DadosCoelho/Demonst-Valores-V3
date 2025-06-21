@echo off
chcp 65001 > nul

echo.
echo =========================================================
echo  Iniciando o processamento da Demonstracao Financeira
echo  (com Ambiente Virtual)
echo =========================================================
echo.

rem Navega para o diretorio onde o script .bat esta localizado
cd /d "%~dp0"

rem --- Verifica e cria o ambiente virtual ---
echo Verificando ambiente virtual...

rem Tenta encontrar o Python principal
for /f "delims=" %%i in ('where python 2^>nul') do set PYTHON_MAIN_CMD="%%i"
if not defined PYTHON_MAIN_CMD (
    echo ERRO: Python nao encontrado no PATH. Por favor, instale o Python e configure-o no PATH.
    goto :end_with_pause
)

if not exist venv\ (
    echo Ambiente virtual "venv" nao encontrado. Criando...
    %PYTHON_MAIN_CMD% -m venv venv
    if %errorlevel% neq 0 (
        echo ERRO: Falha ao criar o ambiente virtual.
        echo Verifique se o Python esta instalado e acessivel.
        goto :end_with_pause
    )
    echo Ambiente virtual criado com sucesso.
) else (
    echo Ambiente virtual "venv" ja existe.
)

rem --- Ativa o ambiente virtual ---
echo Ativando ambiente virtual...
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo ERRO: Falha ao ativar o ambiente virtual.
    goto :deactivate_and_end_with_pause
)
echo Ambiente virtual ativado.
echo.

rem --- Instala dependencias dentro do ambiente virtual ---
echo Instalando/Verificando dependencias no ambiente virtual (pandas, openpyxl)...
pip install pandas openpyxl --quiet
if %errorlevel% neq 0 (
    echo ERRO: Falha ao instalar dependencias.
    echo Certifique-se de que o pip esta funcionando corretamente dentro do ambiente virtual.
    goto :deactivate_and_end_with_pause
)
echo Dependencias verificadas/instaladas.
echo.

rem --- Executa o script Python ---
echo Executando o script processar_dados.py...
python processar_dados.py
if %errorlevel% neq 0 (
    echo ERRO: O script Python encontrou um problema durante a execucao.
    echo Verifique as mensagens de erro acima para detalhes.
    goto :deactivate_and_end_with_pause
)
echo.
echo Processamento concluido!
echo O arquivo index.html foi atualizado com os dados calculados.
echo.
echo Para visualizar os resultados, abra o arquivo index.html em seu navegador
echo ou inicie um servidor web local (se configurado).
echo.

:deactivate_and_end_with_pause
rem Desativa o ambiente virtual
echo Desativando ambiente virtual...
deactivate

:end_with_pause
echo.
echo Pressione qualquer tecla para fechar o terminal...
pause