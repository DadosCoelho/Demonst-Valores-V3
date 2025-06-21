#!/bin/bash

# Define a codificação para UTF-8 (útil para saída no terminal)
export LANG=en_US.UTF-8
export LC_ALL=en_US.UTF-8

echo ""
echo "========================================================="
echo " Iniciando o processamento da Demonstracao Financeira"
echo " (com Ambiente Virtual)"
echo "========================================================="
echo ""

# Navega para o diretorio onde o script .sh esta localizado
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
cd "$SCRIPT_DIR"

# Função para desativar o ambiente virtual e esperar por uma tecla
exit_with_pause() {
    echo ""
    echo "Desativando ambiente virtual..."
    # Verifica se a função deactivate existe e a chama (se o venv foi ativado)
    if command -v deactivate &>/dev/null; then
        deactivate
    fi
    echo "Pressione qualquer tecla para fechar o terminal..."
    read -n 1 -s -r # Lê uma tecla sem eco e sem precisar de Enter
    exit "$1" # Sai com o código de status fornecido
}

# --- Verifica qual comando Python/pip usar ---
if command -v python3 &>/dev/null; then
    PYTHON_MAIN_CMD="python3"
elif command -v python &>/dev/null; then
    PYTHON_MAIN_CMD="python"
else
    echo "ERRO: Python nao encontrado. Certifique-se de que o Python esta instalado e no PATH."
    exit_with_pause 1 # Sai com erro
fi

# --- Verifica e cria o ambiente virtual ---
echo "Verificando ambiente virtual..."
if [ ! -d "venv" ]; then
    echo "Ambiente virtual 'venv' nao encontrado. Criando..."
    "$PYTHON_MAIN_CMD" -m venv venv
    if [ $? -ne 0 ]; then
        echo "ERRO: Falha ao criar o ambiente virtual."
        echo "Verifique se o Python esta instalado e acessivel."
        exit_with_pause 1 # Sai com erro
    fi
    echo "Ambiente virtual criado com sucesso."
else
    echo "Ambiente virtual 'venv' ja existe."
fi

# --- Ativa o ambiente virtual ---
echo "Ativando ambiente virtual..."
source venv/bin/activate
if [ $? -ne 0 ]; then
    echo "ERRO: Falha ao ativar o ambiente virtual."
    exit_with_pause 1 # Sai com erro
fi
echo "Ambiente virtual ativado."
echo ""

# --- Instala dependencias dentro do ambiente virtual ---
echo "Instalando/Verificando dependencias no ambiente virtual (pandas, openpyxl)..."
pip install pandas openpyxl
if [ $? -ne 0 ]; then
    echo "ERRO: Falha ao instalar dependencias."
    echo "Certifique-se de que o pip esta funcionando corretamente dentro do ambiente virtual."
    exit_with_pause 1 # Sai com erro
fi
echo "Dependencias verificadas/instaladas."
echo ""

# --- Executa o script Python ---
echo "Executando o script processar_dados.py..."
python processar_dados.py
if [ $? -ne 0 ]; then
    echo "ERRO: O script Python encontrou um problema durante a execucao."
    echo "Verifique as mensagens de erro acima para detalhes."
    exit_with_pause 1 # Sai com erro
fi
echo ""
echo "Processamento concluido!"
echo "O arquivo index.html foi atualizado com os dados calculados."
echo ""
echo "Para visualizar os resultados, abra o arquivo index.html em seu navegador"
echo "ou inicie um servidor web local (se configurado)."
echo ""

# --- Finaliza com a pausa ---
exit_with_pause 0 # Sai com sucesso