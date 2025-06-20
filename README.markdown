# Demonstração Financeira Calculada

## Descrição
O projeto **Demonstração Financeira Calculada** é uma aplicação que processa dados financeiros de planilhas Excel para gerar visões financeiras calculadas, exibidas em uma interface web. Ele lê dados brutos e planos de contas de um arquivo Excel, realiza cálculos hierárquicos e baseados em fórmulas, e atualiza um arquivo HTML com os resultados em formato JSON para visualização.

## Funcionalidades
- **Leitura de Planilhas Excel**: Processa abas de dados brutos (prefixo "dados") e planos de contas (prefixo "plano") de um arquivo Excel (`DADOS.xlsx`).
- **Hierarquia de Contas**: Suporta contas analíticas, sintéticas e de cálculo, organizadas em uma estrutura hierárquica baseada em códigos (ex.: `1`, `1.1`, `1.1.1`).
- **Cálculos Dinâmicos**:
  - Contas analíticas: Soma valores de códigos brutos vinculados a partir de um DataSource.
  - Contas sintéticas: Soma valores de contas filhas.
  - Contas de cálculo: Avalia fórmulas matemáticas (ex.: `001 - 002`) usando valores de outras contas.
- **Geração de Visões**: Cria visões financeiras combinando cada plano de contas com seus DataSources associados, incluindo valores calculados para todos os períodos.
- **Integração Web**: Atualiza um arquivo `index.html` com os dados calculados em formato JSON, para exibição em uma interface web interativa.
- **Cache de Cálculo**: Utiliza memoização para otimizar cálculos recursivos, reiniciando o cache por combinação plano-DataSource.

## Estrutura do Projeto
- **`index.html`**: Interface web para visualização das demonstrações financeiras. Inclui um dropdown para selecionar visões e exibe os dados processados.
- **`processar_dados.py`**: Script Python que:
  - Lê e valida planilhas Excel.
  - Constrói hierarquias de contas.
  - Calcula valores com base em tipos de conta e fórmulas.
  - Gera visões financeiras e as injeta no `index.html` como JSON.
- **`DADOS.xlsx`**: Arquivo Excel esperado como entrada, contendo:
  - Abas de dados brutos (ex.: `Dados (DS1)`), com colunas para Código, Descrição e períodos (ex.: valores por mês/ano).
  - Abas de planos de contas (ex.: `Plano (Plano A)`), com colunas para Código, Descrição, Tipo (analítica, sintética, cálculo) e vínculos com DataSources.

## Requisitos
- **Python 3.6+**
- **Bibliotecas Python**:
  - `pandas`
  - `openpyxl`
  - `json`
- **Ambiente Web**:
  - Um navegador moderno para visualizar o `index.html`.
  - Opcionalmente, um servidor web local para testar a interface (ex.: `python -m http.server`).

## Instalação
1. Clone o repositório ou copie os arquivos do projeto para um diretório local.
2. Instale as dependências Python:
   ```bash
   pip install pandas openpyxl
   ```
3. Certifique-se de que o arquivo `DADOS.xlsx` está no mesmo diretório que o script `processar_dados.py`.

## Uso
1. Prepare o arquivo `DADOS.xlsx` com as abas de dados brutos e planos de contas, seguindo o formato esperado:
   - **Abas de dados**: Nomeadas como `Dados (nome_do_datasource)`, com colunas `Código`, `Descrição` e colunas de períodos (ex.: `Jan/2023`, `Fev/2023`).
   - **Abas de planos**: Nomeadas como `Plano (nome_do_plano)`, com colunas `Código`, `Descrição`, `Tipo` (analítica, sintética, cálculo) e colunas adicionais para vínculos com DataSources (ex.: `(DS1)`, `(DS2)`).
2. Execute o script Python para processar os dados e atualizar o `index.html`:
   ```bash
   python processar_dados.py
   ```
3. Abra o arquivo `index.html` em um navegador ou sirva-o com um servidor web local:
   ```bash
   python -m http.server 8000
   ```
4. Acesse `http://localhost:8000` no navegador, selecione uma visão no dropdown e visualize os resultados financeiros.

## Formato do Excel (`DADOS.xlsx`)
### Abas de Dados Brutos
- **Nome da aba**: `Dados (nome_do_datasource)` (ex.: `Dados (Financeiro 2023)`).
- **Colunas**:
  - `Código`: Código único do dado bruto (ex.: `COD1`).
  - `Descrição`: Descrição do dado bruto.
  - Colunas adicionais: Valores para cada período (ex.: `Jan/2023`, `Fev/2023`).

### Abas de Planos de Contas
- **Nome da aba**: `Plano (nome_do_plano)` (ex.: `Plano (DRE)`).
- **Colunas**:
  - `Código`: Código hierárquico da conta (ex.: `1`, `1.1`, `1.1.1`).
  - `Descrição`: Nome da conta, com fórmula entre parênteses para contas de cálculo (ex.: `Lucro (001 - 002)`).
  - `Tipo`: `analitica`, `sintetica` ou `calculo`.
  - Colunas adicionais: Vínculos com códigos brutos por DataSource (ex.: coluna `(DS1)` com valor `COD1;COD2`).

## Saída
- O script gera um objeto JSON com visões calculadas, injetado na variável `calculatedViewsData` no `index.html`.
- Cada visão contém:
  - `plan_name`: Nome do plano de contas.
  - `data_source_name`: Nome do DataSource.
  - `periodos`: Lista de períodos (ex.: `["Jan/2023", "Fev/2023"]`).
  - `accounts`: Lista de contas com código, descrição, tipo, nível hierárquico e valores calculados por período.

## Limitações
- Fórmulas em contas de cálculo devem ser expressões matemáticas simples (ex.: `001 + 002`, `001 - 002`) usando códigos de contas.
- Usa `eval()` para avaliar fórmulas, assumindo que são seguras. Evite inputs não confiáveis.
- Assume que o arquivo `index.html` contém uma variável JavaScript `calculatedViewsData` ou uma tag `<script>` para injeção de dados.
- Erros em fórmulas ou dados ausentes resultam em valores `0.0` para a conta afetada.

## Contribuição
1. Faça um fork do projeto.
2. Crie uma branch para sua feature (`git checkout -b feature/nova-funcionalidade`).
3. Commit suas mudanças (`git commit -m 'Adiciona nova funcionalidade'`).
4. Envie para o repositório remoto (`git push origin feature/nova-funcionalidade`).
5. Abra um Pull Request.

## Licença
Este projeto é distribuído sob a licença MIT. Veja o arquivo `LICENSE` para detalhes (se incluído).

## Contato
Para dúvidas ou sugestões, entre em contato com a equipe de desenvolvimento ou abra uma issue no repositório.