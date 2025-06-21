import pandas as pd
import json
import re
import os
from collections import defaultdict
from datetime import datetime
import datetime # Importar datetime explicitamente para usar timedelta

def get_name_in_parentheses(text):
    """Extrai o conteúdo entre o primeiro par de parênteses."""
    if isinstance(text, str):
        match = re.search(r'\((.*?)\)', text)
        if match:
            return match.group(1).strip()
    return None

def extract_formula_from_description(description):
    """
    Tenta extrair uma fórmula simples (ex: 1-2, 1+1.1) de uma descrição entre parênteses.
    NOTA: Esta função é mantida, mas a extração da fórmula para contas 'calculo'
    agora ocorrerá primariamente da coluna 'Tipo' na leitura do plano.
    """
    if isinstance(description, str):
        match = re.search(r'\((.*?)\)', description)
        if match:
            formula_candidate = match.group(1).strip()
            # Verifica se parece uma fórmula de cálculo financeiro simples
            if re.search(r'[\d\.\s\+\-\*\/\(\)]', formula_candidate):
                 return formula_candidate
    return None

def get_parent_code(account_code):
    """Retorna o código do pai para um dado código de conta (X.Y.Z -> X.Y)."""
    if isinstance(account_code, str):
        parts = account_code.split('.')
        if len(parts) > 1:
            return '.'.join(parts[:-1])
    return None # Retorna None se for código de nível superior ou inválido

def build_account_hierarchy(accounts_list):
    """Constrói uma estrutura de dicionário para lookup rápido e para identificar filhos."""
    # Dicionário principal para lookup rápido por código
    account_dict = {acc["codigo"]: acc for acc in accounts_list if "codigo" in acc and acc["codigo"] is not None}
    # Dicionário para armazenar filhos diretos de cada conta
    children_map = defaultdict(list)
    # Dicionário para armazenar o nível de cada conta
    level_map = {}

    # Preenche o mapa de filhos e calcula os níveis
    for account in accounts_list:
        if "codigo" not in account or account["codigo"] is None:
             continue
        codigo = account["codigo"]
        parent_code = get_parent_code(codigo)
        if parent_code:
            # Verifica se o pai existe no dicionário antes de adicionar como filho
            if parent_code in account_dict:
                 children_map[parent_code].append(codigo)
            else:
                 # Tratar caso de conta órfã cujo pai não está na lista
                 pass
        # Calcula o nível baseado na quantidade de pontos. Código '1' tem nível 1, '1.1' tem nível 2, etc.
        level_map[codigo] = codigo.count('.') + 1

    return account_dict, children_map, level_map

# Cache para armazenar valores calculados durante a execução de get_calculated_value
# Este cache é limpo para cada combinação Plano-DataSource
calculation_cache = {}

def safe_float_conversion(value, default=0.0):
    """Tenta converter um valor para float, retornando um default em caso de erro ou valor ausente."""
    if pd.notnull(value):
        try:
            # Tentar converter para float. Isso pode falhar se o valor for texto não numérico.
            return float(value)
        except (ValueError, TypeError):
            # print(f"Aviso: Valor não numérico '{value}' encontrado. Tratando como {default}.")
            return default
    return default # Retorna default para None/NaN

def get_calculated_value(account_code, period, account_dict, children_map, raw_data_for_ds_dict, ds_name):
    """
    Função recursiva para obter ou calcular o valor de uma conta para um período,
    usando dados de um DataSource específico.
    Usa o cache global (que é limpo por combinação Plano-DataSource) para memoização.
    """
    # Chave para o cache: (código da conta, período, nome do data source)
    cache_key = (account_code, period, ds_name)
    if cache_key in calculation_cache:
        return calculation_cache[cache_key]

    account = account_dict.get(account_code)
    if not account:
        # print(f"Aviso: Conta '{account_code}' referenciada na fórmula não encontrada no plano.")
        calculation_cache[cache_key] = 0.0 # Conta não existe no plano, valor é 0
        return 0.0

    value = 0.0 # Valor padrão inicial

    if account["tipo"].lower() == "analitica":
        # Para analíticas, buscamos e SOMAMOS os valores dos códigos brutos vinculados no data source atual
        data_code_mapping_str = account["data_sources"].get(ds_name) # string como 'COD1; COD2' ou 'COD1'

        if data_code_mapping_str and raw_data_for_ds_dict: # Verifica se há mapeamento para este DS e se o DataSource tem dados
            # Divide a string por ';' e limpa espaços, filtrando strings vazias
            linked_raw_codes = [code.strip() for code in str(data_code_mapping_str).split(';') if code.strip()]

            total_analitica_value = 0.0
            for raw_code in linked_raw_codes:
                # Busca o dado bruto correspondente no dicionário pré-carregado deste DataSource
                raw_data_entry = raw_data_for_ds_dict.get(raw_code)

                if raw_data_entry and raw_data_entry.get("valores"):
                    # Busca o valor bruto para o período específico
                    raw_value = raw_data_entry["valores"].get(period)
                    # Tenta converter o valor bruto para float, trata None/NaN/erros como 0.0
                    numeric_raw_value = safe_float_conversion(raw_value)
                    total_analitica_value += numeric_raw_value
                else:
                    # Aviso se um dos códigos brutos vinculados não foi encontrado no DataSource
                    pass

            value = total_analitica_value # O valor da analítica é a soma dos valores brutos vinculados

        else:
            # print(f"Aviso: Nenhum vínculo válido de dado bruto encontrado para conta analítica '{account_code}' no DataSource '{ds_name}'.")
            value = 0.0 # Nenhum vínculo válido encontrado para este DS, valor é 0

    elif account["tipo"].lower() == "sintetica":
        # Para sintéticas, somamos os valores dos filhos diretos
        direct_children_codes = children_map.get(account_code, [])
        total_sum = 0.0
        for child_code in direct_children_codes:
            # Recursivamente calcula o valor do filho para o mesmo período e DataSource
            child_value = get_calculated_value(child_code, period, account_dict, children_map, raw_data_for_ds_dict, ds_name)
            # Garante que o valor do filho é numérico antes de somar
            total_sum += safe_float_conversion(child_value)

        value = total_sum

    # ALTERAÇÃO CRÍTICA: Verifica se 'tipo' começa com 'calculo' E se a 'formula' foi extraída
    elif account["tipo"].lower().startswith("calculo") and account.get("formula"):
        # Para contas de cálculo, avaliamos a fórmula (que já foi extraída na leitura do plano)
        formula = account["formula"]
        try:
            # Encontra todos os códigos de contas mencionados na fórmula
            # Procura por padrões de código (ex: 001, 001.01.02) que são números e pontos
            referenced_codes = re.findall(r'(\d+(\.\d+)*)', formula)

            # Cria um ambiente local para a avaliação da fórmula, mapeando códigos para seus valores calculados
            local_vars = {}
            for ref_code_tuple in referenced_codes:
                ref_code = ref_code_tuple[0] # O código completo (ex: '001.01.02')
                # Recursivamente obtém o valor calculado da conta referenciada
                # Passa o account_dict, children_map, raw_data_for_ds_dict e ds_name adiante
                ref_value = get_calculated_value(ref_code, period, account_dict, children_map, raw_data_for_ds_dict, ds_name)
                # Adiciona ao ambiente local com um nome de variável válido (sanitiza códigos)
                var_name = f"v_{ref_code.replace('.', '_').replace('-', '_')}" # ex: v_001_01_02
                local_vars[var_name] = safe_float_conversion(ref_value) # Usa 0.0 se o valor referenciado não for numérico

            # Substitui os códigos na fórmula pelos nomes das variáveis no ambiente local
            eval_formula = formula
            # Substitui os códigos mais longos primeiro para evitar substituir partes de códigos mais curtos
            referenced_codes.sort(key=lambda x: len(x[0]), reverse=True)
            for ref_code_tuple in referenced_codes:
                ref_code = ref_code_tuple[0]
                var_name = f"v_{ref_code.replace('.', '_').replace('-', '_')}"
                # Usa regex para substituir apenas o código completo e não partes dele (usa word boundary \b)
                eval_formula = re.sub(r'\b' + re.escape(ref_code) + r'\b', var_name, eval_formula)
            
            # Avalia a expressão usando eval().
            value = eval(eval_formula, {}, local_vars)

            # Garantir que o resultado da avaliação seja numérico
            value = safe_float_conversion(value)

        except Exception as e:
            # print(f"Erro CRÍTICO ao avaliar fórmula '{formula}' (após subst: '{eval_formula}') para conta '{account_code}' no DataSource '{ds_name}', período '{period}': {e}")
            value = 0.0 # Em caso de erro na fórmula, o valor é 0


    # Armazena o resultado no cache para esta combinação Plano-DataSource e período
    calculation_cache[cache_key] = value
    return value


def processar_planilha_integrado(caminho_excel='DADOS.xlsx'):
    """
    Processa a planilha Excel para extrair dados, planos, vincular, calcular
    e estruturar em um objeto JSON de visões calculadas.
    """
    all_raw_data = {} # Armazena dados brutos por DataSource: {ds_name: {periodos:[], data:{codigo:{descricao, valores:{periodo:valor}}}}}
    all_plans = {} # Armazena definições de planos por nome de plano: {plan_name: {"accounts_list":[{codigo, descricao, tipo, formula, data_sources}], "linked_data_sources":[]}}
    calculated_views = [] # Lista final de visões calculadas

    try:
        xls = pd.ExcelFile(caminho_excel, engine='openpyxl')
    except FileNotFoundError:
        print(f"Erro: Arquivo '{caminho_excel}' não encontrado.")
        return None
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        print("Por favor, verifique se o arquivo Excel está salvo em um formato compatível (.xlsx ou .xls) e não está corrompido.")
        return None

    print("\n--- Lendo dados brutos (abas 'dados') ---")
    for sheet_name in xls.sheet_names:
        # Ignorar arquivos temporários do Excel e abas que não começam com 'dados'
        if sheet_name.startswith('~$') or not sheet_name.lower().startswith("dados"):
            continue

        ds_name = get_name_in_parentheses(sheet_name)
        if not ds_name:
            ds_name = sheet_name.strip()
        print(f"Processando aba de Dados brutos: '{sheet_name}' -> Nome: '{ds_name}'")
        try:
            df_dados = xls.parse(sheet_name)
            colunas_dados_raw = df_dados.columns.tolist()
            if len(colunas_dados_raw) < 3:
                 print(f"Aviso: A aba '{sheet_name}' de dados não tem colunas de dados/períodos suficientes. Ignorando.")
                 continue

            coluna_codigo_raw = colunas_dados_raw[0]
            coluna_descricao_raw = colunas_dados_raw[1]
            colunas_periodos_raw = colunas_dados_raw[2:]

            # Limpar linhas sem código na primeira coluna
            df_dados.dropna(subset=[coluna_codigo_raw], inplace=True)
            # Tentar limpar a primeira linha (cabeçalho) caso seja um DataFrame multi-indexado ou com lixo inicial
            if not df_dados.empty and str(df_dados.iloc[0][coluna_codigo_raw]).strip().lower() == str(coluna_codigo_raw).strip().lower():
                df_dados = df_dados.iloc[1:].copy()
                df_dados.reset_index(drop=True, inplace=True)
                colunas_dados_raw = df_dados.columns.tolist()
                coluna_codigo_raw = colunas_dados_raw[0]
                coluna_descricao_raw = colunas_dados_raw[1]
                colunas_periodos_raw = colunas_dados_raw[2:]
                print(" - Primeira linha removida, parece ser cabeçalho repetido.")


            registros_dados = []
            periodos_ds = [] # Coleta os períodos deste DataSource

            # Processa cabeçalhos de período para obter a lista de períodos formatados
            for periodo_header_raw in colunas_periodos_raw:
                if isinstance(periodo_header_raw, pd.Timestamp):
                    # Formata datas como 'MMM/AAAA'
                    formatted_header = periodo_header_raw.strftime('%Y-%m-%d %H:%M:%S') # Mantém o formato original de timestamp
                else:
                    # Mantém como string e limpa espaços
                    formatted_header = str(periodo_header_raw).strip()
                periodos_ds.append(formatted_header)

            # Processa as linhas de dados
            for index, row in df_dados.iterrows():
                # Pega o código e descrição da linha
                codigo_dado = str(row[coluna_codigo_raw]).strip() if pd.notnull(row[coluna_codigo_raw]) else None
                descricao_dado = str(row[coluna_descricao_raw]).strip() if pd.notnull(row[coluna_descricao_raw]) else None

                if not codigo_dado: # Pula linhas sem código na coluna de código
                    continue

                valores_por_periodo = {}
                # Itera sobre as colunas de período para coletar os valores
                for i, periodo_header_raw in enumerate(colunas_periodos_raw):
                    periodo_formatted = periodos_ds[i] # Usa o cabeçalho formatado/limpo
                    valor = row[periodo_header_raw]

                    # Aqui NÃO fazemos conversão forte para float, apenas armazenamos o valor como lido
                    # A conversão numérica e tratamento de None/NaN/erros é feita na função get_calculated_value
                    if pd.isna(valor):
                        valor = None # Armazena None para NaN

                    valores_por_periodo[periodo_formatted] = valor

                registros_dados.append({
                    "codigo": codigo_dado,
                    "descricao": descricao_dado,
                    "valores": valores_por_periodo
                })

            # Armazena os dados brutos deste DS como um dicionário {codigo: {...}} para lookup rápido
            if registros_dados:
                all_raw_data[ds_name] = {
                    "periodos": periodos_ds,
                    "data": {item["codigo"]: item for item in registros_dados if item["codigo"] is not None}
                }
                print(f" - {len(registros_dados)} registros de dados brutos extraídos e indexados por código para '{ds_name}' com {len(periodos_ds)} períodos.")
            else:
                print(f" - Nenhuns registros de dados brutos válidos encontrados na aba '{sheet_name}'.")


        except Exception as e:
            print(f"Erro ao processar a aba '{sheet_name}': {e}")
            print("Verifique as colunas ('Código', 'Descrição', Períodos...) e o formato dos dados.")


    print("\n--- Lendo Planos de Contas (abas 'plano') ---")
    for sheet_name in xls.sheet_names:
        # Ignorar arquivos temporários do Excel e abas que não começam com 'plano'
        if sheet_name.startswith('~$') or not sheet_name.lower().startswith("plano"):
            continue

        plan_name = get_name_in_parentheses(sheet_name)
        if not plan_name:
             plan_name = sheet_name.strip()
        print(f"Processando aba de Plano: '{sheet_name}' -> Nome: '{plan_name}'")
        try:
            df_plano = xls.parse(sheet_name)

            colunas_plano_raw = df_plano.columns.tolist()
            if len(colunas_plano_raw) < 3: # Mínimo: Código, Descrição, Tipo
                 print(f"Aviso: A aba '{sheet_name}' de plano não tem as 3 colunas básicas esperadas. Ignorando.")
                 continue

            coluna_codigo_raw = colunas_plano_raw[0]
            coluna_descricao_raw = colunas_plano_raw[1]
            coluna_tipo_raw = colunas_plano_raw[2]
            colunas_vinculo_raw = colunas_plano_raw[3:] # Da 4ª coluna em diante

            data_source_columns_map = {} # Mapeamento de cabeçalho de coluna -> Nome DataSource (do parênteses)
            linked_data_sources_in_plan = set() # Rastreia quais DataSources este plano *referencia* nas colunas de vínculo
            # Itera sobre as colunas a partir da 4ª para identificar os DataSources vinculados
            for col_header_raw in colunas_vinculo_raw:
                ds_name_from_col = get_name_in_parentheses(col_header_raw)
                if ds_name_from_col:
                    # Mapeia o cabeçalho original da coluna para o nome do DataSource extraído
                    data_source_columns_map[col_header_raw] = ds_name_from_col
                    # Adiciona o nome do DataSource ao conjunto de DataSources referenciados por este plano
                    linked_data_sources_in_plan.add(ds_name_from_col)
                else:
                    print(f"Aviso: Cabeçalho da coluna '{col_header_raw}' na aba '{sheet_name}' não tem nome de data source entre parênteses. Ignorando esta coluna para vínculos.")


            df_plano.dropna(subset=[coluna_codigo_raw], inplace=True)
            # Tentar limpar a primeira linha (cabeçalho) caso seja um DataFrame multi-indexado ou com lixo inicial
            if not df_plano.empty and str(df_plano.iloc[0][coluna_codigo_raw]).strip().lower() == str(coluna_codigo_raw).strip().lower():
                df_plano = df_plano.iloc[1:].copy()
                df_plano.reset_index(drop=True, inplace=True)
                print(" - Primeira linha removida, parece ser cabeçalho repetido.")


            accounts_list = []
            for index, row in df_plano.iterrows():
                # Pega os dados básicos da conta
                codigo = str(row[coluna_codigo_raw]).strip() if pd.notnull(row[coluna_codigo_raw]) else None
                descricao = str(row[coluna_descricao_raw]).strip() if pd.notnull(row[coluna_descricao_raw]) else None
                tipo_excel_value = str(row[coluna_tipo_raw]).strip() if pd.notnull(row[coluna_tipo_raw]) else None

                # Mantém o valor do tipo_excel_value exatamente como no Excel para armazenamento
                tipo_to_store = tipo_excel_value

                # Tenta extrair a fórmula do campo 'Tipo' se ele começar com "calculo"
                extracted_formula = None
                if tipo_excel_value and tipo_excel_value.lower().startswith("calculo"):
                    match_formula_in_type = re.search(r'calculo\s*\((.*?)\)', tipo_excel_value, re.IGNORECASE)
                    if match_formula_in_type:
                        extracted_formula = match_formula_in_type.group(1).strip()
                
                # Fallback para o tipo se o valor do Excel estiver vazio
                if not tipo_to_store:
                    print(f"Aviso: Tipo de conta inválido ou ausente para o código {codigo}. Usando 'sintetica' como padrão.")
                    tipo_to_store = "sintetica"

                data_sources_map = {} # Dicionário {ds_name: codigo_bruto_referenciado_string}
                # Itera sobre as colunas de vínculo identificadas para esta linha/conta
                for col_raw, ds_name_mapped in data_source_columns_map.items():
                    data_code_in_plan_cell = row[col_raw] # O valor da célula (pode ser 'COD1; COD2')
                    if pd.notnull(data_code_in_plan_cell):
                         # Armazena a string bruta do vínculo, incluindo ';' se houver, no mapeamento
                        data_sources_map[ds_name_mapped] = str(data_code_in_plan_cell).strip()


                accounts_list.append({
                    "codigo": codigo,
                    "descricao": descricao,
                    "tipo": tipo_to_store,  # Armazena o texto completo, ex: "calculo (001 - 002)"
                    "formula": extracted_formula, # Armazena apenas a fórmula extraída, ex: "001 - 002"
                    "data_sources": data_sources_map,
                    "valores": {} # Inicializa o dicionário de valores
                })

            # Ordenar as contas pelo código para manter a hierarquia visual no JSON (opcional, mas recomendado)
            accounts_list.sort(key=lambda x: x['codigo'])

            if accounts_list:
                # Armazena as informações do plano, incluindo os DataSources que ele referencia
                all_plans[plan_name] = {
                    "accounts_list": accounts_list,
                    "linked_data_sources": sorted(list(linked_data_sources_in_plan.intersection(all_raw_data.keys())))
                }
                print(f" - {len(accounts_list)} contas de plano extraídas e mapeadas para '{plan_name}'.")
                print(f" - Este plano referencia os DataSources encontrados: {all_plans[plan_name]['linked_data_sources']}")
            else:
                print(f" - Nenhuma conta válida encontrada na aba '{sheet_name}'.")


        except Exception as e:
            print(f"Erro ao processar a aba '{sheet_name}': {e}")
            print("Verifique as colunas ('Código', 'Descrição', 'Tipo') e os cabeçalhos/códigos das colunas de vínculo.")

    # --- Etapa de Cálculo e Geração de Visões Integradas ---
    print("\n--- Realizando Cálculos e Gerando Visões Integradas ---")
    global calculation_cache 

    # Itera por cada plano que foi lido com sucesso
    for plan_name, plan_info in all_plans.items():
        accounts_list = plan_info["accounts_list"]
        linked_data_sources = plan_info["linked_data_sources"] 

        if not accounts_list or not linked_data_sources:
             print(f"Pulando cálculo para o plano '{plan_name}': sem contas válidas ou DataSources vinculados encontrados.")
             continue

        print(f"Calculando visões para o plano '{plan_name}'...")

        # Constrói a hierarquia para este plano (uma vez por plano)
        account_dict, children_map, level_map = build_account_hierarchy(accounts_list)

        # Para cada DataSource que este plano referencia E foi lido com sucesso
        for ds_name in sorted(linked_data_sources): 
            # print(f" - Gerando visão calculada para '{plan_name}' usando DataSource '{ds_name}'...") 

            raw_data_info = all_raw_data[ds_name] 
            raw_data_for_ds_dict = raw_data_info["data"] 
            periodos_ds = raw_data_info["periodos"] 

            if not periodos_ds:
                 print(f" - DataSource '{ds_name}' não tem períodos. Pulando cálculo para esta combinação.")
                 continue

            # Limpa o cache de cálculo para a NOVA COMBINAÇÃO (Plano + DataSource)
            calculation_cache = {}

            # Lista para armazenar as contas deste plano COM OS VALORES CALCULADOS para este DataSource
            calculated_accounts_for_view = []

            # Itera por cada conta do plano para calcular seus valores para todos os períodos deste DataSource
            for account in accounts_list:
                account_codigo = account["codigo"]
                # Dicionário para armazenar os valores calculados/obtidos desta conta para todos os períodos deste DataSource
                account_calculated_values = {}
  
                # Itera por cada período deste DataSource
                for period in periodos_ds:
                    # Chama a função de cálculo para obter o valor da conta neste período, usando os dados deste DataSource
                    value = get_calculated_value(account_codigo, period, account_dict, children_map, raw_data_for_ds_dict, ds_name)
                    account_calculated_values[period] = value
  
                # Cria uma cópia da conta original do plano
                calculated_account = account.copy()
                # Adiciona o dicionário de valores calculados/obtidos para este DataSource
                calculated_account["valores"] = account_calculated_values
                # Adiciona o nível hierárquico
                calculated_account["nivel"] = level_map.get(account_codigo, 1)
                # Remove o mapeamento bruto para os DataSources para a saída final do JSON, se presente
                if "data_sources" in calculated_account:
                    del calculated_account["data_sources"]
                
                # Mantém o campo 'formula' na saída JSON se ele tiver sido preenchido
                # Não há necessidade de 'del' se o get("formula") for None, ele simplesmente não será incluído.
  
                # --- INÍCIO DA NOVA LÓGICA: FILTRAR CONTAS SINTÉTICAS COM VALOR ZERO ---
                should_add_account = True
                # Verifica se a conta é sintética e se todos os seus valores são zero
                if calculated_account["tipo"].lower() in ["sintetica", "analitica"]:
                    # Verifica se todos os valores para esta conta sintética ou analitica são 0.0
                    # Isso garante que se houver *qualquer* valor diferente de zero, a conta seja mantida.
                    all_values_are_zero = all(val == 0.0 for val in calculated_account["valores"].values())
                    if all_values_are_zero:
                        should_add_account = False
                
                if should_add_account:
                    calculated_accounts_for_view.append(calculated_account)
                # --- FIM DA NOVA LÓGICA ---

            # Adiciona a visão calculada completa (Plano, DataSource, Períodos, Contas com Valores) à lista final de visões
            calculated_views.append({
                "plan_name": plan_name,
                "data_source_name": ds_name,
                "periodos": periodos_ds, # Inclui os períodos usados por este DataSource (já ordenado)
                "accounts": calculated_accounts_for_view # Lista de contas deste plano com valores calculados para este DataSource
            })
            print(f" - Visão calculada gerada para '{plan_name}' com '{ds_name}'. ({len(calculated_accounts_for_view)} contas)")
    # Retorna o dicionário final contendo a lista de todas as visões calculadas
    return {"calculated_views": calculated_views}


def update_index_html_with_json(caminho_html='index.html', dados_json={}):
    """
    Lê o arquivo index.html, encontra a definição da variável JS alvo e a substitui/injeta
    o JSON calculado. Preserva o restante do HTML.
    """
    js_variable_name = 'calculatedViewsData'
    pattern = r'(const|let|var)\s+' + re.escape(js_variable_name) + r'\s*=\s*({.*?});'

    html_content = None
    try:
        with open(caminho_html, 'r', encoding='utf-8') as f:
            html_content = f.read()
        print(f"Arquivo '{caminho_html}' encontrado. Lendo conteúdo existente.")

    except FileNotFoundError:
        print(f"Erro: Arquivo '{caminho_html}' não encontrado.")
        print("Não é possível atualizar um arquivo que não existe. Por favor, crie o arquivo index.html com a estrutura básica e a variável JavaScript alvo.")
        return False 

    json_data_str = json.dumps(dados_json, indent=4, ensure_ascii=False, sort_keys=False, separators=(',', ': '))
    full_json_definition_line = f"const {js_variable_name} = {json_data_str};"

    match = re.search(pattern, html_content, re.DOTALL)

    if match:
        print(f"Padrão '{match.group(1)} {js_variable_name} = ...;' encontrado no arquivo '{caminho_html}'. Substituindo dados...")
        novo_html_content = html_content[:match.start(2)] + json_data_str + html_content[match.end(2):]

    else:
        print(f"Padrão '{js_variable_name} = ...;' NÃO encontrado no arquivo '{caminho_html}'. Tentando injetar a definição completa da variável antes de </script>...")
        script_end_tag_pattern = r'</script>'
        match_script_end = re.search(script_end_tag_pattern, html_content, re.DOTALL)

        if match_script_end:
            injection_string = f"\n\n        // Dados injetados pelo script Python:\n        {full_json_definition_line}\n\n        " 
            novo_html_content = html_content[:match_script_end.start()] + injection_string + html_content[match_script_end.start():]
            print("Definição completa da variável JSON injetada dentro da tag <script>.")
        else:
            print(f"Erro: Não foi possível encontrar o padrão '{js_variable_name} = ...;' nem a tag </script> no arquivo '{caminho_html}'. Não foi possível injetar os dados.")
            return False 


    try:
        with open(caminho_html, 'w', encoding='utf-8') as f:
            f.write(novo_html_content) 
        print(f"Arquivo '{caminho_html}' atualizado com sucesso!")
        return True 
    except Exception as e:
        print(f"Erro ao escrever no arquivo '{caminho_html}': {e}")
        return False 


# --- Bloco de Execução Principal ---
if __name__ == "__main__":
    now_utc = datetime.datetime.now(datetime.timezone.utc)
    now_br = now_utc - datetime.timedelta(hours=3)
    time_info = f"Atualizado em {now_br.strftime('%d/%m/%Y %H:%M:%S')} (UTC-3) / {now_utc.strftime('%d/%m/%Y %H:%M:%S')} (UTC)"

    nome_arquivo_excel = 'DADOS.xlsx' 
    nome_arquivo_html = 'index.html' 


    print(f"Iniciando processamento integrado do arquivo '{nome_arquivo_excel}'...")

    if not os.path.exists(nome_arquivo_excel):
        print(f"Erro: O arquivo '{nome_arquivo_excel}' não foi encontrado no diretório atual.")
        print("Certifique-se de que a planilha está na mesma pasta do script Python.")
    else:
        dados_extraidos = processar_planilha_integrado(nome_arquivo_excel)

        if dados_extraidos is not None:
            if dados_extraidos.get("calculated_views"):
                dados_extraidos["timestamp_utc"] = datetime.datetime.utcnow().isoformat() + 'Z'
                dados_extraidos["timestamp_local"] = datetime.datetime.now().astimezone().isoformat()

                if update_index_html_with_json(nome_arquivo_html, dados_extraidos):
                    print(f"\nProcessamento integrado concluído. {time_info}")
                    print(f"Arquivo '{nome_arquivo_html}' gerado/atualizado com a estrutura JSON calculada.")

                    num_views = len(dados_extraidos["calculated_views"])
                    print(f"Total de {num_views} visão(ões) calculada(s) gerada(s).")
                    for view in dados_extraidos["calculated_views"]:
                        print(f" - '{view['plan_name']}' com DataSource '{view['data_source_name']}' ({len(view['accounts'])} contas, {len(view['periodos'])} períodos)")

                    print("\nA próxima etapa é abrir o arquivo index.html gerado em um navegador, selecionar uma visão no dropdown e verificar a tabela.")
                    print("Se os dados na tabela parecerem incorretos:")
                    print("1. Verifique a saída do script no console para quaisquer avisos de 'Código de dado bruto ... não encontrado'. Isso indica que os códigos de vínculo na sua planilha Excel não correspondem exatamente aos códigos nas abas de dados. Corrija-os no Excel.")
                    print("2. Verifique as fórmulas das contas do tipo 'calculo' na planilha (entre parênteses na descrição ou na coluna 'Tipo' conforme o novo modelo). A avaliação de fórmula implementada é básica (suporta +, -, *, /, parênteses e códigos de contas como referências numéricas) e pode não suportar fórmulas complexas ou com nomes de texto. Erros na fórmula resultarão em valor 0.0.")
                    print("3. Confirme que os nomes entre parênteses nos cabeçalhos das colunas de vínculo nas abas 'plano' correspondem exatamente aos nomes entre parênteses das abas 'dados'.")
                    print("4. Verifique se as abas de dados e plano têm as colunas obrigatórias na ordem correta (Código, Descrição, Tipo para plano; Código, Descrição, Períodos... para dados).")

                else:
                    print(f"\nProcessamento concluído, mas houve um erro ao atualizar o arquivo '{nome_arquivo_html}'.")


            else:
                print("\nProcessamento concluído, mas nenhuma visão calculada válida foi gerada.")
                print("Verifique se suas abas estão nomeadas corretamente (iniciando com 'plano(...)' e 'dados(...)') e se os Planos referenciam DataSources existentes com dados.")

        else:
            print(f"\nProcessamento falhou devido a erros na leitura inicial da planilha. {time_info}")