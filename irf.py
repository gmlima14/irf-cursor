""" # Importar as bibliotecas necessárias """

# Importa o módulo de classificação do PyCaret
from pycaret.classification import load_model, predict_model
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo  # disponível a partir do Python 3.9
import os
import warnings
warnings.filterwarnings('ignore')

"""# Configuração de caminhos"""

# Caminhos dos arquivos da rede
ARQUIVO_REDE = r'S:\Procurement\FUP\OTP Mensal\OTP - Base.xlsm'
MODELO_REDE = r'S:\Procurement\FUP\IRF - Índice de Risco de Fornecedores\Modelo de Machine Learning\modelo_treinado_alldates.pkl'
CARGA_REDE = r'S:\Procurement\FUP\IRF - Índice de Risco de Fornecedores\Modelo de Machine Learning\carga_fornecedor.csv'

def verificar_caminhos():
    """
    Verifica se todos os arquivos necessários estão disponíveis na rede.
    
    Returns:
        dict: Dicionário com os caminhos dos arquivos encontrados ou None se algum arquivo não for encontrado
    """
    caminhos_disponiveis = {}
    
    # Verifica arquivo de dados
    if os.path.exists(ARQUIVO_REDE):
        caminhos_disponiveis['dados'] = ARQUIVO_REDE
        print(f"✅ Arquivo de dados encontrado na rede: {ARQUIVO_REDE}")
    else:
        print(f"❌ Arquivo de dados não encontrado na rede: {ARQUIVO_REDE}")
        return None
    
    # Verifica modelo
    if os.path.exists(MODELO_REDE):
        caminhos_disponiveis['modelo'] = MODELO_REDE
        print(f"✅ Modelo encontrado na rede: {MODELO_REDE}")
    else:
        print(f"❌ Modelo não encontrado na rede: {MODELO_REDE}")
        return None
    
    # Verifica arquivo de carga
    if os.path.exists(CARGA_REDE):
        caminhos_disponiveis['carga'] = CARGA_REDE
        print(f"✅ Arquivo de carga encontrado na rede: {CARGA_REDE}")
    else:
        print(f"❌ Arquivo de carga não encontrado na rede: {CARGA_REDE}")
        return None
    
    return caminhos_disponiveis

"""# Lê o arquivo excel com os pedidos em aberto"""

def carregar_dados(caminhos):
    """
    Carrega os dados do arquivo Excel da rede e faz limpeza inicial.
    
    Args:
        caminhos (dict): Dicionário com os caminhos dos arquivos
        
    Returns:
        pandas.DataFrame: DataFrame com os dados carregados ou None se houver erro
    """
    try:
        print(f"📊 Carregando dados do arquivo: {caminhos['dados']}")
        df_pedidos_em_aberto = pd.read_excel(caminhos['dados'])
        
        # Remove linhas que contêm dados na coluna Delivery Date
        if 'Delivery Date' in df_pedidos_em_aberto.columns:
            df_pedidos_em_aberto = df_pedidos_em_aberto[df_pedidos_em_aberto['Delivery Date'].isna()]
            print(f"🗑️ Mantidas apenas linhas com coluna 'Delivery Date' vazia")
        # Remove a coluna 'On Time' se existir
        if 'On Time' in df_pedidos_em_aberto.columns:
            df_pedidos_em_aberto = df_pedidos_em_aberto.drop(columns=['On Time'])
            print(f"🗑️ Removida coluna 'On Time'")
        print(f"✅ Dados carregados com sucesso! {len(df_pedidos_em_aberto)} registros encontrados")
        return df_pedidos_em_aberto
    except Exception as e:
        print(f"❌ Erro ao carregar arquivo: {e}")
        return None

"""# Previsão de Atrasos de Pedidos em Aberto"""

def processar_dados(df_pedidos_em_aberto):
    """
    Processa e prepara os dados para análise de machine learning.
    
    Args:
        df_pedidos_em_aberto (pandas.DataFrame): DataFrame com os dados brutos
        
    Returns:
        pandas.DataFrame: DataFrame processado com variáveis calculadas ou None se houver erro
    """
    if df_pedidos_em_aberto is None:
        return None
    
    # Transforma o tipo para categoria
    df_pedidos_em_aberto['MATKL'] = df_pedidos_em_aberto['MATKL'].astype('category')
    df_pedidos_em_aberto['Vendor'] = df_pedidos_em_aberto['Vendor'].astype('category')

    # Converter as colunas de data
    df_pedidos_em_aberto["BEDAT"] = pd.to_datetime(df_pedidos_em_aberto["BEDAT"], errors="coerce")  # Data de emissão
    df_pedidos_em_aberto["Due Date (incl. ex works time)"] = pd.to_datetime(df_pedidos_em_aberto["Due Date (incl. ex works time)"], errors="coerce")  # Entrega prevista

    hoje = datetime.today()

    # Variáveis de tempo
    df_pedidos_em_aberto["MesPedido"] = df_pedidos_em_aberto["BEDAT"].dt.month
    df_pedidos_em_aberto["IdadePedido"] = (hoje - df_pedidos_em_aberto["BEDAT"]).dt.days
    df_pedidos_em_aberto["DiasParaEntrega"] = (df_pedidos_em_aberto["Due Date (incl. ex works time)"] - df_pedidos_em_aberto["BEDAT"]).dt.days

    # Conta a quantidade de pedidos em aberto por fornecedor
    pedidos_abertos_por_fornecedor = df_pedidos_em_aberto['Vendor'].value_counts()

    # Mapeia para o dataframe principal
    df_pedidos_em_aberto['carga_fornecedor'] = df_pedidos_em_aberto['Vendor'].map(pedidos_abertos_por_fornecedor).fillna(0).astype(int)
    
    return df_pedidos_em_aberto

def carregar_modelo(caminhos):
    """
    Carrega o modelo de machine learning treinado da rede.
    
    Args:
        caminhos (dict): Dicionário com os caminhos dos arquivos
        
    Returns:
        object: Modelo carregado ou None se houver erro
    """
    try:
        print(f"🤖 Carregando modelo de machine learning da rede...")
        modelo = load_model(caminhos['modelo'].replace('.pkl', ''))
        print(f"✅ Modelo carregado com sucesso!")
        return modelo
    except Exception as e:
        print(f"❌ Erro ao carregar modelo: {e}")
        return None

def fazer_previsoes(modelo, df_pedidos_em_aberto):
    """
    Faz previsões de atraso usando o modelo de machine learning.
    
    Args:
        modelo (object): Modelo de machine learning carregado
        df_pedidos_em_aberto (pandas.DataFrame): DataFrame com dados processados
        
    Returns:
        pandas.DataFrame: DataFrame com previsões e confiabilidade ou None se houver erro
    """
    if modelo is None or df_pedidos_em_aberto is None:
        return None
    
    try:
        print("🔮 Fazendo previsões...")
        previsoes = predict_model(modelo, data=df_pedidos_em_aberto)
        
        # Renomeia as colunas
        previsoes.rename(columns={
            "prediction_label": "Previsão",
            "prediction_score": "Confiabilidade",
            "carga_fornecedor": "Carga do Fornecedor",
            "EBELN": "PO",
            "EBELP": "Item",
            "BEDAT": "Data de Emissão da PO",
            "Due Date (incl. ex works time)": "Stat. Del. Date",
            "Material Text (AST or Short Text)": "Descrição do Item",
            "Vendor Name": "Fornecedor",
            "MATKL": "Material Number",
            "NetOrderValue": "Valor Net",
            "MesPedido": "Mês do Pedido",
            "IdadePedido": "Idade do Pedido",
            "DiasParaEntrega": "Dias para Entrega",
        }, inplace=True)

        # Altera os valores de 0 e 1
        previsoes["Previsão"] = previsoes["Previsão"].replace({0: "No Prazo", 1: "Atraso"})
        
        print("✅ Previsões concluídas!")
        return previsoes
    except Exception as e:
        print(f"❌ Erro ao fazer previsões: {e}")
        return None

"""# Matriz de Fornecedores"""

def criar_matriz_fornecedores(previsoes, caminhos):
    """
    Cria a matriz de fornecedores com cálculo do índice de risco.
    
    Args:
        previsoes (pandas.DataFrame): DataFrame com previsões
        caminhos (dict): Dicionário com os caminhos dos arquivos
        
    Returns:
        pandas.DataFrame: DataFrame com matriz de fornecedores e índice de risco ou None se houver erro
    """
    if previsoes is None:
        return None
    
    try:
        print("📊 Criando matriz de fornecedores...")
        
        # 1. Agrupar previsões por fornecedor
        agrupado = previsoes.groupby('Vendor').agg(
            pedidos_no_prazo=('Previsão', lambda x: (x == 'No Prazo').sum()),
            pedidos_atrasados=('Previsão', lambda x: (x == 'Atraso').sum()),
            total_pedidos=('Previsão', 'count'),
            confiabilidade_media=('Confiabilidade', 'mean'),
            valor_total=('Valor Net', 'sum'),
            valor_atrasado=('Valor Net', lambda x: (x[previsoes.loc[x.index, 'Previsão'] == 'Atraso']).sum()),
            valor_no_prazo=('Valor Net', lambda x: (x[previsoes.loc[x.index, 'Previsão'] == 'No Prazo']).sum()),
            fornecedor=('Fornecedor', 'first')
        ).reset_index()

        # Ajusta o valor total para float64
        agrupado['valor_total'] = agrupado['valor_total'].astype(float)

        # Lê o arquivo csv de carga de fornecedores da rede
        df_carga = pd.read_csv(caminhos['carga'])

        # Renomeia a coluna no df carga para carga media
        df_carga.rename(columns={'carga_fornecedor': 'carga_media'}, inplace=True)

        # Faz o merge com o DataFrame de pedidos em aberto
        agrupado = agrupado.merge(df_carga[['Vendor', 'carga_media']], on='Vendor', how='left')

        # Calcular taxa de atraso
        agrupado['taxa_no_prazo'] = agrupado['pedidos_no_prazo'] / agrupado['total_pedidos']

        # Calcular taxa de valor
        agrupado['taxa_valor'] = agrupado['valor_no_prazo'] / agrupado['valor_total']

        # Calcula a taxa de carga conforme as regras
        def calcular_taxa_carga(row):
            """
            Calcula a taxa de carga do fornecedor baseada na carga média.
            
            Args:
                row: Linha do DataFrame com carga_media e total_pedidos
                
            Returns:
                float: Taxa de carga calculada
            """
            carga_media = row['carga_media']
            total_pedidos = row['total_pedidos']

            if pd.isna(carga_media) or carga_media <= 2:
                return 1
            taxa = total_pedidos / carga_media
            return min(max(taxa, 1), 1.5)

        agrupado['taxa_carga'] = agrupado.apply(calcular_taxa_carga, axis=1)

        # Índice de risco bruto = taxa de atraso × confiabilidade
        agrupado['indice_bruto'] = agrupado['taxa_no_prazo'] * agrupado['confiabilidade_media'] * agrupado['taxa_valor'] / agrupado['taxa_carga']

        # Normalizar
        agrupado['indice_risco'] = (1 - agrupado['indice_bruto'])*100

        # Arredonda para 2 casas
        agrupado = agrupado.round(2)

        # Arredonda a carga media para zero cargas decimais e transforma em int
        agrupado['carga_media'] = agrupado['carga_media'].apply(lambda x: int(x) if x >= 1 else 0)

        # Renomear para deixar claro
        df_fornecedores = agrupado[[
            'fornecedor', 'Vendor', 'pedidos_no_prazo', 'pedidos_atrasados', 'taxa_no_prazo',
            'total_pedidos', 'carga_media', 'taxa_carga',  'valor_total', 'taxa_valor', 'confiabilidade_media',  'indice_risco'
        ]]

        # Ordena por indice de risco
        df_fornecedores = df_fornecedores.sort_values('indice_risco', ascending=False).reset_index(drop=True)
        df_fornecedores['Ranking'] = df_fornecedores.index + 1

        # Altera a coluna Ranking para a primeira coluna
        primeira_coluna = df_fornecedores.pop('Ranking')
        df_fornecedores.insert(0, 'Ranking', primeira_coluna)

        # Renomeia as colunas do df
        df_fornecedores = df_fornecedores.rename(columns={
            'fornecedor': 'Fornecedor',
            'pedidos_no_prazo': 'PO previstas no prazo',
            'pedidos_atrasados': 'PO previstas atrasadas',
            'taxa_no_prazo': 'Taxa de PO previstas no prazo',
            'total_pedidos': 'Total de PO',
            'carga_media': 'Carga Média de PO',
            'taxa_carga': 'Taxa de Carga',
            'valor_total': 'Valor NET de PO',
            'taxa_valor': 'Taxa de Valor previsto no prazo',
            'indice_risco': 'Índice de Risco',
            'confiabilidade_media': 'Confiabilidade Média'
        })
        
        print("✅ Matriz de fornecedores criada!")
        return df_fornecedores
    except Exception as e:
        print(f"❌ Erro ao criar matriz de fornecedores: {e}")
        return None

"""# Download do arquivo"""

def salvar_resultados(df_fornecedores, previsoes):
    """
    Salva os resultados em arquivo Excel na pasta de histórico da rede.
    
    Args:
        df_fornecedores (pandas.DataFrame): DataFrame com matriz de fornecedores
        previsoes (pandas.DataFrame): DataFrame com previsões detalhadas
        
    Returns:
        bool: True se salvou com sucesso, False caso contrário
    """
    if df_fornecedores is None or previsoes is None:
        return False
    
    try:
        # Data e hora no fuso de Brasília (GMT-3)
        agora = datetime.now(ZoneInfo("America/Sao_Paulo")).strftime('%d-%m-%Y %H-%M')

        # Salva na pasta atual
        caminho_arquivo = f'S:\\Procurement\\FUP\\IRF - Índice de Risco de Fornecedores\\Modelo de Machine Learning\\Histórico de Execuções\\IRF - {agora}.xlsx'
        print(f"💾 Salvando resultados localmente: {caminho_arquivo}")

        # Exporta para Excel com múltiplas abas
        with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
            df_fornecedores.to_excel(writer, sheet_name='Fornecedores', index=False)
            previsoes.to_excel(writer, sheet_name='Pedidos em Aberto', index=False)

        print("✅ Arquivo salvo com sucesso!")
        caminho_arquivo_url = caminho_arquivo.replace('\\', '/')
        print(f"📁 Arquivo disponível em: file:///{caminho_arquivo_url}")
        return True
    except Exception as e:
        print(f"❌ Erro ao salvar arquivo: {e}")
        return False

"""# Função principal"""

def main():
    """
    Função principal que executa todo o fluxo do IRF.
    
    Fluxo:
    1. Verifica arquivos na rede
    2. Carrega e processa dados
    3. Carrega modelo de ML
    4. Faz previsões
    5. Cria matriz de fornecedores
    6. Salva resultados
    """
    print("🚀 IRF - Índice de Risco de Fornecedor (Versão Rede)")
    print("=" * 60)
    
    # Verifica caminhos disponíveis na rede
    caminhos = verificar_caminhos()
    if caminhos is None:
        print("\n❌ Não foi possível encontrar todos os arquivos necessários na rede")
        print("   Verifique se você tem acesso aos seguintes caminhos:")
        print(f"   - {ARQUIVO_REDE}")
        print(f"   - {MODELO_REDE}")
        print(f"   - {CARGA_REDE}")
        return
    
    # Carrega dados
    df_pedidos_em_aberto = carregar_dados(caminhos)
    if df_pedidos_em_aberto is None:
        return
    
    # Processa dados
    df_pedidos_em_aberto = processar_dados(df_pedidos_em_aberto)
    if df_pedidos_em_aberto is None:
        return
    
    # Carrega modelo
    modelo = carregar_modelo(caminhos)
    if modelo is None:
        return
    
    # Faz previsões
    previsoes = fazer_previsoes(modelo, df_pedidos_em_aberto)
    if previsoes is None:
        return
    
    # Cria matriz de fornecedores
    df_fornecedores = criar_matriz_fornecedores(previsoes, caminhos)
    if df_fornecedores is None:
        return
    
    # Salva resultados
    if salvar_resultados(df_fornecedores, previsoes):
        print("\n🎉 Processamento concluído com sucesso!")
    else:
        print("\n❌ Erro ao salvar resultados")

if __name__ == "__main__":
    main() 