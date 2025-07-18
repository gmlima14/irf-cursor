# -*- coding: utf-8 -*-
"""
# Importação das Bibliotecas utilizadas
"""
# Importa as bibliotecas para manipular os dados
import pandas as pd
from pycaret.classification import *
from datetime import datetime
import numpy as np
import matplotlib.pyplot as plt

ARQUIVO_REDE = r'S:\Procurement\FUP\OTP Mensal\OTP - Base.xlsx'

def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")

def carregar_e_filtrar_dados(arquivo_rede):
    """
    Carrega e filtra os dados do arquivo Excel.
    
    Args:
        arquivo_rede (str): Caminho do arquivo Excel
        
    Returns:
        pandas.DataFrame: DataFrame filtrado com as colunas necessárias
    """    
    # Ler o arquivo excel
    log_message(f"📁 Carregando arquivo: {arquivo_rede}")
    df = pd.read_excel(arquivo_rede)
    log_message(f"✅ Arquivo carregado com {len(df)} registros iniciais")

    # Filtra apenas os dados que possuem Delivery Date
    if 'Delivery Date' in df.columns:
        df = df[df['Delivery Date'].notna()].copy()
    else:
        log_message("⚠️ Coluna 'Delivery Date' não encontrada no DataFrame")
    
    # # Filtra os registros cuja 'Delivery Date' seja de até 1 ano atrás em relação à data atual
    data_limite = datetime.today() - pd.DateOffset(years=1)
    # df['Delivery Date'] = pd.to_datetime(df['Delivery Date'], errors='coerce')
    df = df[df['Delivery Date'] >= data_limite].copy()
    
    # Manter apenas as colunas necessárias
    colunas_manter = ['BEDAT', 'Due Date (incl. ex works time)', 'MATKL', 'Vendor', 'NetOrderValue', 'On Time']
    df = df[colunas_manter].copy()
    
    log_message(f"✅ DataFrame filtrado com {len(df)} registros e {len(df.columns)} colunas")
    log_message(f"📊 Colunas mantidas: {list(df.columns)}")
    
    return df

def converter_datas_e_criar_variaveis_temporais(df):
    """
    Converte colunas de data e cria variáveis temporais.
    
    Args:
        df (pandas.DataFrame): DataFrame com os dados
        
    Returns:
        pandas.DataFrame: DataFrame com datas convertidas e variáveis temporais
    """    
    # Converter as colunas de data
    df["BEDAT"] = pd.to_datetime(df["BEDAT"], errors="coerce")  # Data de emissão
    df["Due Date (incl. ex works time)"] = pd.to_datetime(df["Due Date (incl. ex works time)"], errors="coerce")  # Entrega prevista

    # Dias para a entrega (em dias úteis)
    df["Dias Para Entrega"] = [
        np.busday_count(bedat.date(), due_date.date()) if pd.notnull(bedat) and pd.notnull(due_date) else np.nan
        for bedat, due_date in zip(df["BEDAT"], df["Due Date (incl. ex works time)"])
    ]
    
    # Inverte a coluna On Time
    df['On Time'] = df['On Time'].replace({1: 0, 0: 1})
    
    # Converter para categoria
    df['MATKL'] = df['MATKL'].astype('category')
    df['Vendor'] = df['Vendor'].astype('category')
    
    return df

def calcular_carga_fornecedor(df):
    """
    Calcula a carga do fornecedor para todos os registros.
    
    Args:
        df (pandas.DataFrame): DataFrame com os dados
        
    Returns:
        pandas.DataFrame: DataFrame com a coluna carga_fornecedor calculada
    """
    log_message("📈 Iniciando cálculo da carga do fornecedor...")
    
    # Inicializar nova coluna
    df['carga_fornecedor'] = 0
    
    df = df.sort_values(['Vendor', 'BEDAT']).reset_index(drop=True)

    df_start = df[['Vendor', 'BEDAT']].copy()
    df_start = df_start.rename(columns={'BEDAT':'date'})
    df_start['start'] = 1
    df_start['end'] = 0

    df_end = df[['Vendor', 'Due Date (incl. ex works time)']].copy()
    df_end = df_end.rename(columns={'Due Date (incl. ex works time)':'date'})
    df_end['start'] = 0
    df_end['end'] = 1

    df_events = pd.concat([df_start, df_end])
    df_events = df_events.sort_values(['Vendor', 'date']).reset_index(drop=True)

    df_events['open_pedidos'] = (
        df_events.groupby('Vendor')['start'].cumsum() - df_events.groupby('Vendor')['end'].cumsum()
    )

    df_open = df_events[['Vendor', 'date', 'open_pedidos']].drop_duplicates(subset=['Vendor', 'date'])

    df = pd.merge(df, df_open, left_on=['Vendor', 'BEDAT'], right_on=['Vendor', 'date'], how='left')

    df['carga_fornecedor'] = df['open_pedidos'] - 1

    df = df.drop(columns=['date', 'open_pedidos'])
    
    log_message("✅ Cálculo da carga do fornecedor concluído")
    return df

def treinar_e_salvar_modelo(df, caminho_salvamento):
    """
    Executa todo o pipeline de treinamento do modelo: prepara dados, configura experimento,
    treina modelo e salva o resultado.
    
    Args:
        df (pandas.DataFrame): DataFrame com os dados
        caminho_salvamento (str): Caminho onde salvar o modelo
        
    Returns:
        object: Modelo final treinado
    """
    log_message("🤖 Iniciando treinamento do modelo...")
    
    # Remove as colunas BEDAT e Due Date
    df = df.drop('BEDAT', axis=1)
    df = df.drop('Due Date (incl. ex works time)', axis=1)
    
    # Configura o experimento que será iniciado
    log_message("⚙️ Configurando experimento PyCaret...")
    s = setup(df, target='On Time', session_id=109, fix_imbalance=True)

    # Treina os três modelos
    log_message("Treinando modelo XGBoost...")
    modelo_xgb = create_model('xgboost')
    log_message("Treinando modelo LightGBM...")
    modelo_lgbm = create_model('lightgbm', verbose=False)
    log_message("Treinando modelo CatBoost...")
    modelo_cat = create_model('catboost')

    # Tunando os modelos
    log_message("🔧 Tunando modelo XGBoost...")
    modelo_xgb_tunado = tune_model(modelo_xgb, optimize='MCC', n_iter=5)
    log_message("🔧 Tunando modelo LightGBM...")
    modelo_lgbm_tunado = tune_model(modelo_lgbm, optimize='MCC', n_iter=5, verbose=False)
    log_message("🔧 Tunando modelo CatBoost...")
    modelo_cat_tunado = tune_model(modelo_cat, optimize='MCC', n_iter=5)

    # Plota a importância das variáveis para cada modelo
    plot_model(modelo_xgb_tunado, plot='feature', save=True)
    plot_model(modelo_lgbm_tunado, plot='feature', save=True)
    plot_model(modelo_cat_tunado, plot='feature', save=True)

    # Plota a matriz de confusão para cada modelo
    plot_model(modelo_xgb_tunado, plot='confusion_matrix', save=True)
    plot_model(modelo_lgbm_tunado, plot='confusion_matrix', save=True)
    plot_model(modelo_cat_tunado, plot='confusion_matrix', save=True)

    # Finaliza os modelos tunados
    log_message("🔧 Finalizando modelos tunados...")
    modelo_xgb_final = finalize_model(modelo_xgb_tunado)
    modelo_lgbm_final = finalize_model(modelo_lgbm_tunado)
    modelo_cat_final = finalize_model(modelo_cat_tunado)

    # Faz o blend dos modelos
    log_message("🤝 Fazendo blend dos modelos...")
    modelo_blend = blend_models(estimator_list=[modelo_xgb_final, modelo_lgbm_final, modelo_cat_final], fold=5, optimize='MCC', choose_better=True, verbose=False)
    modelo_blend_final = finalize_model(modelo_blend)

    # Salvando o modelo blendado
    log_message(f"💾 Salvando modelo blendado em: {caminho_salvamento}")
    save_model(modelo_blend_final, caminho_salvamento)

    log_message("✅ Modelo blendado treinado e salvo com sucesso!")
    return modelo_blend_final

def main():
    """
    Função principal que executa todo o pipeline de treinamento do modelo.
    """
    log_message("🚀 Iniciando pipeline de treinamento do modelo IRF...")
    log_message("=" * 60)
    
    # Carregar e filtrar dados
    df = carregar_e_filtrar_dados(ARQUIVO_REDE)
    
    # Converter datas e criar variáveis temporais
    df = converter_datas_e_criar_variaveis_temporais(df)
    
    # Calcular carga do fornecedor
    df = calcular_carga_fornecedor(df)

    # Salvar o dataframe com a carga do fornecedor
    df.to_excel(r'S:\Procurement\FUP\IRF - Índice de Risco de Fornecedores\Modelo de Machine Learning\dados_treinamento.xlsx', index=False)
    
    # Treinar e salvar modelo
    modelo_final = treinar_e_salvar_modelo(df, r'S:\Procurement\FUP\IRF - Índice de Risco de Fornecedores\Modelo de Machine Learning\modelo_treinado_blend_ytd')
    
    log_message("=" * 60)

if __name__ == "__main__":
    main()
