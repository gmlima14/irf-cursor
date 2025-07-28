# IRF - Índice de Risco de Fornecedor

## 🚀 Versão Atual

A versão atual (`irf.py`) foi simplificada para utilizar **apenas um modelo de machine learning já treinado** (modelo blendado ou LightGBM, conforme definido em `MODELO_BLEND`).

### 📁 **Arquivos de Entrada (Rede):**

- **Dados:** `S:\Procurement\FUP\OTP Mensal\OTP - Base.xlsx`
- **Modelo:** `S:\Procurement\FUP\IRF - Índice de Risco de Fornecedores\Modelo de Machine Learning\modelo_treinado_lightgbm.pkl` *(ou modelo_treinado_blend.pkl, conforme configuração)*
- **Carga:** `S:\Procurement\FUP\IRF - Índice de Risco de Fornecedores\Modelo de Machine Learning\carga_fornecedor.csv`

### 💾 **Local de Salvamento:**
- Salva na pasta de histórico da rede: `S:\Procurement\FUP\IRF - Índice de Risco de Fornecedores\Modelo de Machine Learning\Histórico de Execuções\`

## 🔧 **Funções do Sistema**

### **1. verificar_caminhos()**
- **Função:** Verifica se todos os arquivos necessários estão disponíveis na rede
- **Retorna:** Dicionário com caminhos ou None se arquivo não encontrado

### **2. carregar_dados(caminhos)**
- **Função:** Carrega dados do Excel e faz limpeza inicial
- **Ações:** Remove linhas com 'Delivery Date' preenchida e coluna 'On Time'
- **Retorna:** DataFrame limpo ou None se erro

### **3. processar_dados(df_pedidos_em_aberto)**
- **Função:** Processa dados para análise de machine learning
- **Ações:** Converte tipos, calcula variáveis de tempo, mapeia carga por fornecedor
- **Retorna:** DataFrame processado com variáveis calculadas

### **4. carregar_modelo(caminho_modelo)**
- **Função:** Carrega o modelo de machine learning treinado (apenas um modelo, já blendado ou LightGBM)
- **Retorna:** Modelo carregado ou None se erro

### **5. fazer_previsoes(modelo, df_pedidos_em_aberto)**
- **Função:** Faz previsões de atraso usando o modelo
- **Ações:** Executa predições, renomeia colunas, converte valores (0/1 → No Prazo/Atraso)
- **Retorna:** DataFrame com previsões e confiabilidade

### **6. criar_matriz_fornecedores(previsoes, df_carga, caminhos)**
- **Função:** Cria matriz de fornecedores com índice de risco
- **Ações:** Agrupa por fornecedor, calcula taxas, aplica fórmula do índice de risco
- **Retorna:** DataFrame com ranking de fornecedores

### **7. salvar_resultados(df_fornecedores, previsoes)**
- **Função:** Salva resultados em Excel na pasta de histórico
- **Ações:** Cria arquivo com duas abas (Fornecedores e Pedidos em Aberto)
- **Retorna:** True se sucesso, False se erro

### **8. main()**
- **Função:** Função principal que executa todo o fluxo
- **Fluxo:** Verifica → Carrega → Processa → Carrega Modelo → Prevê → Cria Matriz → Salva

## Como executar o código

### Opção 1: Script Batch (Windows)
1. Clique duas vezes no arquivo `run_previsao.bat`
2. O script irá automaticamente:
   - Ativar o ambiente virtual
   - Executar o código Python
   - Aguardar você pressionar uma tecla para sair

### Opção 2: Execução manual
Se preferir executar manualmente:

```bash
# Ativar o ambiente virtual
.venv\Scripts\activate

# Executar o código
python irf.py
```

## 🔍 **Como funciona:**

1. **Verificação da rede:** O código verifica se todos os arquivos estão disponíveis na rede
2. **Mensagens claras:** Mostra exatamente onde encontrou cada arquivo
3. **Salvamento organizado:** Salva na pasta de histórico com data/hora
4. **Tratamento de erros:** Mensagens claras se algum arquivo não for encontrado
5. **Documentação completa:** Todas as funções têm descrições detalhadas

## Dependências
Certifique-se de que todas as dependências estão instaladas:
```bash
pip install -r requirements.txt
```

## Notas
- O ambiente virtual já está configurado na pasta `.venv`
- Os scripts automaticamente ativam o venv antes de executar o código
- O arquivo de saída será salvo com o nome `IRF - [data_hora].xlsx`
- **Requisito:** Acesso à rede da empresa para os arquivos
- **Documentação:** Todas as funções incluem docstrings com parâmetros e retornos
- Se algum arquivo não for encontrado na rede, o código mostrará uma mensagem clara

## 🚨 Problemas Resolvidos

- ❌ **Erro de caminho de rede**: Código tentava acessar arquivos em `S:\Procurement\...`
- ❌ **Erro de linter**: Problemas de tipagem no pandas
- ❌ **Erro de ExcelWriter**: Problema com o engine xlsxwriter

## ✅ Versão Corrigida

- ✅ **Arquivos locais**: Procura automaticamente arquivos na pasta do projeto
- ✅ **Tratamento de erros**: Mensagens claras quando arquivos estão faltando
- ✅ **Código limpo**: Sem erros de linter
- ✅ **ExcelWriter corrigido**: Usa openpyxl em vez de xlsxwriter

## Observações Finais
- Agora o sistema utiliza **apenas um modelo já treinado** (blendado ou LightGBM, conforme definido em `MODELO_BLEND` no início do código).
- Não é mais necessário manter múltiplos arquivos de modelo para previsão.
- O fluxo está mais simples, robusto e fácil de manter.
- Se precisar treinar um novo modelo blendado, utilize o script de treinamento apropriado e substitua o arquivo de modelo na pasta da rede. 