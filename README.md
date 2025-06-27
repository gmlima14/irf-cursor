# IRF - Índice de Risco de Fornecedor

## 🚀 Versão Atual

A versão atual `irf.py` foi desenvolvida para usar **apenas arquivos da rede** e inclui documentação completa das funções.

### 📁 **Arquivos de Entrada (Rede):**

- **Dados:** `S:\Procurement\FUP\OTP Mensal\OTP - Base.xlsm`
- **Modelo:** `S:\Procurement\FUP\IRF - Índice de Risco de Fornecedores\Modelo de Machine Learning\modelo_treinado_alldates.pkl`
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

### **4. carregar_modelo(caminhos)**
- **Função:** Carrega modelo de machine learning treinado
- **Retorna:** Modelo carregado ou None se erro

### **5. fazer_previsoes(modelo, df_pedidos_em_aberto)**
- **Função:** Faz previsões de atraso usando o modelo
- **Ações:** Executa predições, renomeia colunas, converte valores (0/1 → No Prazo/Atraso)
- **Retorna:** DataFrame com previsões e confiabilidade

### **6. criar_matriz_fornecedores(previsoes, caminhos)**
- **Função:** Cria matriz de fornecedores com índice de risco
- **Ações:** Agrupa por fornecedor, calcula taxas, aplica fórmula do índice de risco
- **Retorna:** DataFrame com ranking de fornecedores

### **7. salvar_resultados(df_fornecedores, previsoes)**
- **Função:** Salva resultados em Excel na pasta de histórico
- **Ações:** Cria arquivo com duas abas (Fornecedores e Pedidos em Aberto)
- **Retorna:** True se sucesso, False se erro

### **8. main()**
- **Função:** Função principal que executa todo o fluxo
- **Fluxo:** Verifica → Carrega → Processa → Prevê → Cria Matriz → Salva

## Como executar o código com ativação automática do ambiente virtual

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

O código original tinha alguns problemas que foram corrigidos na versão `irf_corrigido.py`:

- ❌ **Erro de caminho de rede**: Código tentava acessar arquivos em `S:\Procurement\...`
- ❌ **Erro de linter**: Problemas de tipagem no pandas
- ❌ **Erro de ExcelWriter**: Problema com o engine xlsxwriter

## ✅ Versão Corrigida

A versão `irf_corrigido.py` resolve todos esses problemas:

- ✅ **Arquivos locais**: Procura automaticamente arquivos na pasta do projeto
- ✅ **Tratamento de erros**: Mensagens claras quando arquivos estão faltando
- ✅ **Código limpo**: Sem erros de linter
- ✅ **ExcelWriter corrigido**: Usa openpyxl em vez de xlsxwriter

## 🚀 Versão Rede/OneDrive

A versão `irf_rede_oneprive.py` foi criada especificamente para usar arquivos da rede e OneDrive conforme solicitado:

### 📁 **Arquivos de Entrada:**

**Dados de Pedidos:**
- **Rede:** `S:\Procurement\FUP\OTP Mensal\OTP - Base.xlsm`
- **OneDrive:** `C:\Users\CSUGAB01\OneDrive - ANDRITZ AG\General - Follow up - Araucária\OTP - Base.xlsm`

**Modelo e Carga (sempre da rede):**
- **Modelo:** `S:\Procurement\FUP\IRF - Índice de Risco de Fornecedores\Modelo de Machine Learning\modelo_treinado_alldates.pkl`
- **Carga:** `S:\Procurement\FUP\IRF - Índice de Risco de Fornecedores\Modelo de Machine Learning\carga_fornecedor.csv`

### 💾 **Local de Salvamento:**
- Se usar arquivo da **rede** → Salva na pasta atual
- Se usar arquivo do **OneDrive** → Salva no OneDrive: `C:\Users\CSUGAB01\OneDrive - ANDRITZ AG\General - Follow up - Araucária\`

## Como executar o código com ativação automática do ambiente virtual

### Opção 1: Script Batch (Windows)
1. Clique duas vezes no arquivo `run_previsao.bat`
2. O script irá automaticamente:
   - Ativar o ambiente virtual
   - Executar o código Python (versão rede/OneDrive)
   - Aguardar você pressionar uma tecla para sair

### Opção 2: Script PowerShell (Windows)
1. Abra o PowerShell
2. Navegue até a pasta do projeto
3. Execute: `.\run_previsao.ps1`
4. O script irá automaticamente:
   - Ativar o ambiente virtual
   - Executar o código Python (versão rede/OneDrive)
   - Aguardar você pressionar uma tecla para sair

### Opção 3: Execução manual
Se preferir executar manualmente:

```bash
# Ativar o ambiente virtual
.venv\Scripts\activate

# Executar o código que usa rede/OneDrive
python irf_rede_oneprive.py
```

### Opção 4: Comando único (PowerShell)
```powershell
.\ativar_e_executar.ps1
```

## 🔍 **Como funciona:**

1. **Verificação automática:** O código verifica primeiro se o arquivo está na rede, depois no OneDrive
2. **Mensagens claras:** Mostra exatamente onde encontrou cada arquivo
3. **Salvamento inteligente:** Salva no local apropriado baseado na origem dos dados
4. **Tratamento de erros:** Mensagens claras se algum arquivo não for encontrado

## 📁 Arquivos Necessários

Para que o código funcione, você precisa ter na pasta do projeto:

1. **`modelo_treinado_alldates.pkl`** - Modelo de machine learning treinado
2. **`carga_fornecedor.csv`** - Dados de carga dos fornecedores
3. **Arquivo Excel** - Com pedidos em aberto (qualquer arquivo .xlsx ou .xls)

## Dependências
Certifique-se de que todas as dependências estão instaladas:
```bash
pip install -r requirements.txt
```

## Notas
- O ambiente virtual já está configurado na pasta `.venv`
- Os scripts automaticamente ativam o venv antes de executar o código
- O arquivo de saída será salvo com o nome `IRF - [data_hora].xlsx`
- **Recomendação:** Use o OneDrive para melhor sincronização com o Teams
- Se algum arquivo não for encontrado, o código mostrará uma mensagem clara
- A versão corrigida procura automaticamente os arquivos necessários na pasta
- Se algum arquivo estiver faltando, o código mostrará uma mensagem clara do que precisa ser adicionado 