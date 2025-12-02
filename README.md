# 💼 Automação Bancária — Itaú & Caixa  
### Selenium + Pandas + Orquestração Automática

Este projeto automatiza o processo de:

✔ Login no Itaú e na Caixa  
✔ Download automático de extratos bancários  
✔ Organização das pastas e limpeza de arquivos  
✔ Leitura e padronização dos extratos com Pandas  
✔ Geração de histórico consolidado diário  
✔ Execução completa orquestrada via `run_all.py`

O objetivo é eliminar atividades manuais do financeiro e garantir que todos os extratos do dia sejam baixados, processados e consolidados automaticamente.

---

# 📌 1. Estrutura Geral do Projeto

```
/app
│   itau-main.py
│   itau-pandas.py
│   caixa-main.py
│   caixa-pandas.py
│   consolidado-pandas.py
│   run_all.py
│   orchestrator.log
│   start_app.bat
│
└── creds.py       ← (criado pelo usuário – NÃO está no Git)
```

---

# 🔐 2. Sobre o arquivo `creds.py` (obrigatório)

Este arquivo **NÃO é distribuído** no repositório porque contém:

- credenciais bancárias  
- diretórios internos  
- dados sensíveis  

Você deve criar o seu próprio `creds.py` dentro da pasta `/app`.

## ➤ Modelo de `creds.py`

```python
import os

# ===============================
# CREDENCIAIS CAIXA
# ===============================
ppswd_caixa = "SUA_SENHA_CAIXA"
accont_caixa = "SEU_USUARIO_CAIXA"

# ===============================
# CREDENCIAIS ITAÚ
# ===============================
accont_itau = "SEU_OPERADOR_ITAU"
ppswd_itau = "SUA_SENHA_ITAU"

# ===============================
# DIRETÓRIOS
# ===============================

# ITAÚ — diretórios da estrutura interna
diretorio_base_itau = r"C:\CAMINHO\ITAU"
diretorio_file_base_itau = os.path.join(diretorio_base_itau, "02-FILES")
diretorio_consolidado_itau = os.path.join(diretorio_base_itau, "01-CONSOLIDADO")

# CAIXA — diretórios da estrutura interna
diretorio_base_caixa = r"C:\CAMINHO\CAIXA"
diretorio_file_base_caixa = os.path.join(diretorio_base_caixa, "02-FILES")
diretorio_consolidado_caixa = os.path.join(diretorio_base_caixa, "01-CONSOLIDADO")

# CONSOLIDADO GERAL (Itaú + Caixa)
diretorio_saida = r"C:\CAMINHO\CONSOLIDADO_GERAL"
```

---

# 🧰 3. Instalação das Dependências

### Criar ambiente virtual (opcional, recomendado)

```bash
python -m venv .venv
.venv\Scripts\activate
```

### Instalar bibliotecas

```bash
pip install -r requirements.txt
```

Ou manualmente:

```bash
pip install selenium pandas xlrd openpyxl python-dotenv
```

---

# ▶️ 4. Como Rodar o Projeto

### 🔥 Rodar tudo automaticamente (modo recomendado)

```bash
python run_all.py
```

O orquestrador:

1. Executa o Itaú (download de extratos)  
2. Processa os arquivos Itaú  
3. Executa a Caixa (download de extratos)  
4. Processa os arquivos Caixa  
5. Gera o consolidado diário  

Log completo é salvo em:

```
app/orchestrator.log
```

---

# ▶️ 5. Execução manual (script por script)

### Itaú — baixar extratos
```bash
python itau-main.py
```

### Itaú — processar extratos
```bash
python itau-pandas.py
```

### Caixa — baixar extratos
```bash
python caixa-main.py
```

### Caixa — processar extratos
```bash
python caixa-pandas.py
```

### Gerar consolidado geral
```bash
python consolidado-pandas.py
```

---

# 🧠 6. Como funciona a automação

### 📌 1. `itau-main.py`
- Abre o site do Itaú  
- Realiza login como operador  
- Navega pelas contas  
- Baixa o extrato em Excel  
- Salva em `02-FILES`  

### 📌 2. `itau-pandas.py`
- Identifica arquivos baixados no dia  
- Padroniza colunas  
- Converte datas  
- Cria colunas adicionais  
- Salva em `01-CONSOLIDADO`  

### 📌 3. `caixa-main.py`
- Abre o site da Caixa com link direto de login  
- Preenche senha  
- Abre todas as contas  
- Baixa extratos  
- Salva em `02-FILES`  

### 📌 4. `caixa-pandas.py`
- Realiza tratamento e padronização  
- Salva em `01-CONSOLIDADO`  

### 📌 5. `consolidado-pandas.py`
- Junta Itaú + Caixa  
- Exclui duplicidades  
- Padroniza colunas  
- Gera arquivo final:

```
HISTORICO_CONSOLIDADO_YYYY_MM_DD.xlsx
```

---

# 📋 7. Execução automatizada pelo Windows (opcional)

Você pode agendar o Windows Scheduler para rodar o arquivo:

```
start_app.bat
```

Que contém:

```bat
cd /d "C:\CAMINHO\app"
call .venv\Scripts\activate
python run_all.py
```

---

# 📄 8. Logs

Tudo é logado com timestamps:

```
app/orchestrator.log
```

Inclui:

- início/fim de cada script  
- erros detalhados  
- tempo de execução  
- status geral  

---

# 🔒 9. Segurança

- `creds.py` está 100% ignorado pelo Git  
- Nenhuma credencial é enviada ao repositório  

---

# 📄 10. Licença

Uso interno (Financeiro / TI).  
Livre para manutenção e melhorias internas.
