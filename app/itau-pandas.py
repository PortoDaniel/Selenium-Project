# PANDAS E CRIAÇÃO DE ARQUIVOS CONSOLIDADOS
import pandas as pd
import os, shutil, datetime
import creds

# Diretórios
diretorio_base = creds.diretorio_base_itau
diretorio_file_base = creds.diretorio_file_base_itau
diretorio_consolidado = creds.diretorio_consolidado_itau

data_hoje = datetime.date.today()

# Lista de arquivos do dia
arquivos_xlsx = [
    os.path.join(diretorio_file_base, f)
    for f in os.listdir(diretorio_file_base)
    if f.lower().endswith((".xls", ".xlsx"))
    and datetime.date.fromtimestamp(os.path.getctime(os.path.join(diretorio_file_base, f))) == data_hoje
]

print(f"\n🔎 {len(arquivos_xlsx)} arquivos encontrados na pasta 02-FILES\n")

dfs = []
for i, f in enumerate(arquivos_xlsx, start=1):
    print(f"📂 [{i}/{len(arquivos_xlsx)}] Lendo arquivo: {os.path.basename(f)}")

    try:
        df = pd.read_excel(f, header=None, engine="openpyxl")
        print(f"   ✅ Arquivo carregado - Shape: {df.shape}")

        # 🧠 Detectar automaticamente a linha do cabeçalho
        header_row = None
        for idx, row in df.iterrows():
            if any(str(cell).strip().lower() == "data" for cell in row):
                header_row = idx
                break

        if header_row is None:
            print("   ⚠️ Cabeçalho não encontrado, pulando arquivo.\n")
            continue

        # Extrair metadados
        nome_xls = df.iloc[2, 1]
        agencia_xls = df.iloc[3, 1]
        conta_xls = df.iloc[4, 1]
        data_atualizada_xls = str(df.iloc[1, 1]).split()[0]

        print(f"   🏦 Nome: {nome_xls} | Agência: {agencia_xls} | Conta: {conta_xls} | Data: {data_atualizada_xls}")

        # Aplicar o cabeçalho correto
        df.columns = df.iloc[header_row]
        df = df.iloc[header_row + 1:].reset_index(drop=True)

        # Adicionar metadados fixos
        df["nome"] = nome_xls
        df["agencia"] = agencia_xls
        df["banco"] = "ITAU"
        df["conta"] = conta_xls
        df["data_atualizada"] = data_atualizada_xls

        dfs.append(df)
        print("   ➕ DataFrame adicionado à lista de histórico\n")

    except Exception as e:
        print(f"   ❌ Erro ao processar {os.path.basename(f)}: {e}\n")

# 🧩 Consolidar DataFrames
if not dfs:
    print("❌ Nenhum arquivo válido encontrado. Encerrando.")
    exit()

historico = pd.concat(dfs, ignore_index=True)
print(f"\n📊 Histórico consolidado - Total de linhas: {len(historico)}")

# 🔹 Detectar nome da coluna de data
col_data = next((c for c in historico.columns if "data" in str(c).lower()), None)
if col_data:
    historico["data"] = pd.to_datetime(historico[col_data], errors="coerce", dayfirst=True)
    print(f"📅 Coluna de data identificada: '{col_data}'")
else:
    print("⚠️ Nenhuma coluna de data encontrada. Pulando conversão de datas.")

# 🔹 Detectar nome da coluna de lançamento
col_lanc = next((c for c in historico.columns if "lan" in str(c).lower()), None)
if col_lanc:
    filtro = historico[historico[col_lanc].astype(str).str.contains("SALDO", na=False, case=False)]
else:
    print("⚠️ Nenhuma coluna de lançamento encontrada. Pulando filtro de saldos.")
    filtro = historico.copy()

# 🔹 Filtrar saldos
saldo_total = (
    filtro[filtro[col_lanc].str.contains("SALDO TOTAL", case=False, na=False)]
    .sort_values("data")
    .groupby("conta", as_index=False)
    .tail(1)
)

saldo_anterior = filtro[filtro[col_lanc].str.contains("SALDO ANTERIOR", case=False, na=False)]

resultado = (
    pd.concat([saldo_total, saldo_anterior])
    .drop_duplicates(subset=["conta"], keep="first")
    .reset_index(drop=True)
)

# 🧾 Salvar arquivos
os.makedirs(diretorio_consolidado, exist_ok=True)
arquivo_hist = os.path.join(diretorio_consolidado, f"HISTORICO-{pd.Timestamp.today().strftime('%d_%m_%Y')}.xlsx")
arquivo_cons = os.path.join(diretorio_consolidado, f"CONSOLIDADO-{pd.Timestamp.today().strftime('%d_%m_%Y')}.xlsx")

historico.to_excel(arquivo_hist, index=False, engine="openpyxl")
print(f"💾 Histórico salvo em: {arquivo_hist}")

resultado.to_excel(arquivo_cons, index=False, engine="openpyxl")
print(f"💾 Consolidado salvo em: {arquivo_cons}\n")

print("✅ Processo finalizado com sucesso!")
