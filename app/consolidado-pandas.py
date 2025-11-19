import os
import pandas as pd
from datetime import date
import creds

# Caminhos das pastas de origem
itau_dir = creds.diretorio_consolidado_itau
caixa_dir = creds.diretorio_consolidado_caixa

# Caminho base de saída (pasta)
saida_dir = creds.diretorio_saida
os.makedirs(saida_dir, exist_ok=True)

# Data de hoje
data_hoje = date.today()
data_formatada = data_hoje.strftime("%Y_%m_%d")

# Nome do arquivo de saída
nome_arquivo = f"HISTORICO_CONSOLIDADO_{data_formatada}.xlsx"
saida = os.path.join(saida_dir, nome_arquivo)

# Lista apenas os arquivos HISTORICO-*.xlsx criados hoje
arquivos = []
for pasta in [itau_dir, caixa_dir]:
    if os.path.exists(pasta):
        for f in os.listdir(pasta):
            caminho = os.path.join(pasta, f)
            if (
                os.path.isfile(caminho)
                and f.lower().startswith("historico-")
                and f.lower().endswith(".xlsx")
            ):
                data_criacao = date.fromtimestamp(os.path.getctime(caminho))
                if data_criacao == data_hoje:
                    arquivos.append(caminho)

print(f"📂 Arquivos HISTORICO criados hoje ({data_hoje}):")
for a in arquivos:
    print(" -", a)

# Lê e concatena todos os arquivos HISTORICO
dfs = []
for f in arquivos:
    print(f"\n📄 Lendo {os.path.basename(f)}")
    df = pd.read_excel(f)

    # 🔹 Corrige colunas duplicadas (se existirem)
    df = df.loc[:, ~df.columns.duplicated()].copy()

    # 🔹 Detecta o banco
    banco_nome = ""
    if "banco" in df.columns and not df["banco"].empty:
        banco_nome = str(df["banco"].iloc[0]).strip().upper()

    # 🔹 Remove colunas desnecessárias por banco
    if "ITAU" in banco_nome:
        colunas_excluir = ["Razão Social", "CPF/CNPJ"]
    elif "CAIXA" in banco_nome:
        colunas_excluir = ["ag./origem"]
    else:
        colunas_excluir = []

    for c in colunas_excluir:
        if c in df.columns:
            df.drop(columns=c, inplace=True, errors="ignore")
            print(f"🧹 Coluna '{c}' removida ({banco_nome})")

    # 🔹 Padronizar as colunas finais (mesma ordem e nomes)
    colunas_padrao = [
        "data", "lançamento", "valor (R$)", "saldo (R$)",
        "nome", "banco", "agencia", "conta", "data_atualizada"
    ]

    # Renomeia se necessário
    renomear = {
        "Data": "data",
        "Lançamento": "lançamento",
        "Valor (R$)": "valor (R$)",
        "Saldo (R$)": "saldo (R$)"
    }
    df.rename(columns=renomear, inplace=True)

    # Adiciona colunas ausentes com valor nulo
    for col in colunas_padrao:
        if col not in df.columns:
            df[col] = None

    # Reordena e remove duplicatas
    df = df[colunas_padrao]
    df = df.loc[:, ~df.columns.duplicated()].copy()

    # Reseta o índice
    df.reset_index(drop=True, inplace=True)

    dfs.append(df)
    print(f"✅ Banco detectado: {banco_nome} | Colunas padronizadas.")

# 🔹 Junta todos os DataFrames em um só
if dfs:
    df_final = pd.concat(dfs, ignore_index=True)
    df_final.reset_index(drop=True, inplace=True)
    df_final = df_final.loc[:, ~df_final.columns.duplicated()]

    # 🔹 Normaliza coluna de data para datetime com hora
    df_final["data"] = pd.to_datetime(df_final["data"], errors="coerce")
    df_final["data"] = df_final["data"].apply(
        lambda d: pd.Timestamp(d) if not pd.isna(d) else None
    )

    # 🔹 Salva no Excel com formatação de data real
    df_final.to_excel(saida, index=False, engine="openpyxl")

    # 🔹 Log final
    total_itau = len(df_final[df_final["banco"].str.upper() == "ITAU"])
    total_caixa = len(df_final[df_final["banco"].str.upper() == "CAIXA"])
    total_geral = len(df_final)

    print(f"\n✅ Registros Itaú: {total_itau}")
    print(f"✅ Registros Caixa: {total_caixa}")
    print(f"📊 Total consolidado: {total_geral}")
    print(f"\n💾 Arquivo consolidado salvo em:\n{saida}")

else:
    print("⚠️ Nenhum arquivo HISTORICO encontrado para hoje.")
