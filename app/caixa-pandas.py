import pandas as pd
import os, time, shutil
import re
from datetime import datetime, date
import creds
# Diretórios
diretorio_base = creds.diretorio_base_caixa
diretorio_file_base = creds.diretorio_file_base_caixa
diretorio_consolidado = creds.diretorio_consolidado_caixa

data_hoje = date.today()

# 🧹 Excluir apenas arquivos e subpastas criadas hoje no diretório consolidado
if os.path.exists(diretorio_consolidado):
    for arquivo in os.listdir(diretorio_consolidado):
        caminho_arquivo = os.path.join(diretorio_consolidado, arquivo)
        try:
            # Obtém a data de criação
            data_criacao = date.fromtimestamp(os.path.getctime(caminho_arquivo))

            # Só remove se o arquivo/pasta for de hoje
            if data_criacao == data_hoje:
                if os.path.isfile(caminho_arquivo) or os.path.islink(caminho_arquivo):
                    os.unlink(caminho_arquivo)
                    print(f"🗑️ Arquivo removido: {arquivo}")
                elif os.path.isdir(caminho_arquivo):
                    shutil.rmtree(caminho_arquivo)
                    print(f"🗑️ Pasta removida: {arquivo}")

        except Exception as e:
            print(f"⚠️ Erro ao excluir {caminho_arquivo}: {e}")
else:
    os.makedirs(diretorio_consolidado)
    print(f"📁 Diretório criado: {diretorio_consolidado}")

print("✅ Limpeza concluída: apenas arquivos criados hoje foram removidos.")

# Lista de arquivos .xls
arquivos_xls = [
    os.path.join(diretorio_file_base, f)
    for f in os.listdir(diretorio_file_base)
    if f.lower().endswith((".xls", ".xlsx"))
    and date.fromtimestamp(os.path.getctime(os.path.join(diretorio_file_base, f))) == data_hoje
]

print(f"\n🔎 {len(arquivos_xls)} arquivos criados hoje encontrados na pasta 02-FILES\n")
for a in arquivos_xls:
    print(" -", os.path.basename(a))

dfs = []
for i, f in enumerate(arquivos_xls, start=1):
    print(f"📂 [{i}/{len(arquivos_xls)}] Lendo arquivo: {os.path.basename(f)}")

    # 🔹 Extrai agência e conta do nome do arquivo
    match = re.search(r"AG(\d+)_CC(\d+-\d+)", os.path.basename(f))
    if match:
        agencia = match.group(1)
        conta = match.group(2)
    else:
        agencia = None
        conta = None

    print(f"🏦 Agência: {agencia}")
    print(f"💳 Conta: {conta}")

    # 🔹 Lê o Excel (pulando a primeira linha "Extrato de ...")
    try:
        df = pd.read_excel(f, header=None, skiprows=2, engine="xlrd")
        print(f"   ✅ Arquivo carregado via xlrd - Shape: {df.shape}")

    except Exception as e:
        print(f"⚠️ Erro ao ler com xlrd ({e}). Tentando método alternativo com read_html...")
        try:
            df = pd.read_html(f, header=None, skiprows=2, decimal=",", thousands=".")[0]
            print(f"   ✅ Arquivo carregado via read_html - Shape: {df.shape}")
        except Exception as e2:
            print(f"❌ Falha ao ler o arquivo {os.path.basename(f)} com qualquer método ({e2}). Pulando...\n")
            continue

# 🔸 Verifica se o arquivo está vazio
    if df.empty or df.shape[0] == 0:
        print(f"⚠️ Arquivo vazio detectado ({os.path.basename(f)}). Nenhuma movimentação. Pulando...\n")
        continue

    # 🔹 Define colunas originais
    df.columns = [
        "Data Lançamento",
        "Data Movimento",
        "Histórico",
        "Documento",
        "Valor Lançamento",
        "Saldo",
        "CPF/CNPJ",
        "Nome/Razão Social"
    ]

    # 🔹 Renomeia colunas conforme solicitado
    df.rename(columns={
        "Data Lançamento": "data",
        "Histórico": "lançamento",
        "Documento": "ag./origem",
        "Valor Lançamento": "valor (R$)",
        "Saldo": "saldo (R$)"
    }, inplace=True)

    # 🔹 Remove
    df.drop(columns=["Data Movimento"], inplace=True)
    df.drop(columns=["Nome/Razão Social"], inplace=True)
    df.drop(columns=["CPF/CNPJ"], inplace=True)

    # 🔹 Cria colunas adicionais
    df["nome"] = "7LM EMPREENDIMENTOS - CAIXA"
    df["banco"] = "CAIXA"
    df["agencia"] = agencia
    df["conta"] = conta
    df['data_atualizada'] = datetime.now().strftime('%d/%m/%Y')


    # 🔹 Adiciona à lista de DataFrames
    dfs.append(df)
    print("   ➕ DataFrame adicionado à lista de histórico\n")


# 🔹 Junta todos os DataFrames em um só
df_consolidado = pd.concat(dfs, ignore_index=True)

# 🔹 Cria o nome do arquivo com data/hora
arquivo_consolidado = os.path.join(
    diretorio_consolidado,
    f"CONSOLIDADO-{pd.Timestamp.today().strftime('%d_%m_%Y')}.xlsx"
)

# 🔹 Exporta para Excel
df_consolidado.to_excel(arquivo_consolidado, index=False, engine="openpyxl")

print(f"💾 Arquivo consolidado salvo em:\n{arquivo_consolidado}")

#HISTORICO

df_historico = pd.concat(dfs, ignore_index=True)

# 🔹 Ordena do mais recente para o mais antigo
df_historico = df_historico.sort_values("data", ascending=False)

# 🔹 Cria um novo DataFrame apenas com a linha mais recente de cada conta
df_historico = df_historico.groupby(["agencia", "conta"], as_index=False).first()

arquivo_historico = os.path.join(
    diretorio_consolidado,
    f"HISTORICO-{pd.Timestamp.today().strftime('%d_%m_%Y')}.xlsx"
)

# 🔹 Exporta para Excel
df_historico.to_excel(arquivo_historico, index=False, engine="openpyxl")