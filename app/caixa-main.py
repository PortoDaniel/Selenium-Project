from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time, os, shutil, glob
import creds
from datetime import date
from selenium.common.exceptions import TimeoutException

# Diretórios
diretorio_base = creds.diretorio_base_caixa
diretorio_file_base = creds.diretorio_file_base_caixa
diretorio_consolidado = creds.diretorio_consolidado_caixa

# Data atual
data_hoje = date.today()

# Limpeza de pastas — apenas arquivos criados hoje
for pasta in [diretorio_file_base, diretorio_consolidado]:
    os.makedirs(pasta, exist_ok=True)
    for arquivo in os.listdir(pasta):
        caminho = os.path.join(pasta, arquivo)

        # Ignora qualquer caminho que contenha 'ITAU' (maiúsculo ou minúsculo)
        if "ITAU" in caminho.upper():
            continue

        try:
            if os.path.isfile(caminho):
                data_criacao = date.fromtimestamp(os.path.getctime(caminho))
                if data_criacao == data_hoje:
                    os.remove(caminho)
                    print(f"🗑️ Arquivo removido: {arquivo}")
            elif os.path.isdir(caminho):
                data_criacao = date.fromtimestamp(os.path.getctime(caminho))
                if data_criacao == data_hoje:
                    shutil.rmtree(caminho)
                    print(f"🗑️ Pasta removida: {arquivo}")
        except Exception as e:
            print(f"⚠️ Erro ao remover {arquivo}: {e}")

# Configuração do Firefox
options = webdriver.FirefoxOptions()
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.dir", diretorio_file_base)
options.set_preference("browser.helperApps.neverAsk.saveToDisk",
    "application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream")
options.set_preference("browser.download.manager.showWhenStarting", False)
options.set_preference("browser.download.useDownloadDir", True)
options.set_preference("pdfjs.disabled", True)
options.set_preference("browser.download.alwaysOpenPanel", False)
options.set_preference("permissions.default.geo", 1)
options.set_preference("geo.prompt.testing", True)
options.set_preference("geo.prompt.testing.allow", True)

pswd = creds.ppswd_caixa
accont = creds.accont_caixa

driver = webdriver.Firefox(service=Service(), options=options)
driver.get(f"https://loginx.caixa.gov.br/auth/realms/r_inter_siper/protocol/openid-connect/auth?client_id=cli-web-gce&login_hint={accont}&redirect_uri=https%3A%2F%2Fgerenciador.caixa.gov.br%2Fempresa&response_mode=query&response_type=code&scope=openid&nonce=b117bf80c4384ef3637f5a0a8042d410c580a5f86119198a5dbd21a2bbc0c1c9")

wait = WebDriverWait(driver, 20)

# Login
wait.until(EC.presence_of_element_located((By.ID, "password"))).send_keys(pswd)
driver.find_element(By.ID, "kc-login").click()

time.sleep(2)
# Fechar aviso inicial, se aparecer
wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "msg-close"))).click()
time.sleep(2)

def houve_bloqueio_saldo(driver, timeout=3):
    """
    True se aparecer a mensagem de bloqueio de saldo.
    Usa o <p class='mensagem__descricao'>...permissão de acesso ao saldo...</p>
    """
    xpath_msg = ("//p[contains(@class,'mensagem__descricao') "
                 "and contains(normalize-space(.), 'permissão de acesso ao saldo')]")
    try:
        WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((By.XPATH, xpath_msg))
        )
        return True
    except TimeoutException:
        return False

def esperar_extrato_ou_bloqueio(driver, timeout=10):
    """
    Espera EITHER: (a) a mensagem de bloqueio OU (b) algum conteúdo do painel de extrato.
    Retorna 'bloqueio' ou 'ok'.
    """
    xpath_bloqueio = ("//p[contains(@class,'mensagem__descricao') "
                      "and contains(normalize-space(.), 'permissão de acesso ao saldo')]")
    # Ajuste o seletor 'ok' para algo que sempre exista quando o extrato carrega
    xpath_extrato_ok = "//a[contains(normalize-space(.), 'Extrato CAIXA')]"

    end = time.time() + timeout
    while time.time() < end:
        # bloqueio?
        if driver.find_elements(By.XPATH, xpath_bloqueio):
            return "bloqueio"
        # extrato disponível?
        if driver.find_elements(By.XPATH, xpath_extrato_ok):
            return "ok"
        time.sleep(0.25)
    # Timeout: tenta uma última checagem
    return "bloqueio" if driver.find_elements(By.XPATH, xpath_bloqueio) else "ok"

# Loop das contas
while True:
    # Abre o seletor de contas
    abrir_contas = wait.until(EC.element_to_be_clickable((By.ID, "selectConta")))
    abrir_contas.click()

    lista = wait.until(EC.presence_of_element_located((By.ID, "listaContas")))
    contas = lista.find_elements(By.TAG_NAME, "li")
    total_contas = len(contas)

    print(f"🔢 {total_contas} contas encontradas.\n")
    ActionChains(driver).move_by_offset(5, 5).click().perform()
    time.sleep(2)
    for i in range(total_contas):
        # Reabre o seletor e recarrega a lista (pois o DOM é recriado a cada clique)
        abrir_contas1 = wait.until(EC.presence_of_element_located((By.ID, "selectConta")))

        # Simula o movimento e o clique humano
        time.sleep(1)
        ActionChains(driver).move_to_element(abrir_contas1).pause(0.3).click().perform()

        print(f"\n🔁 Reabrindo lista de contas (iteração {i+1})")

        try:
            # Verifica se aparece o modal do Pix Automático
            modal_pix = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "mat-dialog-container"))
            )
            print("⚠️ Modal Pix Automático detectado. Tentando fechar...")

            # Tenta clicar no botão Fechar dentro do modal
            botao_fechar = modal_pix.find_element(By.XPATH, ".//button//span[contains(text(), 'Fechar')]")
            driver.execute_script("arguments[0].click();", botao_fechar)
            print("✅ Modal Pix Automático fechado com sucesso.")

            # Espera o modal sumir
            WebDriverWait(driver, 5).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, "mat-dialog-container"))
            )

        except TimeoutException:
            # Se o modal não apareceu, segue normalmente
            pass
        except Exception as e:
            print(f"⚠️ Erro ao tentar fechar o modal Pix: {e}")

        lista = wait.until(EC.presence_of_element_located((By.ID, "listaContas")))
        contas = lista.find_elements(By.TAG_NAME, "li")
        total_contas = len(contas)  # Atualiza caso a lista mude

        # Garante que o índice ainda existe
        if i >= total_contas:
            print("⚠️ Lista de contas mudou, encerrando loop.")
            break

        conta = contas[i]
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", conta)
        nome_conta = conta.text.strip()
        
        if not nome_conta:
            continue

        print(f"\n=========== CONTA {i+1} de {total_contas} ===========")
        print(f"➡️ Acessando conta: {nome_conta}")

        # Clica na conta
        ActionChains(driver).move_to_element(conta).click().perform()
        time.sleep(3)

        ActionChains(driver).move_by_offset(2, 0).click().perform()

        estado = esperar_extrato_ou_bloqueio(driver, timeout=3)
        if estado == "bloqueio":
            print("⛔ Conta sem permissão de acesso ao saldo — pulando para a próxima.")
            continue

        # Espera o botão "Saldo e Extratos"
        saldo_extratos = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[.//span[contains(normalize-space(.), 'Saldo e Extratos')]]")
        ))
        saldo_extratos.click()



        # Espera o link "Extrato CAIXA" e clica
        extrato_caixa = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(normalize-space(.), 'Extrato CAIXA')]")
        ))
        extrato_caixa.click()

        # Espera o botão XLS e clica
        botao_xls = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//img[contains(@src, 'xls.svg')]/ancestor::button")
        ))
        botao_xls.click()

        print(f"💾 Clique no botão XLS realizado para {nome_conta}")

        # Pausa breve antes de reabrir a lista
        time.sleep(2)

        # Busca o arquivo XLS mais recente baixado
        arquivos = [
            os.path.join(diretorio_file_base, f)
            for f in os.listdir(diretorio_file_base)
            if f.lower().endswith(".xls")
        ]

        if arquivos:
            arquivo_recente = max(arquivos, key=os.path.getctime)

            # Extrai apenas o número da conta após "CC::"
            conta_numero = None
            agencia = None
            if "CC::" in nome_conta:
                partes = nome_conta.split("CC::")
                agencia = partes[0].split("AG:")[-1].strip()       # Pega o que vem depois de AG:
                conta_numero = partes[-1].strip()   
            else:
                conta_numero = nome_conta.strip()

            # Monta o nome novo do arquivo
            data_agora = time.strftime("%Y-%m-%d")
            novo_nome = f"Extrato_{data_agora}_AG{agencia}_CC{conta_numero}.xls"

            # Caminho final
            novo_caminho = os.path.join(diretorio_file_base, novo_nome)

            # Renomeia o arquivo
            os.rename(arquivo_recente, novo_caminho)
            print(f"✅ Arquivo renomeado para: {novo_nome}")
        else:
            print("⚠️ Nenhum arquivo XLS encontrado para renomear.")

        ActionChains(driver).move_by_offset(2, 0).click().perform()


    print("\n✅ Processo Finalizado!")
    driver.quit()
    print("Processo de Busca Finalizado!")
    break
