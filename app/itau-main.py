from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains

#SISTEMAS
import re
import os, time, shutil
from datetime import datetime, date
import creds

data_hoje = date.today()

#APAGAR ARQUIVOS DO DIRETORIO
diretorio_base = creds.diretorio_base_itau
diretorio_file_base = creds.diretorio_file_base_itau
diretorio_consolidado = creds.diretorio_consolidado_itau

# 🧹 Apaga arquivos criados hoje no diretório de arquivos base
for arquivo in os.listdir(diretorio_file_base):
    caminho_arquivo = os.path.join(diretorio_file_base, arquivo)
    if os.path.isfile(caminho_arquivo):
        data_criacao = date.fromtimestamp(os.path.getctime(caminho_arquivo))
        if data_criacao == data_hoje:
            os.remove(caminho_arquivo)
            print(f"🗑️ Arquivo removido: {arquivo}")

# 🧹 Apaga apenas arquivos .xlsx criados hoje no diretório base
for arquivo in os.listdir(diretorio_base):
    caminho_arquivo = os.path.join(diretorio_base, arquivo)
    if os.path.isfile(caminho_arquivo) and arquivo.lower().endswith(".xlsx"):
        data_criacao = date.fromtimestamp(os.path.getctime(caminho_arquivo))
        if data_criacao == data_hoje:
            os.remove(caminho_arquivo)
            print(f"🗑️ Arquivo .xlsx removido: {arquivo}")

#DRIVER E OPTIONS
options = webdriver.FirefoxOptions()

# 2 = usar diretório customizado
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.dir", diretorio_file_base)

# mime types que devem ser baixados automaticamente
options.set_preference(
    "browser.helperApps.neverAsk.saveToDisk",
    "application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream"
)

# não abrir janela de download
options.set_preference("browser.download.manager.showWhenStarting", False)

# não perguntar onde salvar
options.set_preference("browser.download.useDownloadDir", True)

# desativar visualização interna de PDFs (não interfere, mas evita abrir no browser)
options.set_preference("pdfjs.disabled", True)

# evitar popup de aviso
options.set_preference("browser.download.alwaysOpenPanel", False)


service = Service()

driver = webdriver.Firefox(service=service, options=options)
driver.get("https://www.itau.com.br/")


# --COOKIES
try:
    banners = driver.find_elements(By.TAG_NAME, "itau-cookie-consent-banner")
    if banners:
        banner = banners[0]
        shadow_root = driver.execute_script("return arguments[0].shadowRoot", banner)

        try:
            cookie_button = shadow_root.find_element(By.ID, "itau-cookie-consent-banner-accept-cookies-btn")
            cookie_button.click()
            print("Botão de cookies clicado!")
        except NoSuchElementException:
            print("Banner encontrado, mas botão não disponível")

    else:
        print("Nenhum banner de cookies encontrado")

except Exception as e:
    print(f"Erro inesperado ao tentar lidar com cookies: {e}")



# -- ACESSO
accont = creds.accont_itau
psswd = creds.ppswd_itau

driver.find_element(By.ID, "open_modal_more_access").click()
# print("Botão de Acesso Clicado")
acesse_select = Select(driver.find_element(By.ID, "idl-more-access-select-login")).select_by_value("operator")
# print("Botão Selecionado Operador")
driver.find_element(By.ID, "idl-more-access-input-operator").send_keys(accont)
# print("Senha Inserida Operador")
driver.find_element(By.ID, "idl-more-access-submit-button").click()


#SENHA VIRTUAL

WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "campoTeclado")))
time.sleep(1)
botoes = driver.find_elements(By.ID, "campoTeclado")

time.sleep(0.6)
for _ in psswd:
    for i in botoes:
        botao = str(i.text)
        if re.search(_, botao):
            driver.execute_script("arguments[0].click();", i)
            # print(f"Botão {_} Clicado!")
            
Acess = driver.find_element(By.ID, "acessar")
driver.execute_script("arguments[0].click();", Acess)

#BOTÃO DE ACESSO SEM TOKEN
time.sleep(5.8)
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "rdBasico"))).click()
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "btn-continuar"))).click()


#ITERAÇÃO

def fechar_popup(driver, by, seletor, nome="popup"):
    try:
        elem = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((by, seletor))
        )
        elem.click()
        print(f"{nome} fechado")
        return True
    except TimeoutException:
        print(f"Nenhum {nome} aberto")
        return False
    
#Tentativa novamente do Extrato e PDF
def clicar_com_retry(driver, by, locator, descricao, timeout_total=220):
    """
    Espera o elemento aparecer e tenta clicar nele até que seja bem-sucedido
    ou até o timeout total (em segundos). 
    Caso o alerta de erro apareça, tenta clicar no link de reenvio (#lfsupo-link-extrato).
    """
    inicio = time.time()
    ultimo_erro = None

    while True:
        tempo_decorrido = time.time() - inicio
        if tempo_decorrido > timeout_total:
            print(f"❌ {descricao} - Tempo máximo de {timeout_total}s esgotado sem conseguir clicar.")
            
            # 🔄 Verifica se o alerta de erro está visível e tenta clicar no link
            try:
                alerta = driver.find_element(By.ID, "lfsupo-link-extrato")
                if alerta.is_displayed():
                    print("⚠️ Alerta detectado! Tentando clicar em 'acesse o extrato' e repetir a tentativa...")
                    alerta.click()
                    time.sleep(2)  # dá tempo para a página recarregar
                    # Reinicia o tempo e tenta novamente
                    inicio = time.time()
                    continue
            except NoSuchElementException:
                pass

            raise TimeoutException(f"Elemento '{descricao}' não apareceu a tempo.")

        try:
            elem = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((by, locator))
            )
            elem.click()
            print(f"✅ {descricao} - Clique realizado com sucesso após {int(tempo_decorrido)}s.")
            return True

        except (TimeoutException, ElementClickInterceptedException) as e:
            print(f"⚠️ {descricao} - Aguardando elemento aparecer ({int(tempo_decorrido)}s passados)...")
            ActionChains(driver).move_by_offset(2, 0).click().perform()
            ultimo_erro = e
            time.sleep(2)

        except Exception as e:
            print(f"⚠️ Erro inesperado ao tentar clicar em {descricao}: {e}")
            time.sleep(3)


def clicar_pdf_com_retorno(driver, max_tentativas_pdf=20):
    """
    Tenta clicar no botão 'Salvar em Excel' até max_tentativas_pdf vezes.
    Após clicar, espera o modal aparecer e clica no botão 'Salvar'.
    Depois aguarda a mensagem 'Arquivo salvo com sucesso'.
    Se falhar todas, volta no botão 'Extrato' e tenta novamente.
    """

    ultimo_erro = None
    xpath_excel = "//span[normalize-space()='Salvar em Excel']/ancestor::button"
    xpath_extrato = "//h2[contains(.,'Salvar extrato')]/ancestor::legend//button"
    xpath_salvar_modal = "//span[normalize-space()='Salvar']/ancestor::button"
    xpath_mensagem_sucesso = "//div[contains(@class,'voxel-alert__message') and contains(.,'Arquivo salvo com sucesso')]"

    # ==========================
    # 1️⃣ Tentativas no botão Excel
    # ==========================
    for tentativa in range(max_tentativas_pdf):
        try:
            ActionChains(driver).move_by_offset(2, 0).click().perform()
            botao_excel = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, xpath_excel))
            )
            botao_excel.click()
            print(f"✅ Botão 'Salvar em Excel' clicado na tentativa {tentativa+1}")

            # Espera o modal abrir e clica no botão "Salvar"
            botao_salvar = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, xpath_salvar_modal))
            )
            botao_salvar.click()
            print("✅ Botão 'Salvar' do modal clicado com sucesso")

            # Espera a mensagem de sucesso aparecer
            WebDriverWait(driver, 30).until(
                EC.visibility_of_element_located((By.XPATH, xpath_mensagem_sucesso))
            )
            print("🎉 Mensagem 'Arquivo salvo com sucesso' detectada!")
            return True

        except Exception as e:
            time.sleep(2)
            ultimo_erro = e
            print(f"⏳ Tentativa {tentativa+1} falhou para botão Excel.")
            time.sleep(1)

    # ==========================
    # 2️⃣ Se todas falharem → tenta o fallback
    # ==========================
    try:
        print("🔄 Reiniciando fluxo: clicando novamente no botão 'Extrato' antes do Excel...")
        botao_extrato = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, xpath_extrato))
        )
        botao_extrato.click()
        time.sleep(5)

        botao_excel = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, xpath_excel))
        )
        botao_excel.click()
        print("✅ Fluxo Extrato + Excel clicado, aguardando modal...")

        # Espera o modal e clica em Salvar
        botao_salvar = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, xpath_salvar_modal))
        )
        botao_salvar.click()
        print("✅ Botão 'Salvar' do modal clicado no fallback")

        # Espera a mensagem de sucesso
        WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.XPATH, xpath_mensagem_sucesso))
        )
        print("🎉 Mensagem 'Arquivo salvo com sucesso' detectada após fallback!")
        return True

    except Exception as e:
        print("❌ Fluxo Extrato + Excel falhou mesmo após fallback")
        print(f"Erro final: {e}")
        raise e
    

def clicar_perfil_usuario(driver, tentativas=10, tempo_espera=5):
    """
    Repete o clique até encontrar e clicar no elemento com a classe 'perfil-usuario-ni'.
    
    Args:
        driver: Instância do WebDriver.
        tentativas: Quantas vezes tentar clicar antes de gerar exceção.
        tempo_espera: Tempo (em segundos) para esperar entre as tentativas.
    """
    for i in range(1, tentativas + 1):
        try:
            print(f"🔁 Tentativa {i}/{tentativas} — clicando via ActionChains...")
            ActionChains(driver).move_by_offset(2, 0).click().perform()

            # Aguarda até que o elemento apareça
            botao = WebDriverWait(driver, tempo_espera).until(
                EC.presence_of_element_located((By.CLASS_NAME, "perfil-usuario-ni"))
            )

            botao.click()
            print("✅ Elemento 'perfil-usuario-ni' encontrado e clicado com sucesso!")
            return True

        except Exception as e:
            print(f"⚠️ Tentativa {i} falhou: {e}")
            time.sleep(tempo_espera)
            continue

    # Se chegar aqui, todas as tentativas falharam
    raise Exception("❌ Não foi possível clicar no elemento 'perfil-usuario-ni' após várias tentativas.")

def sanitize_filename(nome: str) -> str:
    # remove ou troca caracteres inválidos para Windows
    return re.sub(r'[\\/*?:"<>|]', "-", nome)

array_accont = []
exists_counts = True
while exists_counts:
    
    time.sleep(4)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "seta-direita-ni"))).click()
    general_rows = WebDriverWait(driver, 20).until(
        EC.presence_of_all_elements_located((By.CLASS_NAME, "linha-contas-disponiveis"))
    )
    account = driver.find_elements(By.CLASS_NAME, "linha-contas-disponiveis")
    for _ in account:
        texto = str(_.text)
        texto_tratado = texto.replace(" ", "-")
        array_accont.append(texto_tratado)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "fechar-contas"))).click()
    last_count = len(general_rows) -1
    time.sleep(1.5)
    for i in range(len(general_rows)):
        if i == 0:
            fechar_popup(driver, By.CSS_SELECTOR, "div.icon-content[role='button'][aria-label*='fechar']", "Tutorial")
            print("Entrou condição index 0")
            time.sleep(1)
        
        clicar_perfil_usuario(driver, tentativas=15, tempo_espera=3)

        # time.sleep(1)
        general_rows = WebDriverWait(driver, 20).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "linha-contas-disponiveis"))
        )
        conta = general_rows[i]
        print(f"Processando conta {i}")
        driver.execute_script("arguments[0].click();", conta)
        
        #BOTÃO DE EXTRATO E PDF
        clicar_com_retry(driver, By.ID, "btnExtrato", "Botão Extrato")
        time.sleep(1.5)
        clicar_pdf_com_retorno(driver)

        #RENOMEAR O ARQUIVO
        time.sleep(1.5)

        arquivos = os.listdir(diretorio_file_base)
        arquivos = [os.path.join(diretorio_file_base, f) for f in arquivos]
        ultimo = max(arquivos, key=os.path.getctime)
        nome_limpo = sanitize_filename(array_accont[i])
        n_nome = f"{nome_limpo}({datetime.now().strftime('%d-%m-%Y')}).xlsx"
        novo_nome = os.path.join(diretorio_file_base, n_nome)
        shutil.move(ultimo, novo_nome)
        print(f"Arquivo foi  criado {n_nome}")

        if i == last_count:
            exists_counts = False
            driver.quit()
            print("Processo de Busca Finalizado!")

