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

import pandas as pd

data_hoje = date.today()

#APAGAR ARQUIVOS DO DIRETORIO
diretorio_base = creds.diretorio_base_itau
diretorio_file_base = creds.diretorio_file_base_itau
diretorio_consolidado = creds.diretorio_consolidado_itau

# 🧹 Apaga arquivos criados hoje no diretório de arquivos base
for arquivo in os.listdir(diretorio_file_base):
    caminho_arquivo = os.path.join(diretorio_file_base, arquivo)
    nome_upper = arquivo.upper()

    # IGNORAR arquivos da Caixa, Histórico e Consolidado
    if (
        "CAIXA" in nome_upper or
        "HISTORICO" in nome_upper or
        "CONSOLIDADO" in nome_upper
    ):
        continue

    if os.path.isfile(caminho_arquivo):
        data_criacao = date.fromtimestamp(os.path.getctime(caminho_arquivo))

        # excluir somente se criado HOJE
        if data_criacao == data_hoje:
            os.remove(caminho_arquivo)
            print(f"🗑️ Arquivo removido: {arquivo}")


# 🧹 Apaga apenas arquivos .xlsx criados hoje no diretório base
for arquivo in os.listdir(diretorio_base):
    caminho_arquivo = os.path.join(diretorio_base, arquivo)
    nome_upper = arquivo.upper()

    # IGNORAR arquivos da Caixa, Histórico e Consolidado
    if (
        "CAIXA" in nome_upper or
        "HISTORICO" in nome_upper or
        "CONSOLIDADO" in nome_upper
    ):
        continue

    # só arquivos .xlsx
    if os.path.isfile(caminho_arquivo) and arquivo.lower().endswith(".xlsx"):
        data_criacao = date.fromtimestamp(os.path.getctime(caminho_arquivo))

        # excluir somente se criado HOJE
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
    # Espera até que o banner de cookies esteja presente
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "itau-cookie-consent-banner")))

    # Acessa o banner de cookies
    banners = driver.find_elements(By.TAG_NAME, "itau-cookie-consent-banner")
    if banners:
        banner = banners[0]
        shadow_root = driver.execute_script("return arguments[0].shadowRoot", banner)

        try:
            # Espera até que o botão de aceitação de cookies esteja presente e clicável
            cookie_button = WebDriverWait(shadow_root, 20).until(
                EC.element_to_be_clickable((By.ID, "itau-cookie-consent-banner-accept-cookies-btn"))
            )
            cookie_button.click()
            print("Botão de cookies clicado!")
        except TimeoutException:
            print("O botão de cookies não estava disponível a tempo.")
        except NoSuchElementException:
            print("Banner encontrado, mas botão não disponível.")
    else:
        print("Nenhum banner de cookies encontrado.")

except TimeoutException:
    print("O banner de cookies não foi encontrado a tempo.")
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


def esperar_download_concluir(diretorio, timeout=40):
    """
    Aguarda um novo arquivo aparecer no diretório e
    garante que não é .crdownload / .tmp
    """
    inicio = time.time()
    arquivos_iniciais = set(os.listdir(diretorio))

    while time.time() - inicio < timeout:
        arquivos_atuais = set(os.listdir(diretorio))
        novos = arquivos_atuais - arquivos_iniciais

        arquivos_validos = [
            f for f in novos
            if not f.endswith(".crdownload") and not f.endswith(".tmp")
        ]

        if arquivos_validos:
            return True

        time.sleep(1)

    return False



def clicar_excel_com_retorno(
    driver,
    diretorio_download,
    max_tentativas=10
):
    """
    Clica em 'Salvar em Excel', aguarda o download concluir
    e retorna True para seguir o fluxo.
    """

    xpath_li_excel = "//li[@id='salvarXls']"
    xpath_excel = "//li[@id='salvarXls']//a[contains(@href,'xls')]"

    ultimo_erro = None

    for tentativa in range(max_tentativas):
        try:
            print(f"🔁 Tentativa {tentativa+1}/{max_tentativas} — Excel")

            # 1️⃣ Garante que o menu de exportação existe
            WebDriverWait(driver, 25).until(
                EC.presence_of_element_located((By.XPATH, xpath_li_excel))
            )

            # 2️⃣ Localiza o botão Excel
            botao_excel = WebDriverWait(driver, 25).until(
                EC.presence_of_element_located((By.XPATH, xpath_excel))
            )

            # 3️⃣ Scroll + clique via JS (obrigatório no Itaú)
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", botao_excel
            )
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", botao_excel)

            print("✅ Clique em 'Salvar em Excel' executado")

            # 4️⃣ Aguarda o download real
            if esperar_download_concluir(diretorio_download, timeout=40):
                print("📥 Download concluído com sucesso")
                return True

            raise Exception("Download não detectado")

        except Exception as e:
            ultimo_erro = e
            print(f"⏳ Falha na tentativa {tentativa+1}: {e}")
            time.sleep(2)

    print("❌ Todas as tentativas de download falharam")
    raise ultimo_erro


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

def converter_xls_para_xlsx(caminho_xls):
    caminho_xlsx = caminho_xls.replace(".xls", ".xlsx")

    df = pd.read_excel(caminho_xls, engine="xlrd")
    df.to_excel(caminho_xlsx, index=False, engine="openpyxl")

    os.remove(caminho_xls)
    return caminho_xlsx

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
        clicar_excel_com_retorno(
            driver,
            diretorio_file_base
        )

        #RENOMEAR O ARQUIVO
        time.sleep(2)
        # Arquivo baixado
        arquivos = os.listdir(diretorio_file_base)
        arquivos = [os.path.join(diretorio_file_base, f) for f in arquivos]
        ultimo = max(arquivos, key=os.path.getctime)

        # Renomeia mantendo XLS
        nome_limpo = sanitize_filename(array_accont[i])
        nome_xls = f"{nome_limpo}({datetime.now().strftime('%d-%m-%Y')}).xls"
        novo_xls = os.path.join(diretorio_file_base, nome_xls)
        shutil.move(ultimo, novo_xls)

        # Converte para XLSX real
        novo_xlsx = converter_xls_para_xlsx(novo_xls)

        print(f"📄 Arquivo convertido com sucesso: {os.path.basename(novo_xlsx)}")

        if i == last_count:
            exists_counts = False
            driver.quit()
            print("Processo de Busca Finalizado!")

