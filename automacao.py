from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
from datetime import datetime
import os
import shutil
from selenium.webdriver.support.ui import Select

# ================= Configurações =================

# Caminho da pasta onde os PDFs serão salvos (mudar caminha unico)
PASTA_DESTINO = r"path\pasta_nfe_barueri\NOTAS"

# Configurações do Chrome (download automático)
opcoes = webdriver.ChromeOptions()
opcoes.add_experimental_option(
    "prefs",
    {
        "download.default_directory": os.path.abspath(PASTA_DESTINO),
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
    },
)
opcoes.add_argument("--start-maximized")

# Inicia o navegador
driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()), options=opcoes
)

# ================= Acessa o site =================
driver.get("https://www.barueri.sp.gov.br/nfe/")

# ================= Clica em "Acesso ao Sistema" =================
try:
    botao_acesso = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "link_url_login"))
    )
    botao_acesso.click()
    print("✅ Clicou no link 'Acesso ao Sistema' com sucesso!")
except Exception as e:
    print("❌ Erro ao tentar clicar no link 'Acesso ao Sistema':", e)

time.sleep(1)

# ================= Troca o foco para o iframe =================
try:
    WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrmPrincipal"))
    )
    print("🪟 Foco alterado para o iframe 'ifrmPrincipal'")
except Exception as e:
    print("⚠️ Nenhum iframe encontrado ou não necessário:", e)

# ================= Digita usuário =================
try:
    campo_usuario = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "txtLoginUsuario"))
    )
    campo_usuario.click()
    campo_usuario.clear()
    x = input("Digite o login : ")
    campo_usuario.send_keys(x)  # substitua pelo seu usuário
    print(f"✅ Valor '{x}' digitado com sucesso no campo de login")
except Exception as e:
    print("❌ Erro ao interagir com o campo de login:", e)

# ================= Digita senha =================
try:
    campo_senha = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "txtSenhaUsuario"))
    )
    campo_senha.click()
    y = input("Digite o login : ")
    campo_senha.send_keys(y)
    print("✅ Valor senha digitada com sucesso")
except Exception as e:
    print("❌ Erro ao interagir com o campo de senha:", e)

time.sleep(1)

# ================= Clica em "Acessar" =================
try:
    btn_acessar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "btnConfirmar"))
    )
    btn_acessar.click()
    print("✅ Clicou no botão 'Acessar o Sistema'")
except Exception as e:
    print("❌ Erro ao clicar no botão 'Acessar o Sistema':", e)

# ================= Espera carregar a tela após login =================
time.sleep(2)


# Volta para o contexto principal
driver.execute_script(
    "window.open('FiltroConsultaNF.aspx?TipoOp=Consultar','ifrmPrincipal');"
)
print("✅ Ação 'Consultar NF-e' executada manualmente via JavaScript")


# ================= Aguardar visualização antes de fechar =================
time.sleep(2)

try:
    # Garante que está dentro do iframe certo
    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrmPrincipal"))
    )

    # Espera o checkbox aparecer e ficar clicável
    chk_periodo = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.ID, "chkSelPeriodo"))
    )
    # ================= escreve data inicio =================

    # Clica no checkbox
    chk_periodo.click()
    print("✅ Checkbox 'Selecionar Período' marcado com sucesso!")

except Exception as e:
    print("❌ Erro ao marcar o checkbox 'Selecionar Período':", e)

    # escreve data inicio =================

time.sleep(2)
i = 0
while i == 0:
    data_input = input("Digite a data de inicio no formato dd/mm/aaaa: ")
    # data_input = "01/10/2025"
    try:
        # Tenta converter a string em data
        data_valida = datetime.strptime(data_input, "%d/%m/%Y")
        data_formatada = data_valida.strftime("%d/%m/%Y")

        # Se não der erro, a data está correta
        print(f"✅ Data válida: {data_valida.strftime('%d/%m/%Y')}")
        i = 1
        # ======== Cola o valor no campo da página ========
        try:
            # Garante que está no iframe correto antes de buscar o campo
            driver.switch_to.default_content()
            WebDriverWait(driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrmPrincipal"))
            )

            # Espera o campo de data ficar clicável
            campo_data_inicio = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.ID, "txtIniPeriodo"))
            )

            # Clica, limpa e escreve a data validada
            campo_data_inicio.click()
            campo_data_inicio.clear()
            campo_data_inicio.send_keys(data_formatada)

            print(f"✅ Campo 'Data Início' preenchido com: {data_formatada}")
            time.sleep(2)
        except Exception as e:
            print("❌ Erro ao preencher o campo 'Data Início':", e)

    except ValueError:
        # Se der erro na conversão, o formato está errado
        print("❌ Formato inválido! Use o formato dd/mm/aaaa (ex: 01/12/2025).")

# ================= escreve data fim =================

time.sleep(2)
i = 0
while i == 0:
    data_input = input("Digite a data de fim no formato dd/mm/aaaa: ")
    # data_input = "20/10/2025"
    try:
        # Tenta converter a string em data
        data_valida_fim = datetime.strptime(data_input, "%d/%m/%Y")
        data_formatada_fim = data_valida_fim.strftime("%d/%m/%Y")

        # Se não der erro, a data está correta
        print(f"✅ Data válida: {data_valida_fim.strftime('%d/%m/%Y')}")
        i = 1
        # ======== Cola o valor no campo da página ========
        try:
            # Garante que está no iframe correto antes de buscar o campo
            driver.switch_to.default_content()
            WebDriverWait(driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrmPrincipal"))
            )

            # Espera o campo de data ficar clicável
            campo_data_inicio = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.ID, "txtFimPeriodo"))
            )

            # Clica, limpa e escreve a data validada
            campo_data_inicio.click()
            campo_data_inicio.clear()
            campo_data_inicio.send_keys(data_formatada_fim)

            print(f"✅ Campo 'Data fim' preenchido com: {data_formatada_fim}")
            time.sleep(2)

        except Exception as e:
            print("❌ Erro ao preencher o campo 'Data fim':", e)

    except ValueError:
        # Se der erro na conversão, o formato está errado
        print("❌ Formato inválido! Use o formato dd/mm/aaaa (ex: 01/12/2025).")

WebDriverWait(driver, 15).until(
    EC.element_to_be_clickable((By.ID, "btnConsultar"))
).click()

# # ================= Detecta todas as listas e baixa =================


# ================= Espera a tabela de notas carregar =================

WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "grdNotasFiscais"))
)

# ================= Localiza o select de paginação =================
select_paginas = Select(
    driver.find_element(
        By.XPATH,
        "/html/body/form/div[3]/div[2]/table/tbody/tr[9]/td/div/table/tbody/tr[2]/td/div/select",
    )
)
total_paginas = len(select_paginas.options)
print(f"✅ Total de páginas: {total_paginas}")

# ================= Loop por todas as páginas =================
try:
    # Espera a tabela inicial carregar
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "grdNotasFiscais"))
    )

    # Localiza o select de paginação
    select_paginas = Select(
        driver.find_element(
            By.XPATH,
            "/html/body/form/div[3]/div[2]/table/tbody/tr[9]/td/div/table/tbody/tr[2]/td/div/select",
        )
    )
    total_paginas = len(select_paginas.options)
    print(f"✅ Total de páginas: {total_paginas}")
    n_notas = 0

    for pagina_atual in range(total_paginas):

        # Seleciona a página atual
        select_paginas.select_by_index(pagina_atual)
        print(f"📄 Navegando para a página {pagina_atual + 1}")

        # Espera a tabela atualizar
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "grdNotasFiscais"))
        )
        time.sleep(2)  # Pequena pausa para garantir carregamento

        # Recupera os botões da página atual
        botoes_notas = driver.find_elements(
            By.XPATH,
            "//a[starts-with(@id,'grdNotasFiscais_ctl') and contains(@id,'_ibtnVisualizar')]",
        )
        print(f"✅ Total de notas nesta página: {len(botoes_notas)}")

        for i in range(len(botoes_notas)):

            try:
                # Sempre encontra o botão novamente pelo índice
                botoes_atualizados = driver.find_elements(
                    By.XPATH,
                    "//a[starts-with(@id,'grdNotasFiscais_ctl') and contains(@id,'_ibtnVisualizar')]",
                )
                botao = botoes_atualizados[i]

                # Espera o botão estar clicável
                botao_clicavel = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, botao.get_attribute("id")))
                )

                # Clica via JavaScript
                driver.execute_script("arguments[0].click();", botao_clicavel)
                print(f"📝 Baixando nota {i + 1} desta página...")
                time.sleep(2)  # Pausa para download/processamento

                x = i + 1
                y = pagina_atual + 1
                arquivo_atual = rf"path\wfImagemNota.aspx"
                novo_nome = rf"path\ NFE_linha{x}_pagina{y}.aspx"
                os.rename(arquivo_atual, novo_nome)

            except Exception as e_inner:
                print(f"❌ Erro ao baixar nota {i + 1} desta página:", e_inner)

        # Reatualiza o select de páginas após cada página
        select_paginas = Select(
            driver.find_element(
                By.XPATH,
                "/html/body/form/div[3]/div[2]/table/tbody/tr[9]/td/div/table/tbody/tr[2]/td/div/select",
            )
        )

    print("✅ Todas as notas de todas as páginas foram baixadas!")

except Exception as e:
    print("❌ Erro ao processar as notas:", e)
# ================= Fecha o navegador =================

time.sleep(2)
# Pasta onde estão os .aspx
pasta_aspx = r"path\NOTAS"

# Pasta onde os PDFs serão salvos
pasta_pdf = r"path\Notas_pdf"

os.makedirs(pasta_pdf, exist_ok=True)

for arquivo in os.listdir(pasta_aspx):
    if arquivo.lower().endswith(".aspx"):
        caminho_origem = os.path.join(pasta_aspx, arquivo)
        nome_pdf = os.path.splitext(arquivo)[0] + ".pdf"
        caminho_destino = os.path.join(pasta_pdf, nome_pdf)

        shutil.copy2(caminho_origem, caminho_destino)  # copia e renomeia
        print(f"✅ {arquivo} → {nome_pdf}")

print("\n✅ Todos os arquivos foram renomeados para PDF!")

# Debug : Testar renomear por pagina  e converter em pdf
#
