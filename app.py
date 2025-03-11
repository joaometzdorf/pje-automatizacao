from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.support.select import Select
import openpyxl

numero_oab = 259155
planilha_dados_processos = openpyxl.load_workbook("dados_de_processos.xlsx")
planilha_processos = planilha_dados_processos["processos"]

# 1 - Entrar no site https://pje-consulta-publica.tjmg.jus.br/
driver = webdriver.Chrome()
driver.get("https://pje-consulta-publica.tjmg.jus.br/")
sleep(2)

# 2 - Clicar no campo de oab e digitar o número do advogado
campo_numero_oab = driver.find_element(
    By.XPATH, "//input[@id='fPP:Decoration:numeroOAB']"
)
sleep(2)
campo_numero_oab.click()
sleep(1)
campo_numero_oab.send_keys(numero_oab)
sleep(1)

# 3 - Selecionar o estado do advogado
selecao_uf = driver.find_element(
    By.XPATH, "//select[@id='fPP:Decoration:estadoComboOAB']"
)
sleep(1)
opcoes_uf = Select(selecao_uf)
sleep(1)
opcoes_uf.select_by_visible_text("SP")
sleep(1)

# 4 - Clicar em pesquisar
botao_pesquisar = driver.find_element(By.XPATH, "//input[@id='fPP:searchProcessos']")
sleep(1)
botao_pesquisar.click()
sleep(3)

# 5 - Entrar em cada um dos processos e extrair número do processo, número do advogado e nome dos participantes
links_abrir_processo = driver.find_elements(By.XPATH, "//a[@title='Ver Detalhes']")
for link in links_abrir_processo:
    janela_principal = driver.current_window_handle
    link.click()
    sleep(3)
    janelas_abertas = driver.window_handles
    for janela in janelas_abertas:
        if janela not in janela_principal:
            driver.switch_to.window(janela)
            sleep(3)
            numero_processo = driver.find_elements(
                By.XPATH, "//div[@class='propertyView ']//div[@class='col-sm-12 ']"
            )[0]
            participantes = driver.find_elements(
                By.XPATH,
                "//tbody[contains(@id,'processoPartesPoloAtivoResumidoList:tb')]//span[@class='text-bold']",
            )

            lista_participantes = []
            for participante in participantes:
                lista_participantes.append(participante.text)

            # Guardar um participante se houver apenas um
            if len(lista_participantes) == 1:
                planilha_processos.append(
                    [numero_oab, numero_processo.text, lista_participantes[0]]
                )
            else:
                planilha_processos.append(
                    [numero_oab, numero_processo.text, ",".join(lista_participantes)]
                )
            # 6 - Salvar os dados em uma planilha
            planilha_dados_processos.save("dados_de_processos.xlsx")
            driver.close()
    # 7 - Repetir até finalizar todos os dados daquele advogado
    driver.switch_to.window(janela_principal)
