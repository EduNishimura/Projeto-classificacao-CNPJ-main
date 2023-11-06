from time import sleep
from Tratamentos import *
import pandas as pd
from playwright.sync_api import sync_playwright
import openpyxl

# definicao de inputs
# nome do arquivo xlsx
arquivo = openpyxl.load_workbook("cnpjs.xlsx")
# nome da planilha
planilha_principal = arquivo["Planilha1"]
CNAE = arquivo['CNAE']
# linha para iniciar o programa
celula_inicio = 2

# tratando o arquivo
tirando_separadores_e_none(arquivo, planilha_principal)
tirando_duplicatas_e_espacos(arquivo, planilha_principal)
tirando_cei(arquivo, planilha_principal)
adicionando_cnae(arquivo)


# Abrindo o Navegador
with sync_playwright() as p:
     # abrindo o navegador
     navegador = p.chromium.launch(headless=False)

# coletando os dados e passando para a planilha
for c in range(celula_inicio, len(planilha_principal['A'])):
    try:

        cedula_atual = planilha_principal[f'A{c}'].value

        # Entrando no site
        pagina = navegador.new_page()
        pagina.goto("https://www.informecadastral.com.br/")
        sleep(2)
        pagina.fill('xpath=//*[@id="cnpj"]', f'{cedula_atual}')
        pagina.locator('xpath=//*[@id="formSearch"]/div/button').click()
        sleep(3)

        # Pegando informações
        razao_social = pagina.locator('xpath=/html/body/div/div[2]/div/div/div[2]/div[1]/div/h1').text_content()
        planilha_principal[f'B{c}'].value = razao_social

        estado_municipio = pagina.locator('xpath=/html/body/div/div[2]/div/div/div[3]/div[2]/div[2]/div[2]/div[1]/p').text_content()
        partes_estado_mun = estado_municipio.split("|")
        estado = partes_estado_mun[0]
        planilha_principal[f'C{c}'].value = estado
        municipio = partes_estado_mun[1]
        planilha_principal[f'D{c}'].value = municipio

        nome_especifico = pagina.locator("xpath=//div[contains(@class, 'col-md-10')]/p").text_content()
        planilha_principal[f'J{c}'].value = nome_especifico

        # Código CNAE
        codigo_cnae = pagina.locator("xpath=//div[contains(@class, 'col-md-2')]/p").text_content()
        planilha_principal[f'E{c}'].value = codigo_cnae

        # CNAE SEÇÃO E DIVISÃO
        planilha_principal[f'F{c}'].value = int(codigo_cnae[:2])
        for cell in range(1, len(CNAE["C"])):
            if CNAE[f"C{cell}"].value == planilha_principal[f'F{c}'].value:
                planilha_principal[f'F{c}'].value = CNAE[f'A{cell}'].value
                planilha_principal[f'G{c}'].value = CNAE[f'B{cell}'].value
                planilha_principal[f'H{c}'].value = CNAE[f'C{cell}'].value
                planilha_principal[f'I{c}'].value = CNAE[f'D{cell}'].value
        print(f'{c}º')
    except Exception as e:
        print(f'{Exception}, {e}')
        # Salvando
    arquivo.save(arquivo)
# tabela = pd.read_excel(arquivo)
# print(tabela)
