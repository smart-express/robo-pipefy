import pandas as pd
from playwright.sync_api import sync_playwright
from datetime import datetime
import re
import time

nomeDoArquivo = "NOME.xlsx"
df = pd.read_excel(nomeDoArquivo, dtype=str)


def novoDado(coluna, valor, container, pagina):
     botao_Perfil = container.locator("div").filter(has_text=coluna).locator("text=Clique aqui para adicionar")
     botao_Perfil.click()
     pagina.get_by_test_id("star-form-connection-button").wait_for()
     pagina.get_by_test_id("star-form-connection-button").click()
     pagina.get_by_role("combobox", name="Pesquisar").wait_for()
     pagina.get_by_role("combobox", name="Pesquisar").fill(valor)
     time.sleep(1)
     
     if pagina.get_by_text("0 resultados").is_visible():
        pagina.get_by_test_id("show-done-cards-checkbox").click()
        pagina.get_by_role("button", name=valor).nth(1).wait_for()
        pagina.get_by_role("button", name=valor).nth(1).click()
        pagina.get_by_test_id("save-button").wait_for()
        time.sleep(1)
        pagina.get_by_test_id("save-button").click()
        
     else:
        pagina.get_by_role("button", name=valor).nth(1).wait_for()
        pagina.get_by_role("button", name=valor).nth(1).click()
        pagina.get_by_test_id("save-button").wait_for()
        pagina.get_by_test_id("save-button").click()
    
with sync_playwright() as p:
    navegador = p.chromium.launch(headless=False)
    context = navegador.new_context()
    pagina = context.new_page()
    pagina.goto("https://www.pipefy.com/")

    pagina.get_by_role("link", name="Entrar").click()
    pagina.get_by_role("textbox", name="E-mail ou nome de usuário").fill("wesley.santos@superaholdings.com.br")
    pagina.get_by_role("button", name="Continuar").click()
    pagina.get_by_role("textbox", name="Senha").fill("MotoSerra@2")
    pagina.get_by_role("button", name="Acessar o Pipefy").click()
   
    pagina.get_by_test_id("databases-tab-button").wait_for()
    pagina.get_by_test_id("databases-tab-button").click()
    pagina.get_by_role("link", name="BLACKLIST (copy 1)").wait_for()
    pagina.get_by_role("link", name="BLACKLIST (copy 1)").click()
    
    for index, row in df.iterrows():
        container = pagina.get_by_test_id("open-record-infos-container")
        pagina.get_by_role("textbox", name="Buscar registros").wait_for()
        pagina.get_by_role("textbox", name="Buscar registros").fill(row["Nome"])
        pagina.get_by_role("cell", name=row["Nome"], exact=True).first.click()
        
        print(row["Nome"])
        print(index + 1, len(df))
        
        novoDado("Nome Promotor", str(row["Nome"]), container, pagina)
        pagina.get_by_test_id("open-record-btn-close").click()
        df.drop(index=index, inplace=True)
        df.to_excel(nomeDoArquivo, index=False)