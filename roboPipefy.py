import pandas as pd
from playwright.sync_api import sync_playwright
from datetime import datetime
import re
import time

nomeDoArquivo = "UFNOVO.xlsx"
df = pd.read_excel(nomeDoArquivo, dtype=str)

def alterarCampoPreenchido(coluna, valor, container, pagina):
    container.get_by_text(coluna, exact=True).locator("..").hover()
    botao_Editar = container.get_by_text(coluna, exact=True).locator("..").get_by_role("button", name="Editar")
    botao_Editar.wait_for()
    botao_Editar.click()
    
    if container.get_by_role("textbox").is_visible():
         pagina.get_by_test_id("textfield-input").fill("")
         pagina.get_by_test_id("textfield-input").fill(valor)
         
         pagina.get_by_test_id("save-button").click()
         time.sleep(1)
    else:
         pagina.get_by_test_id("select-field").wait_for()
         pagina.get_by_test_id("select-field").click()
         pagina.get_by_test_id("select-portal-content").wait_for()
         pagina.get_by_role("grid", name="grid").wait_for()
         if not pagina.get_by_role("grid", name="grid").get_by_role("option", name=valor, exact=True).is_visible():
            pagina.get_by_role("grid", name="grid").evaluate("el => el.scrollTo({ top: 250, behavior: 'smooth' })")
            pagina.get_by_role("grid", name="grid").get_by_role("option", name=valor).scroll_into_view_if_needed()
         pagina.get_by_role("option", name=re.compile(rf"^{valor}$"), exact=True).locator("div").first.wait_for()
         pagina.get_by_role("option", name=re.compile(rf"^{valor}$"), exact=True).locator("div").first.click()
         pagina.get_by_test_id("save-button").click()
         time.sleep(1)

def novoDado(coluna, valor, container, pagina):
     botao_Perfil = container.locator("div").filter(has_text=coluna).locator("text=Clique aqui para adicionar")
     botao_Perfil.click()
     botao_digite_aqui = pagina.get_by_placeholder("Digite aqui ...")
     botao_escolha =  pagina.get_by_placeholder("Escolha uma opção")

     if botao_digite_aqui.is_visible():
         pagina.get_by_test_id("textfield-input").fill(valor)
         pagina.get_by_test_id("save-button").click()
         time.sleep(1)
         
     elif botao_escolha.is_visible():
         pagina.get_by_test_id("select-field").wait_for()
         pagina.get_by_test_id("select-field").click()
         pagina.get_by_test_id("select-portal-content").wait_for()
         pagina.get_by_role("grid", name="grid").wait_for()
         if not pagina.get_by_role("grid", name="grid").get_by_role("option", name=valor, exact=True).is_visible():
            pagina.get_by_role("grid", name="grid").evaluate("el => el.scrollTo({ top: 250, behavior: 'smooth' })")
            pagina.get_by_role("grid", name="grid").get_by_role("option", name=valor).scroll_into_view_if_needed()
         pagina.get_by_role("option", name=re.compile(rf"^{valor}$"), exact=True).locator("div").first.wait_for()
         pagina.get_by_role("option", name=re.compile(rf"^{valor}$"), exact=True).locator("div").first.click()
         pagina.get_by_test_id("save-button").click()
         time.sleep(1)

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
    pagina.get_by_role("link", name="USUARIOS EXPRESS").wait_for()
    pagina.get_by_role("link", name="USUARIOS EXPRESS").click()
    
    for index, row in df.iterrows():
        container = pagina.get_by_test_id("open-record-infos-container")
        pagina.get_by_role("searchbox", name="Buscar registros").wait_for()
        pagina.get_by_role("searchbox", name="Buscar registros").fill(row["Nome"])
        pagina.get_by_role("cell", name=row["Nome"], exact=True).first.click()
        
        print(row["Nome"])
        print(index + 1, len(df))
   
        time.sleep(1)
        
        if "Codigo" in df.columns:
            if container.locator("div").filter(has_text="Codigo").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("Codigo", str(row["Codigo"]).zfill(6), container, pagina)
                
            else:
                alterarCampoPreenchido("Codigo", str(row["Codigo"]).zfill(6), container, pagina)
        
        if "Razão Social" in df.columns:
            if container.locator("div").filter(has_text="Razão Social").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("Razão Social", str(row["Razão Social"]), container, pagina)
                
            else:
                alterarCampoPreenchido("Razão Social", str(row["Razão Social"]), container, pagina)
        
        if "Perfil" in df.columns:
            if container.locator("div").filter(has_text="Perfil").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("Perfil", str(row["Perfil"]), container, pagina)
                
            else:
                alterarCampoPreenchido("Perfil", str(row["Perfil"]), container, pagina)
        
        if "Banco Novo" in df.columns:
            if container.locator("div").filter(has_text="* Banco Novo").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("* Banco Novo", str(row["Banco Novo"]), container, pagina)
                
            else:
                alterarCampoPreenchido("* Banco Novo", str(row["Banco Novo"]), container, pagina)
        
        if "CodigoBanco" in df.columns:
            if container.locator("div").filter(has_text="CodigoBanco").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("CodigoBanco", str(row["CodigoBanco"]), container, pagina)
                
            else:
                alterarCampoPreenchido("CodigoBanco", str(row["CodigoBanco"]), container, pagina)
        
        if "Empresa" in df.columns:
            if container.locator("div").filter(has_text="Empresa").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("Empresa", str(row["Empresa"]), container, pagina)
            
            else:
                alterarCampoPreenchido("Empresa", str(row["Empresa"]), container, pagina)
        
        if "Projeto" in df.columns:
            if container.locator("div").filter(has_text="Projeto").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("Projeto", str(row["Projeto"]), container, pagina)
            
            else:
                alterarCampoPreenchido("Projeto", str(row["Projeto"]), container, pagina)
        
        if "Login" in df.columns:
            if container.locator("div").filter(has_text="Login").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("Login", str(row["Login"]).zfill(11), container, pagina)
            
            else:
                alterarCampoPreenchido("Login", str(row["Login"]).zfill(11), container, pagina)
                
        if "Titulo de Eleitor" in df.columns:
            if container.locator("div").filter(has_text="Titulo de Eleitor").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("Titulo de Eleitor", str(row["Titulo de Eleitor"]), container, pagina)
            
            else:
                alterarCampoPreenchido("Titulo de Eleitor", str(row["Titulo de Eleitor"]), container, pagina)
        
        if "CadastroPessoa" in df.columns:
            if container.locator("div").filter(has_text="CadastroPessoa").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("CadastroPessoa", str(row["CadastroPessoa"]), container, pagina)
            
            else:
                alterarCampoPreenchido("CadastroPessoa", str(row["CadastroPessoa"]), container, pagina)
           
        if "IdPerfilAcessoSistema" in df.columns:
            if container.locator("div").filter(has_text="IdPerfilAcessoSistema").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("IdPerfilAcessoSistema", str(row["IdPerfilAcessoSistema"]), container, pagina)
            
            else:
                alterarCampoPreenchido("IdPerfilAcessoSistema", str(row["IdPerfilAcessoSistema"]), container, pagina)
         
        if "UF" in df.columns:
            if container.locator("div").filter(has_text="UF").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("UF", str(row["UF"]), container, pagina)
            
            else:
                alterarCampoPreenchido("UF", str(row["UF"]), container, pagina)
        
        if "UF Novo" in df.columns:
            if container.locator("div").filter(has_text="UF Novo").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("UF Novo", str(row["UF Novo"]), container, pagina)
            
            else:
                alterarCampoPreenchido("UF Novo", str(row["UF Novo"]), container, pagina)
        
        if "Telefone" in df.columns:
            if container.locator("div").filter(has_text="Telefone").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("Telefone", str(row["Telefone"]), container, pagina)
            
            else:
                alterarCampoPreenchido("Telefone", str(row["Telefone"]), container, pagina)
                 
        if "CPF" in df.columns:
            if container.locator("div").filter(has_text="CPF").locator("text=Clique aqui para adicionar").is_visible():
             
                 novoDado("CPF", str(row["CPF"]).zfill(11), container, pagina)
            
            else:
                alterarCampoPreenchido("CPF", str(row["CPF"]).zfill(11), container, pagina)
                
        if "CNPJ" in df.columns:
            if container.locator("div").filter(has_text="CNPJ").locator("text=Clique aqui para adicionar").is_visible():
             
                 novoDado("CNPJ", str(row["CNPJ"]).zfill(11), container, pagina)
            
            else:
                alterarCampoPreenchido("CNPJ", str(row["CNPJ"]).zfill(11), container, pagina)        
        
        if "PossuiContabilidade" in df.columns:
            if container.locator("div").filter(has_text="PossuiContabilidade").locator("text=Clique aqui para adicionar").is_visible():
             
                 novoDado("PossuiContabilidade", str(row["PossuiContabilidade"]).zfill(11), container, pagina)
            
            else:
                alterarCampoPreenchido("PossuiContabilidade", str(row["PossuiContabilidade"]).zfill(11), container, pagina)
              
        if "CustoOperacional" in df.columns:   
            if container.locator("div").filter(has_text="CustoOperacional").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("CustoOperacional", str(row["CustoOperacional"]), container, pagina)
            
            else:
                alterarCampoPreenchido("CustoOperacional", str(row["CustoOperacional"]), container, pagina)
                
        if "Lider" in df.columns:
            if container.locator("div").filter(has_text="Lider").locator("text=Clique aqui para adicionar").is_visible():
                
                novoDado("Lider", str(row["Lider"]), container, pagina)
                
            else:
                alterarCampoPreenchido("Lider", str(row["Lider"]), container, pagina)
                
        if "Superior" in df.columns:
            if container.locator("div").filter(has_text="Superior").locator("text=Clique aqui para adicionar").is_visible():
                
                novoDado("Superior", str(row["Superior"]), container, pagina)
                
            else:
                alterarCampoPreenchido("Superior", str(row["Superior"]), container, pagina)
                
        if "Coordenador" in df.columns:
            if container.locator("div").filter(has_text="Coordenador").locator("text=Clique aqui para adicionar").is_visible():
                
                novoDado("Coordenador", str(row["Coordenador"]), container, pagina)
                
            else:
                alterarCampoPreenchido("Coordenador", str(row["Coordenador"]), container, pagina)
                
        if "Gerente" in df.columns:
            if container.locator("div").filter(has_text="Gerente").locator("text=Clique aqui para adicionar").is_visible():
                
                novoDado("Gerente", str(row["Gerente"]), container, pagina)
                
            else:
                alterarCampoPreenchido("Gerente", str(row["Gerente"]), container, pagina)
          
        if "Concluir" in df.columns:
            pagina.get_by_test_id("open-record-infos-title-status-btn").click()
            pagina.get_by_test_id("status-316811481").get_by_test_id("pstyle-anchor").wait_for()
            pagina.get_by_test_id("status-316811481").get_by_test_id("pstyle-anchor").click()
           
        if "Admissao" in df.columns:
            if container.locator("div").filter(has_text="Admissao").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("Admissao", str(row["Admissao"]), container, pagina)
            else:
                alterarCampoPreenchido("Admissao", str(datetime.strptime(str(row["Admissao"]), "%d/%m/%Y").strftime("%Y-%m-%d")), container, pagina)
        
        if "OBSERVAÇÃO" in df.columns:
            if container.locator("div").filter(has_text="OBSERVAÇÃO").locator("text=Clique aqui para adicionar").is_visible():
                
                novoDado("OBSERVAÇÃO", str(row["OBSERVAÇÃO"]), container, pagina)
                
            else:
                alterarCampoPreenchido("OBSERVAÇÃO", str(row["OBSERVAÇÃO"]), container, pagina)
        
        if "Status" in df.columns:
            if container.locator("div").filter(has_text="* Status").locator("text=Clique aqui para adicionar").is_visible():
                novoDado("* Status", str(row["Status"]), container, pagina)
            else:
                alterarCampoPreenchido("* Status", str(row["Status"]), container, pagina)


        time.sleep(1)
        pagina.get_by_test_id("open-record-btn-close").click()
        df.drop(index=index, inplace=True)
        df.to_excel(nomeDoArquivo, index=False)