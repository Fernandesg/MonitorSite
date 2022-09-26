from openpyxl import Workbook, load_workbook
from datetime import date
from playwright.sync_api import sync_playwright

passwords = open('credenciais.txt', 'r')
login = []

for linhas in passwords:
    linhas = linhas.strip()
    login.append(linhas)
usuario_me = login[0][14:-1]
senha_me = login[1][12:-1]
site = login[2][8:-1]

tabela = load_workbook('notas.xlsm', data_only=True)
aba_ativa = tabela['REQUISIÇÕES PENDENTES']

with sync_playwright() as p:
  browser = p.chromium.launch(channel="chrome", headless=False)
  page = browser.new_page()
  page.goto(site)

  # LOGIN ME
  page.locator('xpath=//*[@id="LoginName"]').fill(usuario_me)
  page.locator('xpath=//*[@id="RAWSenha"]').fill(senha_me)
  page.locator('xpath=//*[@id="SubmitAuth"]').click()
  page.wait_for_timeout(1)

  # ANALISA STATUS DA REQUISIÇÃO E ATUALIZA PLANILHA
  for celula in aba_ativa['I']:
    linha = celula.row
    if celula.value == 'Pendente' and aba_ativa[f'E{linha}'].value == None:
        cnpj = aba_ativa[f'D{linha}'].value
        reqPendente = aba_ativa[f'B{linha}'].value
        page.goto(f'https://www.me.com.br/DO/Request/Home.mvc/Show/{reqPendente}')
        statusRequisicao = page.locator('//*[@id="formRequest"]/div/div[2]/div[2]/p[2]/span[2]').inner_html().strip()
        tituloReq = page.locator('xpath=//*[@id="formRequest"]/div/div[2]/div[1]/p[1]').inner_html().strip()

        if statusRequisicao == 'APROVADO':
          #CRIAR PRE-PEDIDO
          page.locator('xpath=//*[@id="btnEmergency"]').click()
          page.locator('xpath=/html/body/div[1]/div[3]/div/button[1]/span').click()
          page.locator('xpath=//*[@id="MEComponentManager_MEButton_2"]').click()
          page.locator('xpath=//*[@id="CGC"]').fill(cnpj)
          page.keyboard.press('Enter')
          page.locator('xpath=//*[@id="grid"]/div[2]/table/tbody/tr/td[1]/div/input').click()
          page.locator('xpath=//*[@id="btnSalvarSelecao"]').click()
          page.locator('xpath=//*[@id="btnVoltarPrePedEmergencial"]').click()
          page.locator('xpath=//*[@id="Resumo"]').fill(tituloReq)
          filiaisPrePedido = page.locator('//select[@name="LocalCobranca"]').inner_html().split('\n')
          indice = [i for i, s in enumerate(filiaisPrePedido) if 'VERO SANTO ANTONIO DA PATRULHA' in s][0]
          page.locator('//select[@name="LocalCobranca"]').select_option(index=indice-1)
          page.locator('xpath=//*[@id="DataEntrega"]').fill(date.today().strftime('%d/%m/%Y'))
          page.locator('xpath=//*[@id="MEComponentManager_MEButton_3"]').click()
          page.locator('xpath=/html/body/main/form[2]/table[3]/tbody/tr[1]/td/input[1]').click()
          page.locator('xpath=//*[@id="MEComponentManager_MEButton_2"]').click()
          page.locator('xpath=//*[@id="MEComponentManager_MEButton_2"]').click()
          page.locator('xpath=//*[@id="formItemStatusHistory"]/div/b[1]/a').click()
          numPrePedido = page.locator('xpath=/html/body/main/div/div[1]/div[1]/p').inner_html().strip()
          statusPrePedido = page.locator('xpath=/html/body/main/div/div[1]/div[2]/div[2]/p[1]/span[2]').inner_html().strip()
          aba_ativa[f'F{linha}'] = date.today().strftime('%d/%m/%Y')
          aba_ativa[f'E{linha}'] = numPrePedido
          
    if celula.value == 'Pendente' and aba_ativa[f'E{linha}'].value != None:
      print('Possui pré pedido')
      prePedidoPendente = aba_ativa[f'E{linha}'].value
      print(prePedidoPendente)
      page.goto(f'https://www.me.com.br/VerPrePedidoWF.asp?Pedido={prePedidoPendente}&SuperCleanPage=false&Origin=home')
      statusPrePedido = page.locator('xpath=/html/body/main/div/div[1]/div[2]/div[2]/p[1]/span[2]').inner_html().strip()[:8]
      if statusPrePedido == 'APROVADO':
        numPedidoSAP = page.locator('xpath=/html/body/main/div/div[1]/div[1]/p[1]').inner_html().strip()
        aba_ativa[f'G{linha}'] = numPedidoSAP

tabela.save('Tabelateste.xlsx')


