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

tabela = load_workbook('notas.xlsx')

aba_ativa = tabela.active
requisicoes = []
with sync_playwright() as p:
  browser = p.chromium.launch(channel="chrome",headless=False)
  page = browser.new_page()
  page.goto(site)

  # LOGIN ME
  page.locator('xpath=//*[@id="LoginName"]').fill(usuario_me)
  page.locator('xpath=//*[@id="RAWSenha"]').fill(senha_me)
  page.locator('xpath=//*[@id="SubmitAuth"]').click()
  page.wait_for_timeout(1)

  # ANALISA STATUS DA REQUISIÇÃO E ATUALIZA PLANILHA
  for celula in aba_ativa['H']:
    if celula.value == 'Pendente':
      linha = celula.row
      reqPendente = aba_ativa[f'B{linha}'].value
      page.goto(f'https://www.me.com.br/VerRequisicaoWF.asp?Req={reqPendente}&SuperCleanPage=false&Origin=home')
      statusRequisicao = page.locator('//*[@id="formRequest"]/div/div[2]/div[2]/p[2]/span[2]').inner_html().strip()
      if statusRequisicao == 'APROVADO':
        #CRIAR PRE-PEDIDO
        if aba_ativa[f'E{linha}'].value == None:
          aba_ativa[f'E{linha}'] = date.today().strftime('%d/%m/%Y')

tabela.save('Tabelateste.xlsx')


