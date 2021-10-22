from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import openpyxl

navegador = webdriver.Chrome()
book = openpyxl.Workbook()

options = Options()

url = {
    'BRL': '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]',
    'Url_Processador': '//*[@id="root"]/div/div[2]/div/div[2]/div[4]/div[1]/span',
    'Url_Placa_Mae': '//*[@id="root"]/div/div[2]/div/div[2]/div[3]/div[2]/div[1]/span[1]',
    'Url_ram': '//*[@id="root"]/div/div[2]/div/div[2]/div[4]/div[1]/span',
    'Url_ssd': '//*[@id="root"]/div/div[2]/div/div[2]/div[4]/div[1]/span',
    'Url_fonte': '//*[@id="valVista"]',
  
    # Cotação usado somente pois peguei valores do AliExpress, e o mesmo apresenta valores em UYU;
    'Cotacao': 'https://www.google.com/search?client=opera-gx&q=uyu+para+brl&sourceid=opera&ie=UTF-8&oe=UTF-8',
    'Processador': 'https://pt.aliexpress.com/item/4000262737817.html?spm=a2g0s.8937460.0.0.60c52e0eBZnzwy',
    'Placa_mae': 'https://pt.aliexpress.com/item/4000750170401.html?spm=a2g0s.8937460.0.0.60c52e0eBZnzwy',
    'Ram': 'https://pt.aliexpress.com/item/32601791617.html?spm=a2g0s.8937460.0.0.60c52e0eBZnzwy',
    'SSD': 'https://pt.aliexpress.com/item/4001316177223.html?spm=a2g0s.8937460.0.0.60c52e0eBZnzwy',
    'Fonte': 'https://www.terabyteshop.com.br/produto/13329/fonte-gamemax-gp650-650w-80-plus-bronze-pfc-ativo-black',
}

botoes = {
    'botao_ram': '//*[@id="root"]/div/div[2]/div/div[2]/div[7]/div/div[2]/ul/li[4]/div/span',
    'botao_ram2': '//*[@id="root"]/div/div[2]/div/div[2]/div[7]/div/div[1]/ul/li[2]/div/span',
    'botao_ssd': '//*[@id="root"]/div/div[2]/div/div[2]/div[7]/div/div/ul/li[3]/div'
}


def encontrar_elemento(endereco):
    return navegador.find_element_by_xpath(url[endereco]).get_attribute('innerHTML')

def apertar_botao(elemento):
    return navegador.find_element_by_xpath(botoes[elemento]).click()

def navegar(site):
    return navegador.get(url[site])

def verificar_precos():
    navegar('Cotacao')
    cotacao = encontrar_elemento('BRL')

    navegar('Processador')
    pp = encontrar_elemento('Url_Processador')
    valor_processador = pp.split().pop(1)

    navegar('Placa_mae')
    pm = encontrar_elemento('Url_Placa_Mae')
    valor_placamae = pm.split().pop(1)

    navegar('Ram')
    apertar_botao('botao_ram')
    apertar_botao('botao_ram2')
    pr = encontrar_elemento('Url_ram')
    valor_ram = pr.split().pop(1)

    navegar('SSD')
    apertar_botao('botao_ssd')
    ps = encontrar_elemento('Url_ssd')
    valor_ssd = ps.split().pop(1)

    navegar('Fonte')
    pf = encontrar_elemento('Url_fonte')
    valor_fonte = pf.split().pop(1)

    # Planilha;
    pagina = book['Sheet']
    pagina.append(['Cotação do dia: ', cotacao])
    pagina.append(['Processador E5 2640 V3', valor_processador])
    pagina.append(['Placa Mãe x99 Machinist', valor_placamae])
    pagina.append(['Memória RAM 8GB 1866MHz', valor_ram])
    pagina.append(['SSD 512GB', valor_ssd])
    pagina.append(['Fonte Gamemax 650w', valor_fonte])
    book.save('Planilha.xlsx')
    
    print('Processo de escritura na Planilha concluído.')
    navegador.close()


verificar_precos()
