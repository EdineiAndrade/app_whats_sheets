
import time
import datetime
import gspread
from seleniumbase import Driver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PrettyColorPrinter import add_printer

#------Conexão com o Google Sheets
ID_SHEET = '1oe5cwMMjbsBYVTXRNQPTOpVq6qIqdRtZRCf3vpaSQTg'
gc = gspread.service_account(filename='key.json')
sh = gc.open_by_key(ID_SHEET)
ws = sh.worksheet('Contatos')   

plan_01 = sh.get_worksheet(0)
tempo_limite = 3600 
linha = 2

coluna = plan_01.col_values(2)  # 1 representa a coluna B
# Conte o número de linhas preenchidas (não vazias)
total_linhas_preenchidas = len([valor for valor in coluna if valor.strip()])
total = total_linhas_preenchidas
# Obter a data e hora atual
data_h = datetime.datetime.now()
add_printer(1)
driver = Driver(uc=True)
# Abrir o Whatsapp com o chrome
driver.get("https://web.whatsapp.com/")
#esperar conectar
time.sleep(10)
while True:
    try:
        # Tente encontrar o elemento usando o XPath
        elemento = driver.find_element(By.XPATH,'//*[@id="app"]/div/div/div[3]/div[1]/div/div/div[2]/div/canvas')
        print('Aguardando Conectar...')
    except NoSuchElementException:
        # Se o elemento não estiver presente, saia do loop
        print('Conectado!!!')
        break
    # Por exemplo, você pode esperar ou fazer outra coisa
    time.sleep(1)
while True:
    #________VARIÁVEIS DA GOOGLE SHEETS____
    SICAD = plan_01.cell(linha,1).value
    CLIENTE = plan_01.cell(linha,2).value
    CELULAR = plan_01.cell(linha,3).value
    DATA_VECIMENTO = plan_01.cell(linha,4).value
    VALOR = plan_01.cell(linha,5).value
    #sair assim que cliente for ''
    if not CLIENTE:
        break
    MENSSAGEM = "Sr(a) " + CLIENTE + ", Atenção!!! Sua parcela no valor de " + VALOR + " vence em " + DATA_VECIMENTO +"."   
    #____________________________WHATSAPP__________________________________________________________

    driver.get("https://api.whatsapp.com/send?phone=" + CELULAR + "&text=" + MENSSAGEM)
    
    while True:
        try:
            # Verifique se o elemento está presente usando um tempo limite
            elemento_01 = WebDriverWait(driver, tempo_limite).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="action-button"]'))
            )
            # Se o elemento estiver presente, saia do loop
            break
        except:
            # Se o elemento não estiver presente, continue aguardando
            continue
    elemento_01.click() #

    while True:
        try:
            # Verifique se o elemento está presente usando um tempo limite
            elemento_02 = WebDriverWait(driver, tempo_limite).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="fallback_block"]/div/div/h4[2]/a/span'))
            )
            # Se o elemento estiver presente, saia do loop
            break
        except:
            # Se o elemento não estiver presente, continue aguardando
            continue    
    # Localize a tag <a> para obter o link
    elemento_a = driver.find_element(By.XPATH,'//*[@id="fallback_block"]/div/div/h4[2]/a')
    # Obtenha o link (URL) da tag <a>
    link = elemento_a.get_attribute("href")        
    #Navegar para a conversa
    driver.get(link)        
    #Tela de converssas
    time.sleep(5)
    while True:
        try:
            # Verifique se o elemento está presente usando um tempo limite
            elemento_03 = WebDriverWait(driver, tempo_limite).until(
                EC.presence_of_element_located((By.XPATH, '//*[@class="_2xy_p _3XKXx"]/button'))
            )
            # Se o elemento estiver presente, saia do loop
            break
        except:
            # Se o elemento não estiver presente, continue aguardando
            continue
    elemento_03.click() #Clicar em Enviar Menssagem    
    data_h = datetime.datetime.now()
    plan_01.update("F"+ str(linha),"Cliente Avisado: " + str(data_h))#Atualizar planilha do Google Sheets  
    time.sleep(1)       
    restante = total - linha
    print("|Cliente Avisado: " + CLIENTE + "| Linha: " + str(linha) + "| Faltam: " + str(restante))    
    linha += 1
print("Processo Concluído!!!")
print(str(total-1) + " clientes Avisados.")
