from openpyxl import load_workbook
import os
from selenium import webdriver as opSeleneium
from selenium.webdriver.common.keys import Keys
import pyautogui as tempoEspera
import pyautogui as posicaoMouse
import pyautogui as login

caminhoArquivo = "C:\\Users\\SylviaSusiBezerraMis\\Desktop\\RPA.p\\Novo Formulario 2\\Emitir2.xlsx"

plan_open = load_workbook(filename=caminhoArquivo)

sheet_selecionada = plan_open['Planilha1']

for i in range(2, len(sheet_selecionada['A']) + 1):
    
    user=sheet_selecionada['A%s'% i].value
    descricao=sheet_selecionada['B%s'% i].value
    contact=sheet_selecionada['C%s'% i].value
    telephone=sheet_selecionada['D%s'% i].value
    address=sheet_selecionada['E%s'% i].value#
    city=sheet_selecionada['F%s'% i].value
    state=sheet_selecionada['G%s'% i].value
    zipcode=sheet_selecionada['H%s'% i].value
    data=sheet_selecionada['I%s'% i].value
    equipamento=sheet_selecionada['J%s'% i].value
    
    posicaoMouse.click(x=1291, y=1054)
    tempoEspera.sleep(4)
    login.press('pagedown')
    tempoEspera.sleep(1)
    #Primeira opcao
    #navegadorform.find_element_by_xpath('//*[@id="mainpage"]/div[2]/div/div[1]/span/span/span/ul/li/div/div[3]/div[2]/span/span/img').click()
    #posicaoMouse.click(x=1250, y=323)
    posicaoMouse.click('C:\\Users\SylviaSusiBezerraMis\\Downloads\\botao1.PNG')
    tempoEspera.sleep(5)
    #opcaoNH
    posicaoMouse.click(x=584, y=521)
    tempoEspera.sleep(5)
    #primeiro next
    posicaoMouse.click(x=1336, y=730)
    tempoEspera.sleep(4)
    #opnh
    posicaoMouse.click(x=824, y=403)
    login.typewrite(user)
    tempoEspera.sleep(8)
    #ponto escolha de user
    posicaoMouse.click(x=904, y=456)
    tempoEspera.sleep(8)
    #calendario ponto troca
    posicaoMouse.click(x=851, y=527)
    tempoEspera.sleep(4)
    login.typewrite(str(data))
    tempoEspera.sleep(4)
    #posicaoMouse.click(x=875, y=533)
    #tempoEspera.sleep(9)
    #posicaoMouse.click(x=896, y=686)
    tempoEspera.sleep(9)
    #descricao
    posicaoMouse.click(x=791, y=663)
    tempoEspera.sleep(4)
    login.typewrite(descricao)
    tempoEspera.sleep(3)
    #next
    posicaoMouse.click(x=1344, y=846)
    tempoEspera.sleep(4)
    #escolhendoos
    posicaoMouse.click(x=957, y=539)
    tempoEspera.sleep(4)
    posicaoMouse.click(x=1399, y=861)
    tempoEspera.sleep(60)
    #escolhendot480<-----------------------
    #posicaoMouse.click(x=502, y=703)
    posicaoMouse.hotkey('ctrl','f')
    tempoEspera.sleep(2)
    login.typewrite(equipamento)
    tempoEspera.sleep(5)
    if equipamento == "T460":
        posicaoMouse.click('C:\\Users\\SylviaSusiBezerraMis\\Downloads\\botao3.PNG')
    elif equipamento == "P15":
        posicaoMouse.click('C:\\Users\\SylviaSusiBezerraMis\\Downloads\\botao4.PNG')
    tempoEspera.sleep(8)
    posicaoMouse.click(x=1342, y=846)
    tempoEspera.sleep(5)
    #escolhendoendereco
    posicaoMouse.click(x=621, y=461)
    tempoEspera.sleep(4)
    #next
    posicaoMouse.click(x=1368, y=865)
    tempoEspera.sleep(8)
    #Shipping
    #Ship to contact:
    posicaoMouse.click(x=850, y=435)
    tempoEspera.sleep(4)
    login.typewrite(contact)
    tempoEspera.sleep(4)
    #Ship to telephone:
    posicaoMouse.click(x=806, y=470)
    tempoEspera.sleep(4)
    login.typewrite(telephone)
    tempoEspera.sleep(5)
    #Ship to Address:
    posicaoMouse.click(x=826, y=522) 
    tempoEspera.sleep(5)
    login.typewrite(address)
    tempoEspera.sleep(5)
    #Ship to City:
    posicaoMouse.click(x=816, y=556)
    tempoEspera.sleep(5)
    login.typewrite(city)
    tempoEspera.sleep(5)
    #Ship to State:
    posicaoMouse.click(x=840, y=610)
    tempoEspera.sleep(3)
    login.typewrite(state)
    tempoEspera.sleep(3)
    #Ship to Postal:
    posicaoMouse.click(x=864, y=646)
    tempoEspera.sleep(3)
    login.typewrite(zipcode)
    tempoEspera.sleep(3)
    posicaoMouse.click(x=1476, y=207)
    tempoEspera.sleep(3)
    posicaoMouse.click(x=1291, y=1054)
#Salvando Note
posicaoMouse.click(x=1379, y=1055)
tempoEspera.sleep(2)
posicaoMouse.click(x=984, y=90)
tempoEspera.sleep(2)
posicaoMouse.click(x=1020, y=209)
login.alert("O codigo Finalizou")
