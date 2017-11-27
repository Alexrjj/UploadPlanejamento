﻿from selenium.common.exceptions import NoSuchElementException
import os
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import openpyxl

#  Acessa os dados de login fora do script, salvo numa planilha existente, para proteger as informações de credenciais
dados = openpyxl.load_workbook('C:\\gomnet.xlsx')
login = dados['Plan1']
url = 'http://gomnet.ampla.com/'
url2 = 'http://gomnet.ampla.com/Upload.aspx?numsob='
username = login['A1'].value
password = login['A2'].value


driver = webdriver.Chrome()
if __name__ == '__main__':
    driver.get(url)
    # Faz login no sistema
    uname = driver.find_element_by_name('txtBoxLogin')
    uname.send_keys(username)
    passw = driver.find_element_by_name('txtBoxSenha')
    passw.send_keys(password)
    submit_button = driver.find_element_by_id('ImageButton_Login').click()

    # Modifica os campos necessários e envia o anexo de cada sob contido nos arquivos txt.
    with open('sobs.txt') as sobs:
        for sob in sobs:
            sob = sob.strip()
            driver.get(url2 + sob.partition("_")[0])
            try:  # Verifica se a sob foi digitada incorretamente.
                erro = driver.find_element_by_xpath('*//tr/td[contains(text(),'
                                                    '"Não existem dados para serem exibidos.")]')
                if erro.is_displayed():
                    print("Sob " + sob.partition("_")[0] + " não encontrada. Favor verificar.")
            except NoSuchElementException:
                try:  # Verifica se o arquivo já foi anexado.
                    anexo = driver.find_element_by_xpath(
                        "*//a[contains(text(), '" + sob + ".PDF""')]")
                    if anexo.is_displayed():
                        print("Arquivo " + sob + ".PDF já foi anexado.")
                except NoSuchElementException:
                    # Preenche o campo "Descrição" com "PLANEJAMENTO"
                    atividade = driver.find_element_by_id('txtBoxDescricao')
                    atividade.send_keys('PLANEJAMENTO')
                    # Identifica o menu " Categoria de Documento" e seleciona a opção "EXECUCAO"
                    categoria = Select(driver.find_element_by_id('drpCategoria'))
                    categoria.select_by_visible_text('EXECUCAO')
                    # Identifica o menu " Tipo de Documento" e seleciona a opção "OUTROS"
                    documento = Select(driver.find_element_by_id('DropDownList1'))
                    documento.select_by_visible_text('PRÉ APR VISTORIA DE OBRAS')
                    driver.find_element_by_id('fileUPArquivo').send_keys(os.getcwd() + "\\" + sob + ".PDF")
                    driver.find_element_by_id('Button_Anexar').click()
                    try:
                        # Verifica se o arquivo foi anexado com êxito
                        status = driver.find_element_by_xpath('//*[@id="txtBoxMessage"][contains(text(),'
                                                              '"Arquivo salvo com sucesso.")]')
                        if status.is_displayed():
                            print(sob + " anexado com sucesso.")
                    except NoSuchElementException:
                        log = open("log.txt", "a")
                        log.write(sob + " não foi anexado.")
                        log.close()
                        continue
        print("Fim da execução.")
