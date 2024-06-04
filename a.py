from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from time import sleep

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=options)

url = 'https://cgugovbr.sharepoint.com/sites/ou-sfc-dg-cgplag/_layouts/15/storman.aspx?root=Documentos%20Compartilhados%2FGeneral%2FApresenta%C3%A7%C3%B5es%2FAntigas'

driver.get(url)

sleep(15)

try:
    elementos = driver.find_elements(
        By.PARTIAL_LINK_TEXT, 'Histórico de Versões')

    if elementos:
        for elemento in elementos[:1]:
            link_elemento = elemento.get_attribute('href')

            driver.execute_script("window.open(arguments[0]);", link_elemento)
            sleep(5)

            link_excluir = driver.find_element(
                By.XPATH, "//a[text()='Excluir Todas as Versões']")
            sleep(5)

            link_excluir.click()
            sleep(2)

            alert = Alert(driver)
            alert.accept()

    else:
        print('Não há elementos para excluir')

except:
    print('Link "Histórico de Versões" não encontrado na página')
