import pandas as pd
# from funcoes import FuncoesNumericas, FuncoesDataFrame
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from time import sleep
import PySimpleGUI as sg


def ordenar_arquivos():
    """Função para ordenar os arquivos com base no espaço ocupado"""

    data_frame = pd.read_excel('consulta_sharepoint.xlsx').convert_dtypes()

    arquivos = data_frame[data_frame['Tipo de Item'] == 'Item']

    arquivos_ordenados = arquivos.sort_values(
        by='Tamanho do Arquivo', ascending=False)

    '''arquivos_ordenados['Tamanho do Arquivo'] = arquivos_ordenados['Tamanho do Arquivo'].apply(
        FuncoesNumericas.converter_byte)

    FuncoesDataFrame.gera_excel(
        arquivos_ordenados, 'pastas_sharepoint')'''


def gerar_link(caminho) -> str:
    """Função para gerar os links das pastas"""

    base_link = 'https://cgugovbr.sharepoint.com/'
    inicio_link = caminho[:23]
    meio_link = '_layouts/15/storman.aspx?root='
    final_link = quote(str(caminho)[23:])
    # local_pasta = quote(str(nome))

    link = f'{base_link}{inicio_link}{meio_link}{final_link}'

    return link


def calcula_porcentagem(sucesso: int, erro: int, total: int) -> str:
    if total == 0:
        raise ValueError('O valor total não pode ser zero.')

    calculo_sucesso = (sucesso / total) * 100
    calculo_erro = (erro / total) * 100

    porcentagem = f'''Já foram realizadas {sucesso} exclusões, ou seja, {calculo_sucesso:.2f}% do total {total}. Houveram {erro} erros, ou seja, {calculo_erro:.2f}% do total'''

    return porcentagem


def excluir_versoes(links_pastas: list, total: int):
    """Função para excluir as versões dos arquivos"""

    # Define as opções do navegador
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(options=options)

    url_login = 'https://cgugovbr.sharepoint.com/sites/ou-sfc-dg-cgplag/_layouts/15/storman.aspx'

    driver.maximize_window()
    driver.get(url_login)

    sg.popup('Pressione OK após fazer o login', title='Login')

    contador_exclusao = 0
    contador_erro = 0

    with open('log.txt', 'w', encoding='utf-8') as log:
        # Percorre todos os links das pastas
        for link in links_pastas:
            # Acessa o link da pasta e aguarda 15s
            driver.get(link)
            sleep(2)

            # Tenta encontrar o link 'Histórico de Versões' e aguarda 5
            try:
                elementos = driver.find_elements(
                    By.PARTIAL_LINK_TEXT, 'Histórico de Versões')
                # sleep(5)

                # Se encontrou, acessa o link para excluir as versões
                if elementos:
                    # Percorre todos os elementos encontrados
                    for elemento in elementos:
                        # Procura o 'href' do link
                        link_elemento = elemento.get_attribute('href')

                        # Abre o link em uma nova aba
                        driver.execute_script(
                            "window.open(arguments[0]);", link_elemento)

                        # Espera até que a nova aba esteja disponível
                        WebDriverWait(driver, 10).until(
                            EC.number_of_windows_to_be(2))

                        # Obtém todas as alças (handles) das abas
                        handles = driver.window_handles
                        # Alterna para a nova aba
                        driver.switch_to.window(handles[1])

                        # Procura o link 'Excluir Todas as Versões'
                        try:
                            link_excluir = driver.find_element(
                                By.XPATH, "//a[@accesskey='X']")
                            if link_excluir:
                                # Clica no botão 'Excluir Todas as Versões'
                                link_excluir.click()

                                # Lida com o popup de confirmação
                                alert = Alert(driver)
                                alert.accept()

                                # Incrementa o contador de exclusões
                                contador_exclusao += 1

                        except Exception as erro:
                            mensagem = f'\n{"-" * 200}\n{erro}'
                            print(mensagem)
                            log.writelines(mensagem)

                            # Incrementa o contador de erros
                            contador_erro += 1

                        # Fecha a nova aba
                        driver.close()

                        # Volta para a aba principal
                        driver.switch_to.window(handles[0])

            except Exception as erro:
                mensagem = f'\n{"-" * 200}\n{erro}'
                print(mensagem)
                log.writelines(mensagem)

            porcentagem = calcula_porcentagem(
                contador_exclusao, contador_erro, total)

            mensagem = f'\n{"-" * 200}\n{porcentagem}'
            print(mensagem)
            log.writelines(mensagem)

        try:
            url_lixeira = 'https://cgugovbr.sharepoint.com/sites/ou-sfc-dg-cgplag/_layouts/15/AdminRecycleBin.aspx?view=5'
            driver.get(url_lixeira)

            # Esvazia a lixeira
            lixeira = driver.find_element(By.NAME, 'Esvaziar lixeira')

            if lixeira:
                lixeira.click()

                try:
                    btn_confirma = driver.find_element(
                        By.CLASS_NAME, 'od-Button-label')

                    if btn_confirma:
                        btn_confirma.click()

                except Exception as erro:
                    mensagem = f'\n{"-" * 200}\n{erro}'
                    print(mensagem)
                    log.writelines(mensagem)

        except Exception as erro:
            mensagem = f'\n{"-" * 200}\n{erro}'
            print(mensagem)
            log.writelines(mensagem)

    sg.popup('Pressione OK para encerrar a execução')
    driver.quit()


print(f'\n{"-" * 200}\nLendo o arquivo...')
arquivos = pd.read_excel('arquivos_sharepoint.xlsx').convert_dtypes()

links = arquivos['Link'].unique()
total_links = len(arquivos['Link'])

excluir_versoes(links_pastas=links, total=total_links)
