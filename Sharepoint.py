import pandas as pd
from funcoes import FuncoesNumericas, FuncoesDataFrame
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from time import sleep


def ordenar_arquivos(data_frame: pd.DataFrame):
    """Função para ordenar os arquivos com base no espaço ocupado"""

    dados = pd.read_excel('consulta_sharepoint.xlsx').convert_dtypes()

    arquivos = data_frame[data_frame['Tipo de Item'] == 'Item']

    # Agrupado por arquivos
    arquivos_ordenados = arquivos.sort_values(
        by='Tamanho do Arquivo', ascending=False)

    arquivos_ordenados['Tamanho do Arquivo'] = arquivos_ordenados['Tamanho do Arquivo'].apply(
        FuncoesNumericas.converter_byte)

    FuncoesDataFrame.gera_excel(
        arquivos_ordenados, 'pastas_sharepoint')


def gerar_link(caminho) -> str:
    """Função para gerar os links das pastas

    Chamada da função: dataframe.apply(lambda row: teste(row['Caminho'], row['Nome']), axis=1)"""

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


def excluir_versoes(links_pastas: pd.Series):
    """Função para excluir as versões dos arquivos"""

    # Define as opções do navegador
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(options=options)

    url_login = 'https://cgugovbr.sharepoint.com/sites/ou-sfc-dg-cgplag/_layouts/15/storman.aspx'

    driver.get(url_login)
    input('Por favor, faça o login manualmente e, em seguida, pressione Enter para continuar...')

    contador_exclusao = 0
    contador_erro = 0
    total = len(links_pastas)

    # Percorre todos os links das pastas
    for link in links_pastas:
        # Acessa o link da pasta e aguarda 15s
        driver.get(link)
        sleep(2)

        # Tenta encontrar o link 'Histórico de Versões' e aguarda 5
        # print(f'\nProcurando o botão "Histórico de Versões" em {link}...')
        # print('=' * 200)
        try:
            elementos = driver.find_elements(
                By.PARTIAL_LINK_TEXT, 'Histórico de Versões')
            # sleep(5)

            # Se encontrou, acessa o link para excluir as versões
            if elementos:
                # Percorre todos os elementos encontrados
                for elemento in elementos:
                    # Procura o 'href' do link e aguarda 5s
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

                    # Espera mais 5 segundos para garantir que a nova aba seja carregada
                    # sleep(5)

                    # Procura o link 'Excluir Todas as Versões' e aguarda 5s
                    try:
                        link_excluir = driver.find_element(
                            By.XPATH, "//a[@accesskey='X']")
                        if link_excluir:
                            # Clica no botão 'Excluir Todas as Versões' e aguarda 5s
                            link_excluir.click()
                            # sleep(5)

                            # Lida com o popup de confirmação
                            alert = Alert(driver)
                            alert.accept()

                            # Incrementa o contador de exclusões
                            contador_exclusao += 1

                    except Exception as e:
                        print(
                            f'\nLink "Excluir Todas as Versões" não encontrado? {e}')
                        print('=' * 200)

                        # Incrementa o contador de erros
                        contador_erro += 1

                    # Fecha a nova aba
                    driver.close()

                    # Volta para a aba principal
                    driver.switch_to.window(handles[0])

            else:
                print('\nLink "Histórico de Versões" não foi encontrado na página')
                print('=' * 200)

        except TimeoutException:
            print(f'\nTimeout ao tentar acessar {link}')
            print('=' * 200)

        porcentagem = calcula_porcentagem(
            contador_exclusao, contador_erro, total)

        print(porcentagem)

    input('\nPressione Enter para encerrar a execução...')
    print('=' * 200)

    driver.quit()


arquivos = pd.read_excel('arquivos_sharepoint.xlsx').convert_dtypes()

links = list(arquivos['Link'].unique())

excluir_versoes(links_pastas=links)
