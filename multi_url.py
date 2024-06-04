from selenium import webdriver
import time

# Configurar as opções do Chrome
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)

# Inicializar o driver do Chrome com as opções
driver = webdriver.Chrome(options=options)

# Lista de URLs que você quer abrir
urls = [
    'https://www.example.com',
    'https://www.google.com',
    'https://www.wikipedia.org'
]

# Abrir cada URL na mesma janela do Chrome
for url in urls:
    driver.get(url)
    time.sleep(2)  # Espera um pouco para carregar completamente a página

# Esperar até que o usuário encerre o programa
input("Pressione Enter para sair...")

# Fechar o navegador ao sair do programa
driver.quit()
