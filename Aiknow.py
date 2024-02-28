from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Solicitar o link da página ao usuário
url = ("https://docs.google.com/forms/d/e/1FAIpQLScQJPKbrdpYHtW8MdudlLRVyZYwZGGGq3eUsE2mPi1bH_Xqrg/viewform")

# Abrir o navegador (neste caso, Chrome)
driver = WebDriver()

# Abrir o formulário
driver.get(url)

# Função para preencher os campos
def preencher_campos(driver, email, number, tickets, Dcriado, Dfechado):
    wait = WebDriverWait(driver, 10)
    
    # Preencher campo de E-mail
    campo_email = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[aria-label='Seu e-mail']")))
    campo_email.clear()
    campo_email.send_keys(email)
    

    # Preencher campo de Numero
    campo_number = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[aria-labelledby='i5']")))
    campo_number.clear()
    campo_number.send_keys(number)
   
  
    # Preencher campo de Tickets
    campo_tickets = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[aria-labelledby='i9']")))
    campo_tickets.clear()
    campo_tickets.send_keys(tickets)
    
    
    # Preencher campo de Hora que a loja chamou
    campo_Dcriado = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[aria-labelledby='i56']")))
    campo_Dcriado.clear()
    campo_Dcriado.send_keys(Dcriado)
    
    
    # Preencher campo de Hora de conclusão do chamado
    campo_Dfechado = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[aria-labelledby='i61']")))
    campo_Dfechado.clear()
    campo_Dfechado.send_keys(Dfechado)
    

def selecionar_closed(driver):
    wait = WebDriverWait(driver, 10)
    opcao_closed = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[aria-label='Closed']")))
    opcao_closed.click()
    opcao_suporte = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[aria-label='Suporte remoto 1']")))
    opcao_suporte.click()
    

def selecionar_opcoes(driver, nomes_listas_suspensas):
    # Localizar todas as listas suspensas
    time.sleep(2)
    select_elements = driver.find_elements(By.CLASS_NAME, "ry3kXd")
    
    for i, select_element in enumerate(select_elements):
        nome_lista = nomes_listas_suspensas[i]  # Obtém o nome da lista suspensa usando o índice
        
        # Clicar na lista suspensa para abrir as opções
        select_element.click()
        
        # Esperar um curto período para as opções serem carregadas
        time.sleep(2)
        
        while True:
            # Localizar todas as opções dentro da lista suspensa
            options = select_element.find_elements(By.XPATH, "//div[@role='option']")
            
            # Solicitar a entrada do usuário para o texto desejado, usando o nome da lista suspensa
            search_text = input(f"Digite o texto desejado para selecionar na {nome_lista}: ").lower()
            
            # Encontrar a opção mais próxima correspondente ao texto inserido pelo usuário
            selected_option = None
            min_distance = float('inf')  # Definir uma distância mínima inicial como infinito
            
            for option in options:
                option_text = option.text.strip().lower()  # Remove espaços em branco no início e no final e converte para minúsculas
                # Calcular a distância entre a opção e o texto inserido pelo usuário
                distance = abs(len(option_text) - len(search_text))
                if distance < min_distance and search_text in option_text:
                    min_distance = distance
                    selected_option = option
            
            # Selecionar a opção encontrada
            if selected_option:
                selected_option.click()
                print("Opção selecionada com sucesso.")
                break  # Sai do loop while se uma opção válida for selecionada
            else:
                print("Nenhuma opção correspondente encontrada. Por favor, tente novamente.")
        
        # Pausa para dar tempo para o usuário observar a seleção antes de continuar
        time.sleep(2)



nomes_listas_suspensas = ["a origem do chamado", "o BKN da loja", "quem realizou o chamado", "o BKN da loja", "a importância do chamado", "a resolução do chamado"]
selecionar_opcoes(driver, nomes_listas_suspensas)

# Preencher os campos do formulário
email = "vlordaro@driveit.com.br"
number = 48843
tickets = input("Digite o texto para o campo de Tickets: ")
Dcriado = input("Digite a data que o chamado foi aberto: ")
Dfechado = input("Digite a data que o chamado foi encerrado: ")
# Preencher os campos do formulário
preencher_campos(driver, email, number, tickets, Dcriado, Dfechado)

# Chamada da função para selecionar "Closed"
selecionar_closed(driver)

# Aguardar o carregamento do formulário
time.sleep(1)

# Aguardar a entrada do usuário para fechar o navegador
input("Pressione Enter para fechar o navegador...")
driver.quit()
