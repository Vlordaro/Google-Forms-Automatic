from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time
import asynctkinter as tk 
import os
import asyncio
import gspread
from google.oauth2.service_account import Credentials
from selenium.common.exceptions import TimeoutException
import tkinter as tk
from tkinter import messagebox
import telegram

# Definir as credenciais diretamente no código
creds = { credencial google
}

credentials = Credentials.from_service_account_info(creds, scopes=['https://www.googleapis.com/auth/spreadsheets'])

# Autenticar com o Google Sheets
gc = gspread.authorize(credentials)

# Abrir a primeira guia da planilha pelo ID
sh = gc.open_by_key('KeyAPI').sheet1

def get_next_number(sh):
    # Obter todos os valores na coluna 'number'
    numbers = sh.col_values(3)  # Supondo que a coluna 'number' seja a segunda coluna
    # Converter valores para inteiros e obter o maior número
    numbers = [int(num) for num in numbers if num.isdigit()]
    return max(numbers) + 1 if numbers else 1

chrome_driver_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chromedriver.exe")
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--enable-chrome-browser-cloud-management")

service = webdriver.chrome.service.Service(chrome_driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

def preencher_formulario():
    # Solicitar o link da página ao usuário
    url = "seu Url"

    # Abrir o navegador (neste caso, Chrome)
    driver.get(url)

    # Função para preencher os campos
    def preencher_campos(driver, email, number, tickets, Dcriado, primeira_hora, Dfechado, segunda_hora):
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
    
        # Validação da data
        while True:
            if len(Dcriado) == 10 and Dcriado[2] == '-' and Dcriado[5] == '-' and Dcriado[:2].isdigit() and Dcriado[3:5].isdigit() and Dcriado[6:].isdigit():
                confirmacao = input(f"A data de criação que você digitou é {Dcriado}. Está correto? (S/N): ").upper()
                if confirmacao in ("S", "SIM"):
                    break
                elif confirmacao in ("N", "NAO"):
                    Dcriado = input("Digite a data de criação novamente (DD-MM-AAAA): ")
                else:
                    print("Resposta inválida. Por favor, digite S ou N.")
            else:
                Dcriado = input("Formato de data inválido. Use DD-MM-AAAA (ex: 10-06-2024).")

# Validação da primeira hora
        while True:
            if len(primeira_hora) == 5 and primeira_hora[2] == ':' and primeira_hora[:2].isdigit() and primeira_hora[3:].isdigit():
                confirmacao = input(f"A hora que você digitou é {primeira_hora}. Está correto? (S/N): ").upper()
                if confirmacao in ("S", "SIM"):
                    break
                elif confirmacao in ("N", "NAO"):
                    primeira_hora = input("Digite a hora novamente (HH:MM): ")
                else:
                    print("Resposta inválida. Por favor, digite S ou N.")
            else:
                primeira_hora = input("Formato de hora inválido. Use HH:MM (ex: 10:30).")

        while True:
            if len(Dfechado) == 10 and Dfechado[2] == '-' and Dfechado[5] == '-' and Dfechado[:2].isdigit() and Dfechado[3:5].isdigit() and Dfechado[6:].isdigit():
                confirmacao = input(f"A data de fechamento que você digitou é {Dfechado}. Está correto? (S/N): ").upper()
                if confirmacao in ("S", "SIM"):
                    break
                elif confirmacao in ("N", "NAO"):
                    Dfechado = input("Digite a data de criação novamente (DD-MM-AAAA): ")
                else:
                    print("Resposta inválida. Por favor, digite S ou N.")
            else:
                Dfechado = input("Formato de data inválido. Use DD-MM-AAAA (ex: 10-06-2024).")

# Validação da primeira hora
        while True:
            if len(segunda_hora) == 5 and segunda_hora[2] == ':' and segunda_hora[:2].isdigit() and segunda_hora[3:].isdigit():
                confirmacao = input(f"A hora que você digitou é {segunda_hora}. Está correto? (S/N): ").upper()
                if confirmacao in ("S", "SIM"):
                    break
                elif confirmacao in ("N", "NAO"):
                    segunda_hora = input("Digite a hora novamente (HH:MM): ")
                else:
                    print("Resposta inválida. Por favor, digite S ou N.")
            else:
                segunda_hora = input("Formato de hora inválido. Use HH:MM (ex: 10:30).")

        campo_Dcriado = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[aria-labelledby='i56']")))
        campo_Dcriado.clear()
        campo_Dcriado.send_keys(Dcriado)

        campo_primeira_hora = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(text(), 'Hora que a loja chamou')]//following::input[@aria-label='Hora'][1]")))
        campo_primeira_hora.clear()
        campo_primeira_hora.send_keys(primeira_hora[:2])

        campo_primeiro_minuto = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(text(), 'Hora que a loja chamou')]//following::input[@aria-label='Minuto'][1]")))
        campo_primeiro_minuto.clear()
        campo_primeiro_minuto.send_keys(primeira_hora[3:])

        campo_Dfechado = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[aria-labelledby='i61']")))
        campo_Dfechado.clear()
        campo_Dfechado.send_keys(Dfechado)

        campo_segunda_hora = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(text(), 'Hora de conclusão do chamado')]//following::input[@aria-label='Hora']")))
        campo_segunda_hora.clear()
        campo_segunda_hora.send_keys(segunda_hora[:2])

        campo_segundo_minuto = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(text(), 'Hora de conclusão do chamado')]//following::input[@aria-label='Minuto']")))
        campo_segundo_minuto.clear()
        campo_segundo_minuto.send_keys(segunda_hora[3:])
                

    def preencher_ocorre(driver, ocorrido): #Preencher o ocorrido
        wait=webdriver(driver,10)
        campo_ocorre = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[aria-labelledby='i5']")))
        campo_ocorre.clear()
        campo_ocorre.send_keys(ocorrido)

    def selecionar_closed(driver):
        wait = WebDriverWait(driver, 10)
        opcao_closed = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[aria-label='Closed']")))
        opcao_closed.click()
        opcao_suporte = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[aria-label='Suporte remoto 1']")))
        opcao_suporte.click()

    def selecionar_opcoes(driver, nomes_listas_suspensas):
        # Localizar todas as listas suspensas
        time.sleep(1)
        select_elements = driver.find_elements(By.CLASS_NAME, "ry3kXd")
        global resolucao_escolhida  # Declara que a variável é global
        resolucao_escolhida = ""  # Inicializa a variável aqui
        global bkn_escolhido
        bkn_escolhido = ""
        global User
        User = ""
        
        
        
        
        

        for i, select_element in enumerate(select_elements):
            nome_lista = nomes_listas_suspensas[i]  # Obtém o nome da lista suspensa usando o índice
            
            # Clicar na lista suspensa para abrir as opções
            select_element.click()
            
            # Esperar um curto período para as opções serem carregadas
            time.sleep(1)

            if nome_lista.lower() == "o bkn da loja": 
                options = select_element.find_elements(By.XPATH, "//div[@role='listbox']//div[@role='option']") 
                while True:
                    search_text = input(f"Digite {nome_lista}: ").lower()
                    matching_options = [option for option in options if search_text in option.text.strip().lower()]
                    if len(matching_options) == 1:
                        driver.execute_script("arguments[0].scrollIntoView(true);", matching_options[0])
                        matching_options[0].click()

                        wait = WebDriverWait(driver, 10)
                        matching_option = wait.until(EC.visibility_of(matching_options[0]))

                        print("Opção 'o BKN da loja' selecionada com sucesso.")
                        # Armazena o valor da opção selecionada em 'bkn_escolhido'
                        bkn_escolhido = matching_option.get_attribute("data-value") 
                        break
                    else:
                        print(f"Nenhuma opção correspondente encontrada para '{search_text}'. Por favor, tente novamente.")
            elif nome_lista.lower() == "quem realizou o chamado": 
                options = select_element.find_elements(By.XPATH, "//div[@role='listbox']//div[@role='option']") 
                while True:
                    search_text = input(f"Digite {nome_lista}: ").lower()
                    matching_options = [option for option in options if search_text in option.text.strip().lower()]
                    if len(matching_options) == 1:
                        driver.execute_script("arguments[0].scrollIntoView(true);", matching_options[0])
                        matching_options[0].click()

                        wait = WebDriverWait(driver, 10)
                        matching_option = wait.until(EC.visibility_of(matching_options[0]))

                        print("Opção selecionada com sucesso.")
                        # Armazena o valor da opção selecionada em 'bkn_escolhido'
                        User = matching_option.get_attribute("data-value") 
                        break
                    else:
                        print(f"Nenhuma opção correspondente encontrada para '{search_text}'. Por favor, tente novamente.")

            # Se for a lista suspensa de resolução do chamado, guarde o valor da opção selecionada
            elif nome_lista.lower() == "a resolução do chamado":
                # Localizar todas as opções dentro da lista suspensa
                options = select_element.find_elements(By.XPATH, "//div[@role='listbox']//div[@role='option']")

                while True:
                    # Solicitar a entrada do usuário para o texto desejado, usando o nome da lista suspensa
                    search_text = input(f"Digite {nome_lista}: ").lower()

                    # Encontrar todas as opções correspondentes ao texto inserido pelo usuário
                    matching_options = [option for option in options if search_text in option.text.strip().lower()]

                    if len(matching_options) > 1:
                        print(f"Existem várias opções correspondentes para '{search_text}':")
                        for index, option in enumerate(matching_options):
                            print(f"{index + 1}. {option.text.strip()}")
                        selected_index = input("Digite o número da opção desejada: ")
                        if selected_index.isdigit() and 1 <= int(selected_index) <= len(matching_options):
                            driver.execute_script("arguments[0].scrollIntoView(true);", matching_options[int(selected_index) - 1])
                            matching_options[int(selected_index) - 1].click()
                
                            # Adicione um Explicit Wait aqui, após clicar na opção
                            wait = WebDriverWait(driver, 10)
                            matching_option = wait.until(EC.visibility_of(matching_options[int(selected_index) - 1]))
                
                            print("Opção selecionada com sucesso.")
                            # Obter o valor da opção selecionada
                            resolucao_escolhida = matching_option.get_attribute("data-value")  # Use matching_option aqui
                            break
                        else:
                            print("Opção inválida. Por favor, digite um número válido.")
                    elif len(matching_options) == 1:
                        driver.execute_script("arguments[0].scrollIntoView(true);", matching_options[0])
                        matching_options[0].click()
            
                        # Adicione um Explicit Wait aqui, após clicar na opção
                        wait = WebDriverWait(driver, 10)
                        matching_option = wait.until(EC.visibility_of(matching_options[0]))
            
                        print("Opção selecionada com sucesso.")
                        # Obter o valor da opção selecionada
                        resolucao_escolhida = matching_option.get_attribute("data-value")  # Use matching_option aqui
                        break
                    else:
                        print(f"Nenhuma opção correspondente encontrada para '{search_text}'. Por favor, tente novamente.")
            else:
                while True:
                    # Localizar todas as opções dentro da lista suspensa
                    options = select_element.find_elements(By.XPATH, "//div[@role='listbox']//div[@role='option']")
                    
                    # Solicitar a entrada do usuário para o texto desejado, usando o nome da lista suspensa
                    search_text = input(f"Digite {nome_lista}: ").lower()
                    
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
                        driver.execute_script("arguments[0].scrollIntoView(true);", selected_option)
                        selected_option.click()
                        time.sleep(1)
                        print("Opção selecionada com sucesso.")
                        break  # Sai do loop while se uma opção válida for selecionada
                    else:
                        # Verifica se a entrada do usuário corresponde exatamente a uma opção existente
                        exact_match = [option for option in options if search_text == option.text.strip().lower()]
                        if exact_match:
                            driver.execute_script("arguments[0].scrollIntoView(true);", exact_match[0])
                            exact_match[0].click()
                            time.sleep(1)
                            print("Opção selecionada com sucesso.")
                            break
                        else:
                            print("Nenhuma opção correspondente encontrada. Por favor, tente novamente.")
        
                                       
            
            # Pausa para dar tempo para o usuário observar a seleção antes de continuar
            time.sleep(1)
        
        

    # Preencher os campos do formulário
    email = "vlordaro@driveit.com.br"
    number = get_next_number(sh)
    tickets = input("Digite o texto para o campo de Tickets: ")
    Dcriado = input("Digite a data que o chamado foi aberto (DD-MM-AAAA): ")
    primeira_hora = input("Digite a hora de abertura do chamado (HH:MM): ")
    Dfechado = input("Digite a data que o chamado foi encerrado (DD-MM-AAAA): ")
    segunda_hora = input("Digite a hora de fechamento do chamado (HH:MM): ")
    nomes_listas_suspensas = ["a origem do chamado", "a loja", "quem realizou o chamado", "o BKN da loja", "a importância do chamado", "a resolução do chamado"]
    
    # Preencher os campos do formulário
    preencher_campos(driver, email, number, tickets, Dcriado, primeira_hora, Dfechado, segunda_hora)

    # Chamada da função para selecionar as opções das listas suspensas
    selecionar_opcoes(driver, nomes_listas_suspensas)

    # Chamada da função para selecionar "Closed"
    selecionar_closed(driver)

    # Aguardar o carregamento do formulário
    time.sleep(1)

    # Localizar e clicar no botão "Próxima"
    botao_proxima = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Próxima']")))
    botao_proxima.click()

    time.sleep(1)   

    campo_conclusão = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[aria-labelledby='i5']")))
    if resolucao_escolhida == "Dongle desconectado / Headset":  # Verifica a resolução escolhida
        conclusão = "Dongle estava desconectado, foi necessario desabilitar e habilitar o dispositivo para corrigir."
    elif resolucao_escolhida == "Aparelho danificado / Headset":  # Verifica a resolução escolhida
        conclusão = "Headset esta danificado, foi necessário solicitar a aquisição de um novo."
    elif resolucao_escolhida == "Tela sobreposta ao Dashboard / Dashboard":  # Verifica a resolução escolhida
        conclusão = "Dashboard estava com uma tela sobreposta, foi necessário retirar a mesma"
    elif resolucao_escolhida == "Sistema offline / Sistema operacional Linux":  # Verifica a resolução escolhida
        conclusão = "Sistema estava offline, foi necessario desligar o disjuntor da Aiknow e liga-lo novamente para recobrar o sistema."
    elif resolucao_escolhida == "Veículo fantasma / Cameras":  # Verifica a resolução escolhida
        conclusão = "camera estava identificando veiculos fantasmas, foi realizado a abertura de um chamado interno para corrigir o problema."
    elif resolucao_escolhida == "Volume baixo colaborador / Sistema operacional Windows":  # Verifica a resolução escolhida
        conclusão = "Volume do microfone do colaborador estava baixo, foi aumentado o volume."
    elif resolucao_escolhida == "Volume baixo / Soundblaster":  # Verifica a resolução escolhida
        conclusão = "Volume do soundblaster desceu devido a variação de energia, foi aumentado o volume para corrigir o problema."
    elif resolucao_escolhida == "Volume baixo / Headset":  # Verifica a resolução escolhida
        conclusão = "Volume do headset estava baixo, colaborador foi orietnado a aumentar o volume via headset."
    elif resolucao_escolhida == "Volume baixo / Caixa de som":  # Verifica a resolução escolhida
        conclusão = "Caixa de som estava com o volume baixo, foi necessario aumentar o mesmo. Colaborador foi orientado a não mexer na caixa do sistema"   
    elif resolucao_escolhida == "Volume alto colaborador / Sistema operacional Windows":  # Verifica a resolução escolhida
        conclusão = "Volume do microfone do colaborador estava alto, foi ajustado para corrigir o problema"
    elif resolucao_escolhida == "Volume alto Cliente / Sistema operacional Windows":  # Verifica a resolução escolhida
        conclusão = "Volume do microfone do cliente estava alto, foi ajustado para corrigir o problema"
    elif resolucao_escolhida == "Volume alto / Caixa de som":  # Verifica a resolução escolhida
        conclusão = "Caixa de som estava com volume alto. devido a isso cliente nao escutava colaborador. Colaborador foi orientado a não mexer na caixa do sistema"
    elif resolucao_escolhida == "Travado / Soundblaster":  # Verifica a resolução escolhida
        conclusão = "Soundblaster havia travado, foi necessario reiniciar o sistema."
    elif resolucao_escolhida == "Sistema offline / Sistema operacional Windows":  # Verifica a resolução escolhida
        conclusão = "Sistema estava offline, foi necessario desligar o disjuntor da Aiknow e liga-lo novamente para recobrar o sistema."
    elif resolucao_escolhida == "Sistema foi reiniciado / Não houve problemas":  # Verifica a resolução escolhida
        conclusão = "Apos verificar em sistema nao foi encontrado problemas, devido a isso o mesmo foi reiniciado."
    elif resolucao_escolhida == "Retirada da posição original / Cameras":  # Verifica a resolução escolhida
        conclusão = "Camera teve seu angulo de visao alterado. Foi solicitado uma visita para corrigir o problema"
    elif resolucao_escolhida == "Queda de energia":  # Verifica a resolução escolhida
        conclusão = "Houve queda de energia na loja, devido a isso o sistema não iniciou corretamente"
    elif resolucao_escolhida == "Pareado com a cor incorreta / Headset":  # Verifica a resolução escolhida
        conclusão = "Headset havia sido pareado a cor incorreta, devido a isso foi necessário reemparelhar o dispositivo."   
    elif resolucao_escolhida == "Não está ligado / Headset":  # Verifica a resolução escolhida 20
        conclusão = "Headset estava desligado, colaborador foi orientado a como ligar o headset"
    elif resolucao_escolhida == "Travado / Headset":  # Verifica a resolução escolhida
        conclusão = "Headset se encontra travado, foi solicitado a aquisição de um novo"
    elif resolucao_escolhida == "Roi precisou ser redesenhado":  # Verifica a resolução escolhida
        conclusão = "Roi precisou ser redesenhado devido a não estar condizente com a area de atendimento"
    elif resolucao_escolhida == "Não reconhece veículo / Cameras":  # Verifica a resolução escolhida
        conclusão = "Camera esta tendo problemas na identificação de veiculos, devido a isso foi realizado a abertura de um chamado interno para corrigir o problema"
    elif resolucao_escolhida == "Ligado incorretamente / Headset":  # Verifica a resolução escolhida
        conclusão = "Headset havia sido ligado incorretamente, colaborador foi orientado a como ligar o dispositivo corretamente."
    elif resolucao_escolhida == "Loja sem internet":  # Verifica a resolução escolhida
        conclusão = "Loja estava sem internet durante um momento, após isso tudo foi normalizado"
    elif resolucao_escolhida == "Headset não está carregando / Headset":  # Verifica a resolução escolhida
        conclusão = "Headsets se encontravam sem bateria, colaborador foi orientado a colocar os dispositivos para carregar."
    elif resolucao_escolhida == "Headset Mudo / Headset":  # Verifica a resolução escolhida
        conclusão = "Headset estava mutado, foi desmutado para corrigir o problema"
    elif resolucao_escolhida == "Headset desemparelhado / Headset":  # Verifica a resolução escolhida
        conclusão = "Headset estava desemparelhado, colaborador foi orientado a como emparelhar o dispositivo."   
    elif resolucao_escolhida == "Headset com bateria baixa / Headset":  # Verifica a resolução escolhida
        conclusão = "Headset estava com bateria baixa, comlaborador foi orientado a colocar o dispositivo para carregar"
    elif resolucao_escolhida == "Esqueceu a senha de acesso ao Dashboard":  # Verifica a resolução escolhida 30
        conclusão = "Colaborador esqueceu a senha de acesso ao dashboard gerencial, o mesmo foi instruido a como recuperar a senha"
    elif resolucao_escolhida == "Dongle resetado / Headset":  # Verifica a resolução escolhida
        conclusão = "Dongle foi resetado devido a ter sido apertado. Colaborador foi instruido a não apertar o mesmo novamente"
    elif resolucao_escolhida == "Criar acesso ao Dashboard":  # Verifica a resolução escolhida
        conclusão = "Colaborador solicitou que fosse criado o acesso ao dashboard gerencial da loja"
    elif resolucao_escolhida == "Água na lente / Cameras":  # Verifica a resolução escolhida
        conclusão = "Camera estava com agua na lente, devido a isso estava com problemas para identificar veiculos."
    elif resolucao_escolhida == "Carregador perdido / Headset":  # Verifica a resolução escolhida
        conclusão = "Carregador dos headsets foram perdidos, foi solicitado novos a loja"
    elif resolucao_escolhida == "Dashboard não iniciou / Sistema operacional Linux":  # Verifica a resolução escolhida
        conclusão = "Dashboard não iniciou, devido a isso foi necessario reiniciar o sistema para corrigir o problema."
    elif resolucao_escolhida == "Veículo não toca música de entrada":  # Verifica a resolução escolhida
        conclusão = "Veiculo não tocou musica de entrada, o sistema foi reiniciado para corrigir o problema"
    elif resolucao_escolhida == "Sem conexão / Krisp":  # Verifica a resolução escolhida
        conclusão = "Krisp não estava transmitindo a voz do colaborador, foi necessário reiniciar ele para corrigir o problema"
    else:
        conclusão = input("Digite a conclusão do chamado: ")
    campo_conclusão.clear()
    campo_conclusão.send_keys(conclusão)



    # Localizar e clicar no botão de envio
    botao_enviar = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Enviar']")))
    botao_enviar.click()

    #Pede o número do ticket ao usuário
    numero_ticket = get_next_number(sh)
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal do tkinter
    messagebox.showinfo("Chamado cadastrado com sucesso", f"Número do Ticket: {numero_ticket}\nResolução: {resolucao_escolhida}")
    reportar = messagebox.askyesno("Reportar Ticket", "Deseja reportar este ticket ao bot do Telegram?")
    if reportar:
        # Inicializa o bot do Telegram
        global bot
        bot = telegram.Bot(token='Token do Bot')

        # Cria uma janela para a descrição do problema
        report_window = tk.Toplevel()
        report_window.title("Descrição do Problema")

        # Label para o número do ticket
        ticket_label = tk.Label(report_window, text=f"Número do Ticket: {numero_ticket}")
        ticket_label.pack()

        # Campo de texto para a descrição do problema
        descricao_text = tk.Text(report_window, height=5)
        descricao_text.pack()

        # Botão para enviar o relatório
        def enviar_relatorio():
            global opcoes_selecionadas
            descricao = descricao_text.get("1.0", "end-1c")
            mensagem = f"Ticket: {numero_ticket}\nUsuário: {User}\nLoja: {bkn_escolhido}\nDescrição: {descricao}" # Inclui o 'bkn_escolhido' na mensagem

            async def enviar_mensagem_telegram():
                try:
                    await bot.send_message(chat_id='ChatID', text=mensagem)
                except Exception as e:
                    print(f"Erro ao enviar mensagem: {e}")

         # Executa a função assíncrona usando asyncio.run()
            try:
                asyncio.run(enviar_mensagem_telegram()) 
            except Exception as e:
                print(f"Erro: {e}")

            # Fecha a janela
            report_window.destroy()

        enviar_button = tk.Button(report_window, text="Enviar Relatório", command=enviar_relatorio)
        enviar_button.pack()  # Posiciona o botão na janela

    return input("Deseja realizar o procedimento novamente? (S/N): ").strip().upper() == 'S'

# Executar o procedimento enquanto o usuário quiser
while preencher_formulario():
    pass

driver.quit()
