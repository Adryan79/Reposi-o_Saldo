try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from datetime import timedelta
    import time
    import numpy as np
    import cv2
    import os
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains
    import pandas as pd 
    from selenium.common.exceptions import TimeoutException, ElementNotInteractableException, NoAlertPresentException
    from datetime import datetime, timedelta
    # No topo do seu script (após os imports)
    import sys
    import pandas as pd
    # O caminho do arquivo Excel será o primeiro argumento passado pelo servidor
    excel_file_path = sys.argv[1]

    # Use esta variável para ler o arquivo Excel correto
    print(f"Lendo dados do arquivo: {excel_file_path}")
    tabela_produtos_df = pd.read_excel(excel_file_path)

    # Configurar opções do Chrome
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    # Desativar o gerenciador de senhas e a oferta de salvar senhas
    prefs = {
        "credentials_enable_service": False,  # Desativa o serviço de credenciais
        "profile.password_manager_enabled": False  # Desativa o gerenciador de senhas
    }
    options.add_experimental_option("prefs", prefs)

    # Inicializar navegador com as opções
    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico, options=options)

    # --- FUNÇÕES PADRONIZADAS (ADICIONADAS/AJUSTADAS) ---
    def inserir_texto_padrao(driver, locator_type, locator_value, text_to_insert, wait_time=10):
        """
        Função padronizada para inserir texto em um elemento web de forma segura.
        Melhorada para lidar com elementos que podem não estar clicáveis inicialmente.
        """
        try:
            print(f"Tentando inserir texto '{text_to_insert}' em elemento: {locator_type}='{locator_value}'...")
            element = WebDriverWait(driver, wait_time).until(
                EC.visibility_of_element_located((locator_type, locator_value))
            )
            # Tenta clicar no elemento para garantir o foco, caso send_keys falhe direto
            try:
                element.click()
            except ElementNotInteractableException:
                driver.execute_script("arguments[0].click();", element)
                print(f"Element {locator_value} clicked via JavaScript.")

            element.clear()
            element.send_keys(str(text_to_insert))
            print(f"Texto '{text_to_insert}' inserido com sucesso em {locator_value}.")
        except TimeoutException:
            print(f"ERRO: Elemento {locator_type}='{locator_value}' não encontrado/visível em {wait_time}s para inserir texto.")
            raise
        except Exception as e:
            print(f"ERRO inesperado ao inserir texto em {locator_type}='{locator_value}': {e}")
            raise

    def clicar_elemento_padrao(driver, locator_type, locator_value, wait_time=10):
        """
        Função padronizada para clicar em um elemento web de forma segura.
        """
        try:
            print(f"Tentando clicar em elemento: {locator_type}='{locator_value}'...")
            element = WebDriverWait(driver, wait_time).until(
                EC.element_to_be_clickable((locator_type, locator_value))
            )
            element.click()
            print(f"Elemento {locator_value} clicado com sucesso.")
        except TimeoutException:
            print(f"ERRO: Elemento {locator_type}='{locator_value}' não encontrado/clicável em {wait_time}s.")
            raise
        except Exception as e:
            print(f"ERRO inesperado ao clicar em {locator_type}='{locator_value}': {e}")
            raise

    def apagar_Campo(driver, elemento_id, tempo_espera=30):
        """
        Esperar até que o elemento identificado pelo ID seja clicável, clica, limpa o campo e define "0,00".
        Considera usar a função padronizada de inserção para consistência.
        """
        try:
            print(f"Apagando e definindo '0,00' no campo {elemento_id}...")
            # Reusing inserir_texto_padrao for consistency and robustness
            inserir_texto_padrao(driver, By.ID, elemento_id, "0,00", tempo_espera)
            print(f'Valor "0,00" definido com sucesso no elemento {elemento_id}.')
        except Exception as e:
            print(f"Erro ao apagar ou definir o valor do elemento {elemento_id}: {e}")
            raise # Propagate the error

    def digitar_entrada_com_TAB(driver, texto, tab_count=0):
        """
        Digita texto no elemento ativo e simula 'tab' key presses.
        Considerar substituir por 'inserir_texto_padrao' e 'send_keys(Keys.TAB)' no futuro.
        """
        print(f"Digitando '{texto}' e pressionando TAB {tab_count} vezes...")
        active_element = driver.switch_to.active_element
        active_element.send_keys(texto)
        for _ in range(tab_count):
            active_element.send_keys(Keys.TAB)
        time.sleep(1)

    def esperar_e_clicar_simples(driver, elemento_id, tempo_espera=30):
        """
        Espera até que o elemento identificado pelo ID seja clicável e, então, realiza um duplo clique.
        """
        try:
            print(f"Tentando duplo clique no elemento ID: {elemento_id}...")
            elemento = WebDriverWait(driver, tempo_espera).until(
                EC.element_to_be_clickable((By.ID, elemento_id))
            )
            action = ActionChains(driver)
            action.double_click(elemento).perform()
            print(f'Duplo clique realizado com sucesso no elemento {elemento_id}.')
        except Exception as e:
            print(f"Erro ao realizar o duplo clique no elemento {elemento_id}: {e}")
            raise # Propagate the error

    def digitar_entrada(driver, texto, tab_count=0): # Redundant with inserir_texto_padrao or digitar_entrada_com_TAB
        """
        Digita texto no elemento ativo. Considerar substituir por 'inserir_texto_padrao'.
        """
        print(f"Digitando '{texto}' no elemento ativo...")
        driver.switch_to.active_element.send_keys(texto)
        time.sleep(1)

    def esperar_imagem_aparecer(driver, imagem_alvo, timeout=80, threshold=0.7):
        """
        Espera até que uma imagem apareça na tela do navegador.
        """
        print(f"Esperando imagem '{imagem_alvo}' aparecer na tela (timeout: {timeout}s)...")
        if not os.path.exists(imagem_alvo):
            print(f"ERRO: Imagem alvo não encontrada no caminho: {imagem_alvo}")
            return False

        template = cv2.imread(imagem_alvo, cv2.IMREAD_COLOR)
        if template is None:
            print(f"ERRO: Não foi possível carregar a imagem do template: {imagem_alvo}")
            return False
        template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

        start_time = time.time()
        intervalo = 2

        while time.time() - start_time < timeout:
            screenshot = driver.get_screenshot_as_png()
            screen_array = np.frombuffer(screenshot, dtype=np.uint8)
            screen_image = cv2.imdecode(screen_array, cv2.IMREAD_COLOR)
            screen_gray = cv2.cvtColor(screen_image, cv2.COLOR_BGR2GRAY)

            result = cv2.matchTemplate(screen_gray, template_gray, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, _ = cv2.minMaxLoc(result)

            if max_val >= threshold:
                print(f"Imagem '{imagem_alvo}' encontrada com precisão {max_val:.2f}.")
                return True
            else:
                print(f"Imagem '{imagem_alvo}' ainda não visível (precisão: {max_val:.2f}), aguardando...")
                time.sleep(intervalo)
        print(f"Tempo limite atingido ({timeout}s). Imagem '{imagem_alvo}' não encontrada.")
        return False

    def Clique_Ousado(driver, imagem_alvo, timeout=80, precision=0.90):
        """
        Tenta detectar e clicar em uma imagem na tela do navegador com alta precisão.
        """
        print(f"Tentando clicar na imagem '{imagem_alvo}' com precisão >= {precision:.2f} (timeout: {timeout}s)...")
        if not os.path.exists(imagem_alvo):
            print(f"ERRO: Imagem alvo não encontrada: {imagem_alvo}")
            return

        template = cv2.imread(imagem_alvo, cv2.IMREAD_COLOR)
        if template is None:
            print(f"ERRO: Não foi possível carregar a imagem do template: {imagem_alvo}")
            return
        template_h, template_w = template.shape[:2]
        template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

        start_time = time.time()
        intervalo = 2

        while time.time() - start_time < timeout:
            screenshot = driver.get_screenshot_as_png()
            screen_array = np.frombuffer(screenshot, dtype=np.uint8)
            screen_image = cv2.imdecode(screen_array, cv2.IMREAD_COLOR)
            screen_gray = cv2.cvtColor(screen_image, cv2.COLOR_BGR2GRAY)

            result = cv2.matchTemplate(screen_gray, template_gray, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)

            print(f"Precisão da imagem encontrada: {max_val:.2f}")

            if max_val >= precision:
                click_x = max_loc[0] + template_w // 2
                click_y = max_loc[1] + template_h // 2

                # Optional: Draw rectangle and save screenshot for debugging
                cv2.rectangle(screen_image, (max_loc[0], max_loc[1]),
                                (max_loc[0] + template_w, max_loc[1] + template_h),
                                (0, 255, 0), 2)
                cv2.imwrite("clicado.png", screen_image)
                print("Print da área clicada salvo como 'clicado.png'.")

                # Scroll to element before clicking if necessary (more reliable)
                driver.execute_script(f"window.scrollTo(0, {max_loc[1] - 100});") # Scroll a bit above the element
                time.sleep(0.5) # Give it time to scroll

                # Use ActionChains for a more reliable click at coordinates relative to viewport
                ActionChains(driver).move_to_element_with_offset(
                    driver.find_element(By.TAG_NAME, 'body'), click_x, click_y
                ).click().perform()
                print(f"Clicado na imagem '{imagem_alvo}' em X: {click_x}, Y: {click_y}.")
                return
            else:
                print(f"Imagem '{imagem_alvo}' não encontrada com {precision:.2f} de precisão, aguardando...")
                time.sleep(intervalo)

        print(f"ERRO: Imagem '{imagem_alvo}' não encontrada após tempo limite ({timeout}s).")

    def Clique_Ousado_Duas_Vezes(driver, imagem_alvo, timeout=80, precision=0.90):
        """
        Tenta detectar e realizar um duplo clique em uma imagem na tela do navegador.
        """
        print(f"Tentando duplo clique na imagem '{imagem_alvo}' (timeout: {timeout}s)...")
        if not os.path.exists(imagem_alvo):
            print(f"ERRO: Imagem alvo não encontrada: {imagem_alvo}")
            return

        template = cv2.imread(imagem_alvo, cv2.IMREAD_COLOR)
        if template is None:
            print(f"ERRO: Não foi possível carregar a imagem do template: {imagem_alvo}")
            return
        template_h, template_w = template.shape[:2]
        template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

        start_time = time.time()
        intervalo = 2

        while time.time() - start_time < timeout:
            screenshot = driver.get_screenshot_as_png()
            screen_array = np.frombuffer(screenshot, dtype=np.uint8)
            screen_image = cv2.imdecode(screen_array, cv2.IMREAD_COLOR)
            screen_gray = cv2.cvtColor(screen_image, cv2.COLOR_BGR2GRAY)

            result = cv2.matchTemplate(screen_gray, template_gray, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)

            print(f"Precisão da imagem encontrada: {max_val:.2f}")

            if max_val >= precision:
                click_x = max_loc[0] + template_w // 2
                click_y = max_loc[1] + template_h // 2

                cv2.rectangle(screen_image, (max_loc[0], max_loc[1]),
                                (max_loc[0] + template_w, max_loc[1] + template_h),
                                (0, 255, 0), 2)
                cv2.imwrite("clicado.png", screen_image)
                print("Print da área clicada salvo como 'clicado.png'.")

                driver.execute_script(f"window.scrollTo(0, {max_loc[1] - 100});")
                time.sleep(0.5)

                # Find the element at the point and double click it
                element = driver.execute_script(f"return document.elementFromPoint({click_x}, {click_y});")
                if element:
                    action = ActionChains(driver)
                    action.move_to_element(element).double_click().perform()
                    print(f"Duplo clique realizado na imagem '{imagem_alvo}'!")
                    return
                else:
                    print(f"Elemento não encontrado para duplo clique na imagem '{imagem_alvo}'.")
                    break # Exit loop if element cannot be found/clicked

            else:
                print(f"Imagem '{imagem_alvo}' não encontrada com {precision:.2f} de precisão, aguardando...")
                time.sleep(intervalo)
        print(f"ERRO: Imagem '{imagem_alvo}' não encontrada após tempo limite ({timeout}s).")

    def fatura(driver, imagem_alvo, tempo_maximo=3, confidence=0.7):
        """
        Verifica se o pedido está faturado em segundo plano sem interações diretas.
        """
        print(f"Verificando imagem de faturamento '{imagem_alvo}' (tempo máximo: {tempo_maximo}s)...")
        if not os.path.exists(imagem_alvo):
            print(f"ERRO: Imagem alvo não encontrada: {imagem_alvo}")
            return False

        WebDriverWait(driver, tempo_maximo).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

        inicio = time.time()
        intervalo = 0.5
        template = cv2.imread(imagem_alvo, cv2.IMREAD_COLOR)

        if template is None:
            print(f"ERRO: Erro ao carregar a imagem de referência: {imagem_alvo}")
            return False

        template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

        while time.time() - inicio < tempo_maximo:
            screenshot = driver.get_screenshot_as_png()
            screen_array = np.frombuffer(screenshot, dtype=np.uint8)
            screen_image = cv2.imdecode(screen_array, cv2.IMREAD_COLOR)
            screen_gray = cv2.cvtColor(screen_image, cv2.COLOR_BGR2GRAY)

            result = cv2.matchTemplate(screen_gray, template_gray, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, _ = cv2.minMaxLoc(result)

            print(f"Precisão da imagem detectada: {max_val:.2f}")

            if max_val >= confidence:
                print(f"Imagem '{imagem_alvo}' encontrada!")
                return True

            time.sleep(intervalo)

        print(f"Imagem '{imagem_alvo}' não encontrada dentro do tempo máximo.")
        return False

    def esperar_e_clicar(driver, locator, tempo_espera=30, cliques=1): # Changed 'elemento_id' to 'locator' for flexibility
        """
        Espera até que o elemento (identificado por ID ou XPath) esteja presente e realiza múltiplos cliques.
        Preferir usar 'clicar_elemento_padrao' se o número de cliques for 1.
        """
        print(f"Esperando e clicando {cliques} vez(es) em: {locator}...")
        try:
            # Determine locator type based on typical ID or XPath format
            if locator.startswith('/') or locator.startswith('.'): # Simple check for XPath
                by_type = By.XPATH
            else: # Assume ID by default
                by_type = By.ID

            elemento = WebDriverWait(driver, tempo_espera).until(
                EC.element_to_be_clickable((by_type, locator))
            )

            action = ActionChains(driver)
            if cliques == 1:
                elemento.click() # Direct click is often faster for single clicks
            elif cliques == 2:
                action.double_click(elemento).perform()
            else:
                for _ in range(cliques):
                    action.click(elemento)
                action.perform()

            print(f'{cliques} clique(s) realizado(s) com sucesso no elemento {locator}.')
        except Exception as e:
            print(f"Erro ao clicar no elemento '{locator}': {e}")
            raise # Propagate the error
    
    def aceitar_alerta(driver, tempo_espera=10):
        """
        Espera por um alerta e o aceita (clica em OK).
        """
        try:
            WebDriverWait(driver, tempo_espera).until(EC.alert_is_present())
            alerta = driver.switch_to.alert
            print(f"Alerta encontrado com texto: '{alerta.text}'")
            alerta.accept()
            print("Alerta aceito (OK) com sucesso.")
            time.sleep(2)  # Pausa para o alerta fechar e a página se atualizar
        except TimeoutException:
            print("Nenhum alerta presente no tempo esperado.")
        except NoAlertPresentException:
            print("Nenhum alerta presente.")
        except Exception as e:
            print(f"Ocorreu um erro ao lidar com o alerta: {e}")

    # --- INÍCIO DA EXECUÇÃO DO SCRIPT ---
    print("Iniciando automação...")
    navegador.get("http://netsales.manaus.prodatamobility.com.br")
    WebDriverWait(navegador, 30).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    actions = ActionChains(navegador)
    time.sleep(3) # Give browser a moment to fully render

    # --- Login ---
    print("\n--- Iniciando processo de Login ---")
    inserir_texto_padrao(navegador, By.XPATH, '//*[@id="ctl00_cphMainUnAutMaster_txtUser"]', "gimave.adm2")
    # Inserir senha
    inserir_texto_padrao(navegador, By.XPATH, '//*[@id="ctl00_cphMainUnAutMaster_txtPassword"]', "123456") # ATENÇÃO: Considere usar um método mais seguro para senhas!
    time.sleep(1)

    # Clicar no botão Acessar
    clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="ctl00_cphMainUnAutMaster_btnOk"]')
    print("Botão 'Acessar' clicado. Aguardando carregamento...")
    time.sleep(5) # Esperar a página de login processar

    #Clicar em Lançamento Manual
    clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="ctl00_mnuMain_ctl17"]/td/a')
    time.sleep(2) # Esperar a página de login processar

    #Clicar em Rede de Vendas
    clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ddlNetSale"]')
    time.sleep(2) # Esperar a página de login processar

    #Clicar VT
    actions.send_keys('VT').perform()
    time.sleep(1)
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(2) # Esperar a página de login processar

    # Clicar em credenciado Na Lupa
    clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="imgSearchCred"]')
    print("Elemento //*[@id='imgSearchCred'] clicado com sucesso.")
    time.sleep(5) # Esperar a página de login processar

    # --- LIDANDO COM A NOVA JANELA POP-UP ---
    print("\n--- Lidando com a nova janela pop-up ---")

    # Guarda o identificador (handle) da janela original antes de qualquer interação com o pop-up
    janela_original = navegador.current_window_handle
    print(f"Identificador da janela original: {janela_original}")

    try:
        # Espera até que a nova janela seja detectada
        WebDriverWait(navegador, 30).until(EC.number_of_windows_to_be(2))
        print("Nova janela detectada.")

        # Obtém a lista de todos os identificadores de janelas
        janelas_abertas = navegador.window_handles
        nova_janela_handle = None

        # Itera sobre os handles para encontrar a nova janela
        for handle in janelas_abertas:
            if handle != janela_original:
                nova_janela_handle = handle
                break

        if nova_janela_handle:
            # Muda o foco do Selenium para a nova janela
            navegador.switch_to.window(nova_janela_handle)
            print(f"Foco do navegador mudado para a nova janela: {navegador.current_url}")
            time.sleep(5) # Tempo para a página pop-up carregar

            # --- Suas interações com a nova janela entram aqui ---
            # A partir deste ponto, o Selenium está controlando o pop-up.

            clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnSearch"]')
            time.sleep(10)
            clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_grdProvider_ctl02_imgSelect"]')
            time.sleep(10)
            # --- Fim das interações com a nova janela ---

    except TimeoutException:
        print("ERRO: Nova janela pop-up não apareceu no tempo esperado. Prosseguindo na janela original.")
    except Exception as e:
        # Captura qualquer erro que possa ocorrer durante a interação com o pop-up
        print(f"Ocorreu um erro durante a interação com o pop-up: {e}")
    finally:
        # Este bloco é EXTREMAMENTE IMPORTANTE. Ele sempre será executado,
        # mesmo se ocorrer um erro nos blocos 'try' ou 'except'.
        try:
            # Fecha a nova janela pop-up
            if 'nova_janela_handle' in locals() and nova_janela_handle:
                navegador.close()
                print("Nova janela fechada com sucesso.")
        except:
            # Ignora se a janela já estiver fechada
            pass

        # Retorna o foco do Selenium para a janela original
        navegador.switch_to.window(janela_original)
        print("Foco retornado para a janela principal.")
        time.sleep(2) # Pausa para garantir que o foco seja restabelecido
    
     # 1. Obter a data de hoje
    hoje = datetime.now()

    # 2. Calcular as datas Inicio_Venda e Final_Venda
    Inicio_Venda = hoje - timedelta(days=7)
    Final_Venda = hoje - timedelta(days=1)

    # 3. Formatar as datas para o formato "dd/mm/yyyy"
    data_inicio_formatada = Inicio_Venda.strftime("%d/%m/%Y")
    data_final_formatada = Final_Venda.strftime("%d/%m/%Y")

    # 4. Imprimir os resultados (opcional)
    print(f"Data de Início da Venda: {data_inicio_formatada}")
    print(f"Data Final da Venda: {data_final_formatada}")
    # Extraindo dados da linha com segurança
    # A partir daqui, o foco está garantido na janela principal
    Obs = 'Vendas {} a {}'.format(data_inicio_formatada, data_final_formatada)
    # INSERIR OBSERVAÇÃO
    inserir_texto_padrao(navegador, By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtComment"]', Obs)
    time.sleep(2)

    # --- Processar cada linha da tabela (FUNÇÃO REFATORADA) ---
    def processar_linha(driver_instance, row_data): # Pass driver_instance as argument
        try:
            
            # Tente obter o valor e o nome do DataFrame com segurança
            valor_raw = row_data.get("Valor:")
            Nome = row_data.get("Loja", "").strip()

            # Verifique se o valor é válido e converta-o para o formato esperado
            if pd.isna(valor_raw):
                print("AVISO: Valor para 'Valor:' não encontrado ou inválido.")
                return # Pula esta linha
            
            # Garanta que o valor é um número (float) e formate-o com duas casas decimais e vírgula
            valor_numerico = float(valor_raw)
            # Use f-string para formatar com 2 casas decimais e depois substitua o ponto por vírgula
            Valor = str(f"{valor_numerico:.2f}").replace(".", ",")
            
            #INSERIR VALOR
            inserir_texto_padrao(navegador, By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtValue"]', Valor)
            time.sleep(3) # Esperar a página de login processar

            #INSERIR CLIENTE
            # Guarda o handle da janela principal
            janela_principal = navegador.current_window_handle
            
            clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="imgSearchLoja"]')
            time.sleep(5)

            # AQUI VAI ABRIR ESSA JANELA 
            # http://netsales.manaus.prodatamobility.com.br/Aut/ProviderListStore.aspx?pMASTERID=62&pProviderNetSaleId=7&pProviderNetSaleDesc=VTPASSAFACIL
            janela_popup = None  # Inicializa a variável fora do try
            try:
                # Espera a nova janela pop-up
                WebDriverWait(navegador, 10).until(EC.number_of_windows_to_be(2))
                janelas = navegador.window_handles
                for janela in janelas:
                    if janela != janela_principal:
                        janela_popup = janela
                        break

                if janela_popup:
                    # Mudar o foco para a janela pop-up
                    navegador.switch_to.window(janela_popup)
                    print(f"Foco mudado para a janela pop-up: {navegador.current_url}")
                    
                    # Inserir o nome do cliente no campo de texto da pop-up
                    inserir_texto_padrao(navegador, By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtDescription"]', Nome)
                    time.sleep(2)
                    
                    # Clicar no botão de busca da pop-up (lupa)
                    clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnSearch"]')
                    time.sleep(2)
                    
                    # Clicar para selecionar o cliente na grade de resultados
                    clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_grdProvider_ctl02_imgSelect"]')
                    time.sleep(5)
                    
                else:
                    print("ERRO: Janela pop-up não encontrada após o clique.")
                    # Você pode levantar uma exceção aqui ou continuar, dependendo da sua necessidade
                    raise Exception("Pop-up não encontrado")

            except TimeoutException:
                print("ERRO: Janela pop-up do cliente não apareceu no tempo esperado.")
                # O script vai para o finally para tentar fechar a janela, mas não irá mais executar as linhas abaixo
                raise # Mantém o erro para que a automação seja interrompida
            except Exception as e:
                print(f"Ocorreu um erro na janela pop-up do cliente: {e}")
                # O script vai para o finally para tentar fechar a janela
                raise # Mantém o erro para que a automação seja interrompida
            finally:
                # Fechar a janela pop-up e retornar para a principal, se o pop-up foi encontrado
                if janela_popup:
                    try:
                        navegador.close()
                        print("Janela pop-up fechada com sucesso.")
                    except:
                        print("Pop-up já estava fechado.")

                    # Verifica se a janela principal ainda está disponível
                    if janela_principal in navegador.window_handles:
                        navegador.switch_to.window(janela_principal)
                        print("Foco retornado para a janela principal.")
                        time.sleep(2) # Pausa para garantir que o foco seja restabelecido
                    else:
                        print("ERRO: Janela principal foi fechada. Não é possível continuar a automação.")
                        # Levanta um erro fatal se a janela principal não existir mais
                        raise Exception("Janela principal não encontrada.")
                    
            #Clicar em Tipo:
            clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ddlType"]')
            time.sleep(2)

            #Esolher Crédito
            clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ddlType"]/option[2]')
            time.sleep(2)
            
            # CLICAR Em Efetuar Lançamento
            clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnConfirm"]')
            time.sleep(3)

            # CLICAR Em OK - SOLUÇÃO PARA O SEGUNDO ALERTA
            aceitar_alerta(navegador, tempo_espera=10)
            print('Saldo Lançado do Cliente {} No Valor {} Com Sucesso'.format(Nome, Valor))


        # Fim do Loop
        except Exception as e:
            print(f"ERRO ao processar linha: {e}")
            # Decide if this error should stop the loop or just log and continue
            # raise # Uncomment to stop the script on any row processing error

    # Loop para processar cada linha da tabela
    if not tabela_produtos_df.empty: # Check if the DataFrame is empty
        for index, row in tabela_produtos_df.iterrows():
            print(f"\n--- Iniciando processamento da linha {index} do Excel ---")
            processar_linha(navegador, row) # Pass the row as a Series or dictionary
    else:
        print("Nenhuma linha para processar na tabela do Excel.")


    print('\nProcesso de automação concluído com sucesso!')


except Exception as e:
    erro_contador = globals().get('erro_contador', 0) + 1 # Safely increment if it exists, otherwise start at 1
    max_erros = globals().get('max_erros', 5) # Safely get max_erros, default to 5
    print(f'ERRO CRÍTICO na automação ({erro_contador}/{max_erros}): {e}')
    if erro_contador >= max_erros:
        print('Número máximo de erros consecutivos atingido. Encerrando automação.')
    else:
        print('Tentando novamente... (Esta lógica de "tentar novamente" precisa de um loop externo para ser eficaz).')
finally:
    # Garante que o navegador seja fechado, mesmo em caso de erro
    if 'navegador' in locals() and navegador.service.process is not None:
        print("Fechando navegador...")
        navegador.quit()
    print("Fim do script.")