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
    from selenium.common.exceptions import TimeoutException, ElementNotInteractableException
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


    # --- INÍCIO DA EXECUÇÃO DO SCRIPT ---
    print("Iniciando automação...")
    navegador.get("https://kronos.servicenet.inf.br/")
    WebDriverWait(navegador, 30).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    actions = ActionChains(navegador)
    time.sleep(3) # Give browser a moment to fully render

    # --- Lidar com Pop-ups de Cookies/Termos (MELHORADO) ---
    # Tenta clicar no botão de aceitar cookies, se ele aparecer
    try:
        print("Verificando por botão de aceitar cookies...")
        clicar_elemento_padrao(navegador, By.ID, "acceptCookiesButton", wait_time=5)
        print("Cookies aceitos.")
        time.sleep(1)
    except Exception as e:
        print(f"Botão de aceitar cookies não encontrado ou erro ao clicar: {e}. Prosseguindo sem clicar.")

    # --- Login ---
    print("\n--- Iniciando processo de Login ---")
    inserir_texto_padrao(navegador, By.ID, "username", "gabriele.medina")
    # Inserir senha
    inserir_texto_padrao(navegador, By.ID, "password", "Gimave@2025") # ATENÇÃO: Considere usar um método mais seguro para senhas!
    time.sleep(1)

    # Clicar no botão Acessar
    clicar_elemento_padrao(navegador, By.ID, "login-btn")
    print("Botão 'Acessar' clicado. Aguardando carregamento...")
    time.sleep(5) # Esperar a página de login processar

    #Clicar em Conciliação
    clicar_elemento_padrao(navegador, By.LINK_TEXT, "Conciliação")
    time.sleep(2) # Esperar a página de login processar

    #Clicar em Rede
    clicar_elemento_padrao(navegador, By.XPATH, "//a[text()='Rede' and ./span[@class='fa fa-chevron-down']]")
    time.sleep(2) # Esperar a página de login processar

    #Clicar Lançamento Manual
    clicar_elemento_padrao(navegador, By.LINK_TEXT, "Lançamento Manual")
    time.sleep(5) # Esperar a página de login processar


    # --- Processar cada linha da tabela (FUNÇÃO REFATORADA) ---
    def processar_linha(driver_instance, row_data): # Pass driver_instance as argument
        try:
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
            # Tenta obter o valor da planilha. Se não for numérico, define como 0.0.
            # O método .get evita erros se a coluna não existir.
            valor_numerico = pd.to_numeric(row_data.get("Valor:"), errors="coerce")
            if pd.isna(valor_numerico):
                valor_numerico = 0.0
            # Converte o valor numérico para uma string com 2 casas decimais e substitui o ponto por vírgula.
            # Exemplo: 500.0 vira '500.00' e depois '500,00'.
            valor_formatado = f"{valor_numerico:.2f}".replace('.', ',')

            Loja = row_data.get("Loja", "").strip()
            Estabelecimento = row_data.get("Estabelecimento", "").strip()
            Historico = 'Recarga do cliente {} Das Vendas do dia {} Até o dia {}'.format(Loja, data_inicio_formatada, data_final_formatada)
            Doc = ('Vendas {}. a {}').format( data_inicio_formatada, data_final_formatada)

            #Clicar Solicitar
            clicar_elemento_padrao(navegador, By.XPATH, "/html/body/div[3]/div/div[5]/ul/div/a")
            time.sleep(2) # Esperar a página de login processar

            #Clicar Caixa
            clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="select2-id_tipo-container"]')
            time.sleep(2) # Esperar a página de login processar

            #Clicar em Recarga
            try:
                # Usamos WebDriverWait para esperar que o elemento 'Recarga' esteja clicável.
                # Ele procura um elemento <li> que tem a classe 'select2-results__option' E
                # cujo texto visível é 'Recarga'.
                element_recarrega = WebDriverWait(navegador, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'select2-results__option') and text()='Recarga']"))
                )
                element_recarrega.click()
                print("Elemento 'Recarga' clicado com sucesso.")
                time.sleep(2) # Um pequeno delay após o clique
            except TimeoutException:
                print("ERRO: Elemento 'Recarga' não encontrado/clicável em 10s.")
                # Adicione aqui um tratamento de erro ou um exit() se for crítico
                # exit()
            except Exception as e:
                print(f"ERRO inesperado ao clicar em 'Recarga': {e}")
                # exit()

            #Clicar Estabelecimento
            clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="select2-id_id_estabelecimento-container"]')           
            time.sleep(2) # Esperar 
            actions.send_keys(Estabelecimento).perform()
            time.sleep(1)
            actions.send_keys(Keys.ENTER).perform()
            time.sleep(2) # Esperar a página de login processar

            #Clicar EM NÚMERO DE LOJA
            clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="select2-id_cc_loja-container"]')
            time.sleep(2) # Esperar 
            actions.send_keys(Loja).perform()
            time.sleep(1)
            actions.send_keys(Keys.ENTER).perform()
            time.sleep(2) # Esperar a página de login processar

            #ESCERVER EM NÚMERO DE DOCUMENTO
            inserir_texto_padrao(navegador, By.XPATH, '//*[@id="id_num_documento"]' , Doc)
            time.sleep(2) # Esperar a página de login processar

            # ESCRIVER VALOR
            inserir_texto_padrao(navegador, By.XPATH, '//*[@id="id_valor"]' , valor_formatado)
            time.sleep(2)

            #ESCREVER OBSEVAÇÃO
            inserir_texto_padrao(navegador, By.XPATH, '//*[@id="id_observacao"]' , Historico)
            time.sleep(2)

            #CLICAR EM SALVAR
            clicar_elemento_padrao(navegador, By.XPATH, '//*[@id="btn_salvar_lancamento_manual"]')
            time.sleep(5)

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