import smtplib
from datetime import datetime
import pyautogui
import sys
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
import time
from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
# Interface
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

from PIL import Image, ImageTk
import threading

# ========== SPLASH SCREEN ========== 
def mostrar_splash():
    splash = tk.Tk()
    splash.overrideredirect(True)
    splash.geometry("400x300+500+300")  # Centraliza
    splash.configure(bg="black")

    try:
        imagem = Image.open("logo.jpg")
        imagem = imagem.resize((100, 100))
        imagem_tk = ImageTk.PhotoImage(imagem)
        label_img = tk.Label(splash, image=imagem_tk, bg="black")
        label_img.image = imagem_tk
        label_img.pack(pady=(30, 10))
    except Exception as e:
        print("Erro ao carregar imagem:", e)

    tk.Label(splash, text="DH Scripts", font=(
        "Helvetica", 24, "bold"), fg="white", bg="black").pack()
    tk.Label(splash, text="Iniciando...", font=("Helvetica", 12),
             fg="gray", bg="black").pack(pady=(10, 0))
    # Contato no canto inferior direito
    contato_label = tk.Label(splash, text="Cel. (82) 99121-7317",
                             font=("Helvetica", 10),
                             fg="white", bg="black")
    contato_label.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)

    def fechar():
        splash.destroy()

    splash.after(5000, fechar)
    splash.mainloop()

mostrar_splash()
# ========== CONFIGURAÇÕES DE CHAVE DE ACESSO ==========
ARQUIVO_CREDENCIAIS = ''
ID_PLANILHA = ''
NOME_ABA = 'Página1'
ARQUIVO_CHAVE_SALVA = 'chave_acesso.txt'

# ========== FUNÇÕES DE CHAVE DE ACESSO ==========
def carregar_credenciais():
    if not os.path.exists(ARQUIVO_CREDENCIAIS):
        pyautogui.alert(
            f"Arquivo de credenciais '{ARQUIVO_CREDENCIAIS}' não encontrado.")
        return None
    try:
        creds = service_account.Credentials.from_service_account_file(
            ARQUIVO_CREDENCIAIS)
        service = build('sheets', 'v4', credentials=creds)
        return service
    except Exception as e:
        pyautogui.alert(f"Erro ao carregar as credenciais do Google:\n{e}")
        return None

def verificar_chave(chave_usuario):
    service = carregar_credenciais()
    if not service:
        return False
    try:
        sheet = service.spreadsheets()
        resultado = sheet.values().get(spreadsheetId=ID_PLANILHA,
                                       range=f"{NOME_ABA}!A2:C").execute()
        valores = resultado.get('values', [])

        for linha in valores:
            if len(linha) >= 3:
                chave, status, data_expiracao = linha[0].strip(), linha[1].strip().lower(), linha[2].strip()
                if chave_usuario.strip() == chave and status == 'sim':
                    try:
                        data_exp = datetime.strptime(data_expiracao, "%Y-%m-%d").date()
                        hoje = datetime.now().date()
                        if hoje <= data_exp:
                            return True
                        else:
                            pyautogui.alert("❌ Chave expirada.")
                            return False
                    except ValueError:
                        pyautogui.alert(f"⚠️ Data inválida para a chave '{chave}'.")
                        return False
        return False
    except Exception as e:
        pyautogui.alert(f"Erro ao acessar a planilha:\n{e}")
        return False

def salvar_chave_local(chave):
    with open(ARQUIVO_CHAVE_SALVA, 'w') as f:
        f.write(chave)

# ========== VERIFICAÇÃO DA CHAVE ========== 
chave_digitada = pyautogui.prompt("🔐 Digite sua chave de acesso:")

if not chave_digitada:
    pyautogui.alert("❌ Nenhuma chave foi digitada. Encerrando.")
    sys.exit()

if verificar_chave(chave_digitada):
    salvar_chave_local(chave_digitada)
    pyautogui.alert("✅ Chave válida! Bot liberado!")
else:
    pyautogui.alert("❌ Chave inválida, inativa ou expirada. Bot bloqueado.")
    sys.exit()


# === FUNÇÃO DE ENVIO DE E-MAIL DE ALERTA ===
def enviar_email_alerta():
    email_remetente = ""
    senha = ""  # senha de app do Gmail
    email_destinatario = ""
    mensagem = MIMEMultipart()
    mensagem["From"] = email_remetente
    mensagem["To"] = email_destinatario
    mensagem["Subject"] = "REPORT ALERTA!"
    corpo = "O REPORT SERÁ ENVIADO EM 10 MINUTOS!. 🚀"
    mensagem.attach(MIMEText(corpo, "plain"))
    try:
        servidor = smtplib.SMTP("")
        servidor.starttls()
        servidor.login(email_remetente, senha)
        servidor.sendmail(email_remetente, email_destinatario, mensagem.as_string())
        servidor.quit()
        escrever_log("✅ E-mail de alerta enviado com sucesso!")
    except Exception as e:
        escrever_log(f"❌ Erro ao enviar e-mail de alerta: {e}")


# === CONFIGURAÇÕES ===
CAMINHO_PRINT = "tela.jpg"
CAMINHO_PDF = "tela.pdf"


import json

# Planilhas dos Clientes da empresa.
PLANILHAS_DEFAULT = {
    "xxxxxxxx": "",
    "xxxxxxxx": "",
    "xxxxxxxx": "",
    "xxxxxxxx": ""
}

PLANILHAS_PATH = "planilhas_links.json"

def carregar_planilhas():
    if os.path.exists(PLANILHAS_PATH):
        try:
            with open(PLANILHAS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            # Garante que todas as chaves existam
            for k in PLANILHAS_DEFAULT:
                if k not in data:
                    data[k] = PLANILHAS_DEFAULT[k]
            return data
        except Exception as e:
            print(f"[ERRO] Falha ao carregar links salvos: {e}")
    return PLANILHAS_DEFAULT.copy()

def salvar_planilhas(planilhas):
    try:
        with open(PLANILHAS_PATH, "w", encoding="utf-8") as f:
            json.dump(planilhas, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[ERRO] Falha ao salvar links: {e}")

PLANILHAS = carregar_planilhas()

NOME_GRUPO = "GrupoTest"
MENSAGEM = "Segue o print atualizado [gráfico]"


def selecionar_planilhas():
    def renovar_links():
        top = tk.Toplevel(root)
        top.title("Renovação de links das planilhas")
        top.geometry("500x300")
        entries = {}
        for idx, (nome, link) in enumerate(PLANILHAS.items()):
            tk.Label(top, text=nome, font=("Arial", 10, "bold")).grid(row=idx, column=0, sticky="w", padx=10, pady=5)
            entry = tk.Entry(top, width=60)
            entry.insert(0, link)
            entry.grid(row=idx, column=1, padx=5, pady=5)
            entries[nome] = entry
        def salvar():
            for nome, entry in entries.items():
                PLANILHAS[nome] = entry.get().strip()
            salvar_planilhas(PLANILHAS)
            top.destroy()
            messagebox.showinfo("Links atualizados", "Os links das planilhas foram atualizados com sucesso!")
        btn_salvar = tk.Button(top, text="Salvar", command=salvar, font=("Arial", 10), bg="#2196F3", fg="white")
        btn_salvar.grid(row=len(PLANILHAS), column=0, columnspan=2, pady=15)

    root = tk.Tk()
    root.title("Seleção de Planilhas para Envio")
    root.geometry("350x420")
    tk.Label(root, text="Selecione as planilhas:", font=("Arial", 12)).pack(pady=10)

    def ao_fechar():
        root.destroy()
        sys.exit()
    root.protocol("WM_DELETE_WINDOW", ao_fechar)
    planilha_vars = {}
    for nome in PLANILHAS.keys():
        var = tk.BooleanVar()
        chk = tk.Checkbutton(root, text=nome, variable=var, font=("Arial", 10))
        chk.pack(anchor="w", padx=20)
        planilha_vars[nome] = var
    # Opção para mostrar ou ocultar o navegador
    mostrar_bot_var = tk.BooleanVar(value=True)
    chk_mostrar_bot = tk.Checkbutton(root, text="Mostrar o bot em ação (navegador visível)", variable=mostrar_bot_var, font=("Arial", 10))
    chk_mostrar_bot.pack(anchor="w", padx=20, pady=10)
    # Opção para manter login do WhatsApp
    manter_login_var = tk.BooleanVar(value=True)
    chk_manter_login = tk.Checkbutton(root, text="Manter login do WhatsApp (evita QR Code)", variable=manter_login_var, font=("Arial", 10))
    chk_manter_login.pack(anchor="w", padx=20, pady=5)
    # Opção de teste rápido
    teste_rapido_var = tk.BooleanVar(value=False)
    chk_teste_rapido = tk.Checkbutton(root, text="Teste rápido (enviar imediatamente)", variable=teste_rapido_var, font=("Arial", 10))
    chk_teste_rapido.pack(anchor="w", padx=20, pady=5)
    # Opção de margem do print
    margem_var = tk.StringVar(value="poucas")
    frame_margem = tk.LabelFrame(root, text="Margem do print (para baixo)", font=("Arial", 10))
    frame_margem.pack(anchor="w", padx=20, pady=10, fill="x")
    tk.Radiobutton(frame_margem, text="Poucas linhas (400px)", variable=margem_var, value="poucas", font=("Arial", 10)).pack(anchor="w")
    tk.Radiobutton(frame_margem, text="Muitas linhas (550px)", variable=margem_var, value="muitas", font=("Arial", 10)).pack(anchor="w")

    selecionadas = []
    def iniciar():
        selecionadas.clear()
        for nome, var in planilha_vars.items():
            if var.get():
                selecionadas.append(nome)
        if not selecionadas:
            messagebox.showwarning("Seleção obrigatória", "Selecione pelo menos uma planilha!")
            return
        root.destroy()
    frame_botoes = tk.Frame(root)
    frame_botoes.pack(pady=15)
    btn = tk.Button(frame_botoes, text="Iniciar", command=iniciar, font=("Arial", 11), bg="#4CAF50", fg="white")
    btn.pack(side="left", padx=10)
    btn_renovar = tk.Button(frame_botoes, text="Renovação de links", command=renovar_links, font=("Arial", 10), bg="#FFC107", fg="black")
    btn_renovar.pack(side="left", padx=10)
    root.mainloop()
    # Retorna selecionadas e as opções
    return selecionadas, mostrar_bot_var.get(), manter_login_var.get(), teste_rapido_var.get(), margem_var.get()

def escrever_log(msg):
    print(msg)

def iniciar_chrome(mostrar_bot=True, manter_login=True):
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--lang=pt-BR")
    # Se manter_login, usar perfil fixo
    if manter_login:
        profile_path = os.path.abspath("chrome_profile")
        os.makedirs(profile_path, exist_ok=True)
        chrome_options.add_argument(f"--user-data-dir={profile_path}")
    # Headless se mostrar_bot for False (independente do manter_login)
    if not mostrar_bot:
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver


# Função para selecionar a aba pela data do dia
def selecionar_aba_por_data(driver):
    from datetime import datetime
    import unicodedata
    hoje = datetime.now()
    # Gera variações de data: 19/10, 19-10, 19.10, 19_10, 19 Outubro, etc
    dia = hoje.day
    mes = hoje.month
    meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    mes_nome = meses[mes-1]
    padroes = [
        f"{dia:02d}/{mes:02d}",
        f"{dia:02d}-{mes:02d}",
        f"{dia:02d}.{mes:02d}",
        f"{dia:02d}_{mes:02d}",
        f"{dia:02d} {mes_nome}",
        f"{dia} {mes_nome}",
        f"{dia:02d}{mes:02d}",
        f"{dia} de {mes_nome}",
        f"{dia}/{mes}",
        f"{dia}-{mes}",
        f"{dia}.{mes}",
        f"{dia}_{mes}",
    ]
    # Remove acentos para comparar
    def normalizar(txt):
        return unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII').lower()
    try:
        # Já deve estar dentro do iframe do Excel Online
        abas = driver.find_elements(By.CSS_SELECTOR, '[role="tab"], [data-sheet-tab-name], .ms-SheetTab')
        escrever_log(f"[DEBUG] {len(abas)} abas encontradas para seleção por data.")
        for aba in abas:
            nome = aba.text.strip()
            nome_norm = normalizar(nome)
            for padrao in padroes:
                if padrao in nome_norm:
                    try:
                        aba.click()
                        escrever_log(f"✅ Aba selecionada automaticamente: {nome}")
                        time.sleep(2)
                        return True
                    except Exception as e:
                        escrever_log(f"[AVISO] Falha ao clicar na aba '{nome}': {e}")
        escrever_log("[AVISO] Nenhuma aba correspondente à data encontrada. Prosseguindo sem seleção automática.")
    except Exception as e:
        escrever_log(f"[ERRO] Erro ao tentar selecionar aba por data: {e}")
    return False

def tirar_print(driver, link):
    tentativas = 0
    max_tentativas = 3
    while tentativas < max_tentativas:
        driver.get(link)
        escrever_log(f"Aguardando login manual na planilha (OneDrive/Microsoft)... (tentativa {tentativas+1})")
        time.sleep(30)  # Tempo para login manual na primeira execução
        # Verifica se a página carregou corretamente
        erro_detectado = False
        try:
            # Tenta encontrar o iframe do Excel Online
            driver.switch_to.frame(driver.find_element(By.XPATH, "//iframe[contains(@id, 'WacFrame_Excel')]") )
            escrever_log("Entrou no iframe do Excel Online.")
        except Exception:
            erro_detectado = True
            escrever_log("[AVISO] Não foi possível acessar o iframe do Excel Online. Tentando recarregar...")
        # Verifica se o título da página indica erro
        titulo = driver.title.lower()
        if "erro" in titulo or "not found" in titulo or "não encontrado" in titulo or "conexão" in titulo:
            erro_detectado = True
            escrever_log(f"[AVISO] Título da página indica erro: {driver.title}")
        if not erro_detectado:
            break
        tentativas += 1
        time.sleep(5)
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
    if tentativas == max_tentativas:
        escrever_log(f"[ERRO] Não foi possível carregar a planilha após {max_tentativas} tentativas. Pulando...")
        return
    # Selecionar aba pela data do dia
    try:
        selecionar_aba_por_data(driver)
    except Exception as e:
        escrever_log(f"[AVISO] Erro ao tentar selecionar aba por data: {e}")
    # Voltar para o contexto principal para o restante do script
    try:
        driver.switch_to.default_content()
    except Exception:
        pass
    # Resetar zoom no frame principal conforme a margem escolhida
    zoom = '59%' if (not hasattr(tirar_print, 'altura') or tirar_print.altura == 400) else '50%'
    driver.execute_script(f"document.body.style.zoom='{zoom}'")
    time.sleep(1)
    # Scrollar para o topo e esquerda do conteúdo da planilha dentro do iframe
    try:
        iframe = driver.find_element(By.XPATH, "//iframe[contains(@id, 'WacFrame_Excel')]")
        driver.switch_to.frame(iframe)
        # 1. Forçar scroll em todos os <div> visíveis
        divs = driver.find_elements(By.TAG_NAME, 'div')
        count = 0
        for div in divs:
            try:
                driver.execute_script("arguments[0].scrollLeft = 0; arguments[0].scrollTop = 0;", div)
                count += 1
            except Exception:
                pass
        escrever_log(f"[DEBUG] Scroll forçado em {count} divs.")
        # 2. Simular teclas Home e Ctrl+Home usando Actions
        try:
            from selenium.webdriver.common.action_chains import ActionChains
            from selenium.webdriver.common.keys import Keys
            body = driver.find_element(By.TAG_NAME, 'body')
            ActionChains(driver).move_to_element(body).send_keys(Keys.HOME).perform()
            escrever_log("[DEBUG] Tecla HOME enviada.")
            time.sleep(0.5)
            ActionChains(driver).move_to_element(body).key_down(Keys.CONTROL).send_keys(Keys.HOME).key_up(Keys.CONTROL).perform()
            escrever_log("[DEBUG] Ctrl+HOME enviada.")
        except Exception as e:
            escrever_log(f"[DEBUG] Não foi possível enviar teclas HOME/Ctrl+HOME: {e}")
        # 3. Tenta todos os elementos possíveis para scroll horizontal/vertical
        scroll_elements = driver.find_elements(By.CSS_SELECTOR, '[role="grid"], .office-scrollableContainer')
        escrever_log(f"[DEBUG] Elementos de scroll encontrados: {len(scroll_elements)}")
        for el in scroll_elements:
            driver.execute_script("arguments[0].scrollLeft = 0; arguments[0].scrollTop = 0;", el)
        time.sleep(1)
        driver.switch_to.default_content()
    except Exception as e:
        escrever_log(f"[AVISO] Não foi possível scrollar dentro do iframe: {e}")
    time.sleep(1)
    driver.save_screenshot(CAMINHO_PRINT)
    x, y, largura = 0, 100, 1450
    # Usa a altura definida globalmente
    altura = tirar_print.altura if hasattr(tirar_print, 'altura') else 400
    img = Image.open(CAMINHO_PRINT)
    crop = img.crop((x, y, x + largura, y + altura))
    # Salva como JPG para garantir prévia no WhatsApp
    crop.convert('RGB').save(CAMINHO_PRINT, 'JPEG')
    escrever_log(f"📸 Print tirado! (Altura: {altura}px)")

def whatsapp_selenium(driver):
    escrever_log("✅ Abrindo WhatsApp Web...")
    driver.get("https://web.whatsapp.com")
    wait = WebDriverWait(driver, 120)
    caixa_pesquisa = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div[contenteditable='true']"))
    )
    time.sleep(1)
    caixa_pesquisa.send_keys(NOME_GRUPO)
    escrever_log("Aguardando grupo aparecer...")
    grupo = wait.until(
        EC.presence_of_element_located((By.XPATH, f"//span[@title='{NOME_GRUPO}']"))
    )
    driver.execute_script("arguments[0].scrollIntoView();", grupo)
    time.sleep(0.5)
    try:
        grupo.click()
    except Exception:
        driver.execute_script("arguments[0].click();", grupo)
    escrever_log("Aguardando campo de texto do chat...")
    caixa_texto = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div[contenteditable='true'][data-tab]"))
    )
    time.sleep(1)
    caixa_texto.click()
    # Só envia a mensagem no primeiro envio
    if whatsapp_selenium.primeiro_envio:
        caixa_texto.send_keys(MENSAGEM)
        escrever_log("Mensagem digitada. Aguardando botão de anexo...")
        whatsapp_selenium.primeiro_envio = False
    else:
        escrever_log("Aguardando botão de anexo...")
    anexo = wait.until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div/div[1]/div/span/div/div/div[1]/div[1]/span'))
    )
    anexo.click()
    escrever_log("Aguardando ícone 'Fotos e Vídeos'...")
    fotos_videos = wait.until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div/span[6]/div/ul/div/div/div[2]/li/div/span'))
    )
    fotos_videos.click()
    escrever_log("Aguardando campo de upload de imagem...")
    # Garante que está pegando o input correto após clicar em Fotos e Vídeos
    upload = wait.until(
        EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'))
    )
    upload.send_keys(os.path.abspath(CAMINHO_PRINT))
    escrever_log("Imagem anexada. Aguardando processamento do WhatsApp...")
    time.sleep(2)  # Delay extra para garantir que a janela de seleção seja fechada
    escrever_log("Aguardando botão de enviar...")
    # Espera o botão de enviar aparecer com o novo XPath informado
    send_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div/div[3]/div/div[2]/div[2]/div/span/div/div/div/div[2]/div/div[2]/div[2]/div/div/span'))
    )
    send_button.click()
    escrever_log("Arquivo enviado!")
    time.sleep(3)

def apagar_local():
    if os.path.exists(CAMINHO_PRINT):
        os.remove(CAMINHO_PRINT)
    if os.path.exists(CAMINHO_PDF):
        os.remove(CAMINHO_PDF)
    escrever_log("✅ Arquivos enviados e deletados do PC!")

def processar_planilhas(selecionadas, mostrar_bot, manter_login):
    driver = iniciar_chrome(mostrar_bot, manter_login)
    whatsapp_selenium.primeiro_envio = True
    for nome in selecionadas:
        escrever_log(f"📄 Processando planilha: {nome}")
        tirar_print(driver, PLANILHAS[nome])
        whatsapp_selenium(driver)
        apagar_local()
    driver.quit()
    escrever_log(f"✅ Todas as planilhas enviadas: {datetime.now().strftime('%H:%M:%S')}")


# === AGENDAMENTO AUTOMÁTICO ===
import threading
from datetime import timedelta

def gerar_horarios_envio():
    horarios = []
    hora = 8
    minuto = 30
    while True:
        horarios.append(f"{hora:02d}:{minuto:02d}")
        hora += 2
        if hora > 22:
            break
    return horarios

def proximo_horario_envio(horarios):
    agora = datetime.now()
    hoje = agora.date()
    for h in horarios:
        hora, minuto = map(int, h.split(":"))
        envio = datetime.combine(hoje, datetime.min.time()) + timedelta(hours=hora, minutes=minuto)
        if agora < envio:
            return envio
    # Se passou de todos, retorna o primeiro horário do próximo dia
    return datetime.combine(hoje + timedelta(days=1), datetime.min.time()) + timedelta(hours=8, minutes=30)

def loop_agendamento():
    # Mostra interface apenas uma vez no início
    selecionadas, mostrar_bot, manter_login, teste_rapido, margem = None, None, None, None, None
    resultado = selecionar_planilhas()
    if isinstance(resultado, tuple) and len(resultado) == 5:
        selecionadas, mostrar_bot, manter_login, teste_rapido, margem = resultado
    else:
        # Compatibilidade caso não tenha sido atualizado
        selecionadas, mostrar_bot = resultado[:2]
        manter_login = resultado[2] if len(resultado) > 2 else True
        teste_rapido = resultado[3] if len(resultado) > 3 else False
        margem = "poucas"
    # Define altura global para tirar_print
    tirar_print.altura = 400 if margem == "poucas" else 550
    horarios = gerar_horarios_envio()
    escrever_log(f"Horários de envio configurados: {horarios}")
    if teste_rapido:
        escrever_log("[TESTE RÁPIDO] Enviando imediatamente!")
        processar_planilhas(selecionadas, mostrar_bot, manter_login)
        return
    # Controle para não enviar o mesmo alerta mais de uma vez
    alertas_enviados = set()
    parar_automacao = threading.Event()

    def mostrar_parar_automacao():
        janela_parar = tk.Tk()
        janela_parar.title("Parar Automação")
        janela_parar.geometry("200x100")
        janela_parar.resizable(False, False)
        tk.Label(janela_parar, text="Automação em execução", font=("Arial", 10)).pack(pady=10)
        def parar():
            parar_automacao.set()
            janela_parar.destroy()
        btn = tk.Button(janela_parar, text="Parar Automação", command=parar, font=("Arial", 12), bg="red", fg="white")
        btn.pack(pady=10)
        def ao_fechar():
            parar_automacao.set()
            janela_parar.destroy()
        janela_parar.protocol("WM_DELETE_WINDOW", ao_fechar)
        janela_parar.mainloop()

    # Thread para mostrar a interface de parada
    threading.Thread(target=mostrar_parar_automacao, daemon=True).start()

    while True:
        if parar_automacao.is_set():
            escrever_log("Automação parada pelo usuário. Reiniciando...")
            break
        agora = datetime.now()
        proximo = proximo_horario_envio(horarios)
        # Checar se é hora de enviar alerta (10 minutos antes de cada horário)
        for h in horarios:
            hora, minuto = map(int, h.split(":"))
            alerta = datetime.combine(agora.date(), datetime.min.time()) + timedelta(hours=hora, minutes=minuto) - timedelta(minutes=10)
            if 0 <= (alerta - agora).total_seconds() < 60 and h not in alertas_enviados:
                enviar_email_alerta()
                alertas_enviados.add(h)
        delta = (proximo - agora).total_seconds()
        if delta > 0:
            for _ in range(int(min(delta, 60))):
                if parar_automacao.is_set():
                    break
                time.sleep(1)
            if parar_automacao.is_set():
                escrever_log("Automação parada pelo usuário. Reiniciando...")
                break
            continue
        escrever_log(f"[AGENDADO] Iniciando envio das planilhas em {proximo.strftime('%d/%m/%Y %H:%M')}")
        processar_planilhas(selecionadas, mostrar_bot, manter_login)
        # Aguarda 1 minuto para evitar múltiplos envios no mesmo horário
        for _ in range(60):
            if parar_automacao.is_set():
                break
            time.sleep(1)
        if parar_automacao.is_set():
            escrever_log("Automação parada pelo usuário. Reiniciando...")
            break
        # Limpa alertas antigos para o próximo ciclo diário
        if proximo.time().strftime("%H:%M") == horarios[0]:
            alertas_enviados.clear()

if __name__ == "__main__":
    while True:
        loop_agendamento()