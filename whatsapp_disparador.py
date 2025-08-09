import os
import re
import sys
import time
import random
from datetime import datetime
from urllib.parse import quote_plus

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from colorama import Fore, Style, init
init(autoreset=True)

import pygetwindow as gw

if getattr(sys, 'frozen', False):
    dir_path = os.path.dirname(sys.executable)
else:
    dir_path = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE = os.path.join(dir_path, "contacts.xlsx")
OUTPUT_FILE = os.path.join(dir_path, "contacts_sent.xlsx")

PHONE_COL_INDEX = 5  
CNPJ_COL = "CNPJ"
RAZAO_COL = "Razao Social"
PLANO_COL = "Plano"
ENDERECO_COL = "Endereco"

MESSAGE_TEMPLATE = (
    "Olá {razao}, tudo bem? Aqui é da Vivo. Observamos que o período de fidelidade do plano {plano} "
    "(CNPJ {cnpj}) expirou. Queremos oferecer uma renovação com planos maiores e benefícios. "
    "Endereço cadastrado: {endereco}. Posso encaminhar uma proposta rápida?"
)


def configurar_delay():
    print(Fore.CYAN + "\nConfiguração de delay entre mensagens:")
    while True:
        try:
            delay_min = float(input(Fore.GREEN + "Digite o delay mínimo em segundos (ex: 4): ").strip())
            delay_max = float(input(Fore.GREEN + "Digite o delay máximo em segundos (ex: 7): ").strip())
            if delay_min <= 0 or delay_max <= 0:
                print(Fore.RED + "Valores devem ser maiores que zero.")
                continue
            if delay_min > delay_max:
                print(Fore.RED + "O delay mínimo não pode ser maior que o máximo.")
                continue
            print(Fore.YELLOW + f"Delay configurado: entre {delay_min}s e {delay_max}s.\n")
            return delay_min, delay_max
        except ValueError:
            print(Fore.RED + "Digite apenas números válidos.")


def clean_phone(phone_raw: str) -> str:
    s = str(phone_raw) if phone_raw else ""
    digits = re.sub(r"\D", "", s)
    if not digits:
        return ""
    if digits.startswith("55") and len(digits) >= 11:
        return digits
    if len(digits) == 11:
        return "55" + digits
    if len(digits) == 10:
        return "55" + digits
    if len(digits) < 10:
        return ""
    return digits


def get_screen_resolution():
    try:
        import ctypes
        user32 = ctypes.windll.user32
        width = user32.GetSystemMetrics(0)
        height = user32.GetSystemMetrics(1)
        return width, height - 100
    except Exception:
        return 1366, 660


def prepare_driver(profile_num: int, profile_dir: str, headless: bool = False):
    width, height = get_screen_resolution()

    if profile_num == 1:
        from selenium.webdriver.chrome.service import Service as ChromeService
        from selenium.webdriver.chrome.options import Options as ChromeOptions
        from webdriver_manager.chrome import ChromeDriverManager

        options = ChromeOptions()
        if headless:
            options.add_argument("--headless=new")
        options.add_argument(f"--user-data-dir={profile_dir}")
        options.add_argument("--disable-extensions")
        options.add_argument("--profile-directory=Default")
        options.add_argument(f"--window-size={width},{height}")
        options.add_argument("--window-position=0,0")
        options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        driver.set_window_size(width, height)
        driver.set_window_position(0, 0)

    elif profile_num == 2:
        from selenium.webdriver.edge.options import Options as EdgeOptions
        from selenium.webdriver.edge.service import Service as EdgeService

        options = EdgeOptions()
        if headless:
            options.add_argument("--headless=new")
        options.add_argument(f"--user-data-dir={profile_dir}")
        options.add_argument("--disable-extensions")
        options.add_argument("--profile-directory=Default")
        options.add_argument(f"--window-size={width},{height}")
        options.add_argument("--window-position=0,0")
        options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        driver_path = os.path.join(dir_path, "drivers", "msedgedriver.exe")
        service = EdgeService(executable_path=driver_path)
        driver = webdriver.Edge(service=service, options=options)
        driver.set_window_size(width, height)
        driver.set_window_position(0, 0)

    elif profile_num == 3:
        from selenium.webdriver.chrome.service import Service as BraveService
        from selenium.webdriver.chrome.options import Options as BraveOptions
        from webdriver_manager.chrome import ChromeDriverManager

        options = BraveOptions()
        if headless:
            options.add_argument("--headless=new")
        brave_path = r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"
        options.binary_location = brave_path
        options.add_argument(f"--user-data-dir={profile_dir}")
        options.add_argument("--disable-extensions")
        options.add_argument("--profile-directory=Default")
        options.add_argument(f"--window-size={width},{height}")
        options.add_argument("--window-position=0,0")
        options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        service = BraveService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        driver.set_window_size(width, height)
        driver.set_window_position(0, 0)

    else:
        raise ValueError("Perfil inválido")

    return driver


def exit_fullscreen(driver):
    try:
        driver.execute_script("""
            if (document.fullscreenElement) {
                document.exitFullscreen();
            }
        """)
    except Exception:
        pass


def wait_for_login(driver, timeout=60):
    driver.get("https://web.whatsapp.com")
    exit_fullscreen(driver)
    wait = WebDriverWait(driver, timeout)
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[contenteditable='true']")))
        return True
    except Exception:
        return False


def send_message_to(driver, phone: str, message: str, wait: WebDriverWait) -> bool:
    url = f"https://web.whatsapp.com/send?phone={phone}&text={quote_plus(message)}&app_absent=0"
    driver.get(url)
    exit_fullscreen(driver)
    try:
        input_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[contenteditable='true']")))
        time.sleep(1)
        send_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[aria-label='Enviar']")))
        send_button.click()
        time.sleep(2)
        return True
    except Exception as e:
        print(f"Erro ao enviar para {phone}: {e}")
        return False


def escolher_perfil():
    print(Fore.CYAN + "Escolha o perfil do WhatsApp para usar:")
    print(Fore.CYAN + "1 - WhatsApp 1 (Chrome)")
    print(Fore.CYAN + "2 - WhatsApp 2 (Edge)")
    print(Fore.CYAN + "3 - WhatsApp 3 (Brave)")
    print(Fore.CYAN + "4 - Todos os 3 perfis (envio intercalado)")
    escolha = input(Fore.GREEN + "Digite 1, 2, 3 ou 4: ").strip()
    if escolha == "1":
        return 1, os.path.join(dir_path, "whatsapp_profile")
    elif escolha == "2":
        return 2, os.path.join(dir_path, "whatsapp_profile_2")
    elif escolha == "3":
        return 3, os.path.join(dir_path, "whatsapp_profile_3")
    elif escolha == "4":
        return 4, None
    else:
        print(Fore.RED + "Opção inválida, usando perfil padrão (WhatsApp 1 - Chrome).")
        return 1, os.path.join(dir_path, "whatsapp_profile")


def posicionar_janelas():
    width, height = 960, 540  
    x, y = 0, 0               
    time.sleep(5)           

    nomes_janelas = ["Chrome", "Microsoft Edge", "Brave"]

    for nome in nomes_janelas:
        janelas = [w for w in gw.getWindowsWithTitle(nome) if w.visible]
        if not janelas:
            print(f"Janela '{nome}' não encontrada para posicionar.")
            continue
        janela = janelas[0]
        try:
            janela.moveTo(x, y)
            janela.resizeTo(width, height)
            print(f"{nome} posicionada em ({x},{y}) tamanho {width}x{height}")
        except Exception as e:
            print(f"Erro ao posicionar janela {nome}: {e}")


def main():
    print(Fore.MAGENTA + "Bem vindo(a) Ao disparador de mensagens! Desenvolvido por Arthur! (Está na beta, desconsidere os erros)\n")

    delay_min, delay_max = configurar_delay()

    perfil_escolhido, profile_dir = escolher_perfil()

    if not os.path.exists(INPUT_FILE):
        print(Fore.RED + f"Arquivo {INPUT_FILE} não encontrado.")
        return

    df = pd.read_excel(INPUT_FILE, dtype=str)
    for col in ("Status", "SentAt", "Mensagem"):
        if col not in df.columns:
            df[col] = ""

    if df.shape[1] <= PHONE_COL_INDEX:
        print(Fore.RED + f"A planilha não tem coluna na posição {PHONE_COL_INDEX}.")
        return

    phones = df.iloc[:, PHONE_COL_INDEX].astype(str).fillna("").tolist()
    total = len(df)

    if perfil_escolhido in [1, 2, 3]:
        driver = prepare_driver(perfil_escolhido, profile_dir)
        wait = WebDriverWait(driver, 60)
        print(Fore.YELLOW + "Abra o WhatsApp Web no navegador aberto e escaneie o QR code, se necessário.")
        if not wait_for_login(driver):
            print(Fore.RED + "Não foi possível detectar o login no WhatsApp Web.")
            driver.quit()
            return

        for idx in range(total):
            phone_raw = phones[idx]
            phone = clean_phone(phone_raw)
            cnpj = df.at[idx, CNPJ_COL] if CNPJ_COL in df.columns else ""
            razao = df.at[idx, RAZAO_COL] if RAZAO_COL in df.columns else ""
            plano = df.at[idx, PLANO_COL] if PLANO_COL in df.columns else ""
            endereco = df.at[idx, ENDERECO_COL] if ENDERECO_COL in df.columns else ""

            message = MESSAGE_TEMPLATE.format(
                razao=razao or "cliente",
                plano=plano or "seu plano",
                cnpj=cnpj or "",
                endereco=endereco or "seu endereço cadastrado"
            )
            df.at[idx, "Mensagem"] = message

            if not phone:
                df.at[idx, "Status"] = "Telefone inválido"
                df.at[idx, "SentAt"] = datetime.now().isoformat()
                print(Fore.RED + f"[{idx + 1}/{total}] {razao} - telefone inválido, pulando.")
                continue

            print(Fore.CYAN + f"[{idx + 1}/{total}] Enviando para {razao} ({phone}) ...")
            success = send_message_to(driver, phone, message, wait)
            if success:
                df.at[idx, "Status"] = "Enviado"
                df.at[idx, "SentAt"] = datetime.now().isoformat()
                print(Fore.GREEN + "  -> Enviado com sucesso.")
            else:
                df.at[idx, "Status"] = "Falha"
                df.at[idx, "SentAt"] = datetime.now().isoformat()
                print(Fore.RED + "  -> Falha ao enviar.")

            delay = random.uniform(delay_min, delay_max)
            print(Fore.YELLOW + f"  Aguardando {delay:.1f}s...")
            time.sleep(delay)

        df.to_excel(OUTPUT_FILE, index=False)
        print(Fore.MAGENTA + f"Processo finalizado. Resultados salvos em {OUTPUT_FILE}")
        driver.quit()

    elif perfil_escolhido == 4:
        profiles = [
            (1, os.path.join(dir_path, "whatsapp_profile")),
            (2, os.path.join(dir_path, "whatsapp_profile_2")),
            (3, os.path.join(dir_path, "whatsapp_profile_3")),
        ]

        drivers = []
        waits = []
        for profile_num, prof_dir in profiles:
            driver = prepare_driver(profile_num, prof_dir)
            drivers.append(driver)
            waits.append(WebDriverWait(driver, 60))
            print(Fore.YELLOW + f"Abrindo WhatsApp Web no perfil {profile_num}...")
            driver.get("https://web.whatsapp.com")
            exit_fullscreen(driver)
            time.sleep(10)

        posicionar_janelas()

        print(Fore.YELLOW + "Faça login em cada WhatsApp Web (QR code).")
        for i, driver in enumerate(drivers):
            print(Fore.YELLOW + f"Aguardando login no perfil {i+1}...")
            if not wait_for_login(driver):
                print(Fore.RED + f"Não foi possível detectar login no perfil {i+1}.")

        for idx in range(total):
            driver_idx = idx % 3
            driver = drivers[driver_idx]
            wait = waits[driver_idx]

            phone_raw = phones[idx]
            phone = clean_phone(phone_raw)
            cnpj = df.at[idx, CNPJ_COL] if CNPJ_COL in df.columns else ""
            razao = df.at[idx, RAZAO_COL] if RAZAO_COL in df.columns else ""
            plano = df.at[idx, PLANO_COL] if PLANO_COL in df.columns else ""
            endereco = df.at[idx, ENDERECO_COL] if ENDERECO_COL in df.columns else ""

            message = MESSAGE_TEMPLATE.format(
                razao=razao or "cliente",
                plano=plano or "seu plano",
                cnpj=cnpj or "",
                endereco=endereco or "seu endereço cadastrado"
            )
            df.at[idx, "Mensagem"] = message

            if not phone:
                df.at[idx, "Status"] = "Telefone inválido"
                df.at[idx, "SentAt"] = datetime.now().isoformat()
                print(Fore.RED + f"[{idx + 1}/{total}] {razao} - telefone inválido, pulando.")
                continue

            print(Fore.CYAN + f"[{idx + 1}/{total}] Enviando para {razao} ({phone}) no perfil {driver_idx + 1} ...")
            success = send_message_to(driver, phone, message, wait)
            if success:
                df.at[idx, "Status"] = "Enviado"
                df.at[idx, "SentAt"] = datetime.now().isoformat()
                print(Fore.GREEN + "  -> Enviado com sucesso.")
            else:
                df.at[idx, "Status"] = "Falha"
                df.at[idx, "SentAt"] = datetime.now().isoformat()
                print(Fore.RED + "  -> Falha ao enviar.")

            delay = random.uniform(delay_min, delay_max)
            print(Fore.YELLOW + f"  Aguardando {delay:.1f}s...")
            time.sleep(delay)

        df.to_excel(OUTPUT_FILE, index=False)
        print(Fore.MAGENTA + f"Processo finalizado. Resultados salvos em {OUTPUT_FILE}")

        for driver in drivers:
            driver.quit()


if __name__ == "__main__":
    main()
