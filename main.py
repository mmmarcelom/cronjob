import os
import time
import requests
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

LOGIN_URL = r"https://gestao.hubnordeste.tech/Login.aspx?ReturnUrl=%2f%3f"
WEBHOOK = r"https://script.google.com/macros/s/AKfycbxS_FjHCOPZheABAaNhyfQnyddl1r57MtDyLOqs8CyGn-KHgwCD4JaJtwhc_TF2OhvhUA/exec"

load_dotenv()

USUARIO = os.getenv("USUARIO")
SENHA = os.getenv("SENHA")


def obter_token():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    browser.get(LOGIN_URL)
    time.sleep(5)

    # Preenche o login
    browser.find_element(By.ID, "txtEmail").send_keys(USUARIO)
    browser.find_element(By.ID, "txtPass").send_keys(SENHA)
    browser.find_element(By.ID, "Button1").click()

    time.sleep(3)  
    # espera o redirect

    cookies = browser.get_cookies()
    token = next((c["value"] for c in cookies if c["name"] == "erpusertoken"), "")

    browser.quit()
    return token

def enviar_token(token):
    if not token:
        print("❌ Nenhum token capturado")
        return

    resp = requests.post(WEBHOOK, json={"token": token})
    print("Webhook status:", resp.status_code, resp.text)

if __name__ == "__main__":
    token = obter_token()
    print("TOKEN CAPTURADO:", token)
    enviar_token(token)