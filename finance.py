import time
import pyautogui
import pandas as pd
from openpyxl import load_workbook

def short(): time.sleep(0.2)
def long(): time.sleep(1)
def wait(): time.sleep(3)

rel_name = ""

def create_csv():
    global rel_name
    init_date = input("DATA INICIAL (Exemplo: 30122025):\n")
    fin_date = input("DATA FINAL (Exemplo: 30122025):\n")
    rel_name = init_date + " - " + fin_date

    pyautogui.hotkey("alt", "t")
    short()
    pyautogui.press("x")
    short()
    pyautogui.press("c")
    long()
    pyautogui.write(init_date)
    long()
    pyautogui.press("enter")
    long()
    pyautogui.write(fin_date)
    long()
    pyautogui.press("1")
    long()
    pyautogui.press("enter")
    long()
    pyautogui.hotkey("ctrl", "p")
    long()
    pyautogui.write(rel_name)
    long()
    pyautogui.press("enter")
    long()
    pyautogui.press("enter")
    wait()


def clean_sort():
    try:
        df = pd.read_csv(f"{rel_name}.csv", encoding="utf-8")

        df_sorted = df.sort_values(by='Valor Pendente', ascending=False)

        df_sorted.to_excel("final.xlsx", index=False, engine='openpyxl')
        print("\nArquivo salvo como 'final.xlsx'")
    except Exception as e:
        print(f"\nErro ao processar arquivo: {e}")


def block_clients():
    try:
        df = pd.read_excel("final.xlsx", engine='openpyxl')
        client_codes = df.iloc[:, 0]  # Column A (index 0)

        pyautogui.click(x=100, y=100)
        long
        pyautogui.click(x=100, y=100)
        long
        pyautogui.press("left")
        pyautogui.press("left")
        pyautogui.press("left")
        pyautogui.press("left")

        for client_code in client_codes:
            print(f"Bloqueando cliente: {client_code}")
            # Navegação para "Pessoas > Clientes"
            pyautogui.write(str(client_code))
            long()
            pyautogui.press("enter")
            long()
            pyautogui.press("f5")  # Assuming this refreshes or opens detail
            long()
            pyautogui.press("enter")  # Confirm block
            long()
        
        pyautogui.press("esc")

    except Exception as e:
        print(f"\nErro ao bloquear clientes: {e}")


def start_macro():
    create_csv()
    wait
    clean_sort()
    wait
    block_clients
    print("Done.\n")