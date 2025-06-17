import time
import pyautogui
import pandas as pd
from openpyxl import load_workbook

def short(): time.sleep(2)
def long(): time.sleep(2)
def wait(): time.sleep(5)

rel_name = ""

def create_csv():
    global rel_name
    init_date = input("DATA INICIAL (Exemplo: 30122025):\n")
    fin_date = input("DATA FINAL (Exemplo: 30122025):\n")
    rel_name = init_date + " - " + fin_date

    pyautogui.click(x=100, y=100)
    long()
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
    pyautogui.press("enter")
    long()
    pyautogui.press("1")
    long()
    pyautogui.press("enter")
    long()
    pyautogui.hotkey("ctrl", "p")
    long()
    pyautogui.press("enter")
    long()
    pyautogui.write(rel_name)
    long()
    pyautogui.press("enter")
    long()
    pyautogui.press("enter")
    wait()
    print("Done.\n")


def clean_sort():
    try:
        df = pd.read_csv(f"{rel_name}.csv", encoding="utf-8")

        df_sorted = df.sort_values(by='Valor Pendente', ascending=False)

        df_sorted.to_excel("final.xlsx", index=False, engine='openpyxl')
        print("\nSaved: 'final.xlsx'")
    except Exception as e:
        print(f"\nError: {e}")


def block_clients():
    try:
        df = pd.read_excel("final.xlsx", engine='openpyxl')
        client_code_list = df.iloc[:, 0]  # Column A (index 0)

        wait()
        wait()
        # pyautogui.click(x=100, y=100)
        # long()
        # pyautogui.click(x=100, y=100)
        # long()
        pyautogui.press("left")
        pyautogui.press("left")
        pyautogui.press("left")
        pyautogui.press("left")

        for client_code in client_code_list:
            print(f"Block: {client_code}")
            pyautogui.write(str(client_code))
            long()
            pyautogui.press("enter")
            long()
            pyautogui.press("f5")
            long()
            pyautogui.press("enter")
            long()
        
        pyautogui.press("esc")
        print("Done.\n")

    except Exception as e:
        print(f"\nError: {e}")


def start_macro():
    # create_csv()
    clean_sort()
    input("Type anything to start:\n")
    block_clients
    print("Done.\n")


start_macro()