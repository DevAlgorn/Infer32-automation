import time
import pyautogui
import pandas as pd


def preencher_campos(dados):
    """
    Função que recebe um DataFrame e simula o preenchimento de campos usando pyautogui.
    Cada célula é preenchida com TAB após o valor.
    """
    print("Iniciando preenchimento automático...")
    time.sleep(3)  # Tempo para o usuário posicionar o cursor

    for i, linha in dados.iterrows():
        for valor in linha:
            pyautogui.write(str(valor))  # Digita o valor
            pyautogui.press('tab')       # Pressiona TAB para ir ao próximo campo
        pyautogui.press('enter')         # Após a linha, ENTER pode ser usado para garantir o foco correto
        print(f"Linha {i + 1} preenchida.")

    print("Preenchimento finalizado com sucesso.")
