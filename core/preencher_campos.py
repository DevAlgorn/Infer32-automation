import pandas as pd
import pyautogui
import time

# Função para ler o arquivo IW3
def ler_iw3_arquivo(iw3_file_path):
    iw3_data = pd.read_csv(iw3_file_path)  # Supondo que IW3 seja um arquivo CSV
    return iw3_data

# Função para ler o arquivo Excel
def ler_excel_arquivo(excel_file_path):
    excel_data = pd.read_excel(excel_file_path)
    return excel_data

# Função para cruzar as colunas entre os arquivos
def cruzar_colunas(iw3_data, excel_data):
    iw3_colunas = iw3_data.columns.tolist()
    excel_colunas = excel_data.columns.tolist()
    
    colunas_comum = set(iw3_colunas) & set(excel_colunas)
    print(f"Colunas comuns entre o IW3 e o Excel: {colunas_comum}")
    return colunas_comum

# Função para preencher os campos no infer-32
def preencher_campos(iw3_file_path, excel_file_path):
    # Ler os dados do arquivo IW3
    iw3_data = ler_iw3_arquivo(iw3_file_path)
    print(f"Dados do arquivo IW3: {iw3_data}")
    
    # Ler os dados do arquivo Excel
    excel_data = ler_excel_arquivo(excel_file_path)
    print(f"Dados do arquivo Excel: {excel_data}")
    
    # Cruzar as colunas
    colunas_comum = cruzar_colunas(iw3_data, excel_data)
    
    if colunas_comum:
        print("Campos para preenchimento encontrados!")
        
        # A partir daqui, podemos preencher os campos no Infer-32 utilizando pyautogui
        for index, row in excel_data.iterrows():
            for col in colunas_comum:
                dado = row[col]  # Pega o dado correspondente
                pyautogui.write(str(dado))  # Escreve no campo de texto
                pyautogui.press('tab')  # Passa para o próximo campo
                print(f"Preenchendo: {col} com o valor {dado}")
                time.sleep(1)  # Atraso para não sobrecarregar a interface
    else:
        print("Nenhum campo para preenchimento encontrado.")
