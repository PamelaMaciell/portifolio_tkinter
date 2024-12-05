import tkinter as tk
from tkinter import messagebox
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime

# Função para buscar dados do clima
def buscar_previsao():
    try:
        navegador = webdriver.Chrome()
        navegador.get('https://www.google.com')
        navegador.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys('Temperatura de São Paulo')
        navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[3]/center/input[1]').send_keys(Keys.ENTER)

        temperatura = navegador.find_element(By.XPATH, '/html/body/div[3]/div/div[13]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div/div[1]/div[1]').text
        umidade = navegador.find_element(By.XPATH, '//*[@id="wob_hm"]').text

        navegador.quit()  # Fecha o navegador

        return temperatura, umidade
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao buscar os dados: {e}")
        return None, None

# Função para salvar os dados no Excel
def salvar_dados(data_hora, temperatura, umidade):
    try:
        # Carregar a planilha ou criar uma nova
        try:
            wb = openpyxl.load_workbook("historico_temperatura.xlsx")
        except FileNotFoundError:
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.append(["Data/Hora", "Temperatura", "Umidade"])

        sheet = wb.active
        sheet.append([data_hora, temperatura, umidade])

        # Salvar a planilha
        wb.save("historico_temperatura.xlsx")
    except Exception as e:
        messagebox.showerror("Erro ao salvar", f"Erro ao salvar os dados na planilha: {e}")

# Função para atualizar a interface com os dados
def atualizar_interface():
    temperatura, umidade = buscar_previsao()
    if temperatura and umidade:
        data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        salvar_dados(data_hora, temperatura, umidade)
        lbl_resultado.config(text=f"Temperatura: {temperatura}\nUmidade: {umidade}")
    else:
        lbl_resultado.config(text="Erro ao buscar dados.")

# Função para criar a interface gráfica
def criar_interface():
    global lbl_resultado

    root = tk.Tk()
    root.title("Aplicação de Previsão do Tempo")
    root.geometry("400x200")

    # Botão para buscar a previsão
    btn_buscar = tk.Button(root, text="Buscar Previsão", command=atualizar_interface, font=("Arial", 14))
    btn_buscar.pack(pady=20)

    # Label para exibir os resultados
    lbl_resultado = tk.Label(root, text="", font=("Arial", 12))
    lbl_resultado.pack(pady=20)

    root.mainloop()

# Iniciar a aplicação
if __name__ == "__main__":
    criar_interface()