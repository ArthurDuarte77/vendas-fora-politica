import customtkinter as ctk
from tkcalendar import DateEntry
from tkinter import ttk
from datetime import datetime, timedelta
import subprocess
import os

if os.path.exists('resultado.xlsx'):  
    os.remove("resultado.xlsx")
if os.path.exists('items.xlsx'):  
    os.remove("items.xlsx")

# Função chamada ao clicar no botão
def on_date_selected():
    # Obtém as datas selecionadas
    started_data = start_cal.get_date()
    end_date = end_cal.get_date()

    # Verifica se a data inicial é maior que a final
    if started_data > end_date:
        print("Erro: A data inicial não pode ser maior que a data final.")
        return

    root.destroy()  # Fecha a janela após o loop
    # Percorre cada dia no intervalo
    current_date = started_data
    while current_date <= end_date:
        # Formata as datas no estilo DDMMYYYY
        dataInicial = current_date.strftime('%d%m%Y')
        dataFinal = current_date.strftime('%d%m%Y')

        subprocess.run(["python", "main.py", dataInicial, dataFinal])

        # Incrementa para o próximo dia
        current_date += timedelta(days=1)


# Janela principal
root = ctk.CTk()
root.geometry("500x500")
root.iconbitmap('jfa.ico')
root.title("Seletor de Datas")

# Configuração do estilo para aumentar a fonte do DateEntry
style = ttk.Style(root)


# Configuração para ocupar a janela toda
root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure((0, 1, 2, 3, 4), weight=1)

data_atual = datetime.now()

# Label e DateEntry para a data inicial
textoInicial = ctk.CTkLabel(root, text="Data Inicial:", font=("Arial", 16))
textoInicial.grid(row=0, column=0, padx=10, pady=5, sticky="nsew")

start_cal = DateEntry(root, width=30, background='darkblue', foreground='white', borderwidth=2, year=2024 , locale='pt_BR', day=data_atual.day - 1)
start_cal.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")

# Label e DateEntry para a data final
textoFinal = ctk.CTkLabel(root, text="Data Final:", font=("Arial", 16))
textoFinal.grid(row=2, column=0, padx=10, pady=5, sticky="nsew")

end_cal = DateEntry(root, width=30, background='darkblue', foreground='white', borderwidth=2, year=2024 , locale='pt_BR', day=data_atual.day - 1)
end_cal.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")

# Botão para confirmar seleção
confirm_button = ctk.CTkButton(root, text="Confirmar Data", font=("Arial", 16), command=on_date_selected)
confirm_button.grid(row=4, column=0, padx=10, pady=10, sticky="nsew")

# Executar a aplicação
root.mainloop()
