import requests
import json
import pandas as pd
import tkinter as tk
from tkinter import ttk, scrolledtext
from senha import API_KEY

def analisar_dados():
    
    analisar_button.config(state=tk.DISABLED)

    
    progresso_bar['value'] = 0
    root.update_idletasks()

    
    df = pd.read_excel('seuarquivo.xlsx')

    
    data_text = df.to_json(orient='records')

    
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }

    link = "https://api.openai.com/v1/chat/completions"

    id_modelo = "gpt-3.5-turbo"

    
    mensagem_analise = {
        "role": "user",
        "content": "Realize uma análise curta dos dados e forneça recomendações de dashboards com base nos seguintes dados:\n" + data_text
    }

    body_mensagem = {
        "model": id_modelo,
        "messages": [mensagem_analise]
    }

    body_mensagem = json.dumps(body_mensagem)

   
    requisicao = requests.post(link, headers=headers, data=body_mensagem)
    resposta_json = requisicao.json()

    
    mensagem_gerada = resposta_json['choices'][0]['message']['content']

    
    resultado_text.config(state=tk.NORMAL)
    resultado_text.delete(1.0, tk.END)
    resultado_text.insert(tk.END, mensagem_gerada)

    
    resultado_text.tag_add("font_size", 1.0, "end")
    resultado_text.tag_config("font_size", font=("Helvetica", 12))

    resultado_text.config(state=tk.DISABLED)

   
    analisar_button.config(state=tk.NORMAL)


root = tk.Tk()
root.title("Análise e Recomendações de Dashboards")


style = ttk.Style()
style.configure('TButton', font=('Helvetica', 12))
style.configure('TLabel', font=('Helvetica', 14))


root.geometry("600x400")


resultado_text = scrolledtext.ScrolledText(root, width=40, height=10, state=tk.DISABLED)
resultado_text.grid(row=0, column=0, padx=10, pady=10, rowspan=2, sticky='nsew')


analisar_button = ttk.Button(root, text="Analisar Dados", command=analisar_dados)
analisar_button.grid(row=2, column=0, padx=10, pady=10, sticky='nsew')


progresso_bar = ttk.Progressbar(root, orient="horizontal", mode="determinate")
progresso_bar.grid(row=3, column=0, padx=10, pady=10, sticky='nsew')

root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

root.mainloop()
