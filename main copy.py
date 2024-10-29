import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from dotenv import load_dotenv
import pandas as pd
import requests
import json
import os
import numpy as np
import logging

# Configurando o logger
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', filename='lead_update_log.log', filemode='w')

# Carregando as variáveis de ambiente
load_dotenv()

# Função para carregar a planilha de RFs
def carregar_planilha_rf():
    global df_rf
    caminho_arquivo = r"C:\0\EXACT_SALES\rf.xlsx"
    
    try:
        df_rf = pd.read_excel(caminho_arquivo)
        df_rf['RF'] = df_rf['RF'].apply(lambda x: str(x))  # Converte RF para string
        df_rf['RD'] = df_rf['RD'].apply(lambda x: str(x))  # Converte RD para string
        messagebox.showinfo("Sucesso", "Planilha de RFs carregada com sucesso!")
        logging.info("Planilha de RFs carregada com sucesso.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar a planilha: {e}")
        logging.error(f"Erro ao carregar a planilha: {e}")

# Função para consultar dados dos leads na API
def listar_leads():
    api_base_url = "https://api.exactspotter.com/v3/Leads"
    headers = {
        "token_exact": os.getenv("TOKEN_EXACT"),
        "Content-Type": "application/json"
    }
    
    response = requests.get(api_base_url, headers=headers, timeout=30)
    
    if response.status_code == 200:
        try:
            leads_data = response.json()
            if isinstance(leads_data, dict) and 'value' in leads_data:
                leads_data = leads_data['value']

            global leads_list
            leads_list = leads_data
            messagebox.showinfo("Sucesso", "Leads consultados com sucesso!")
            logging.info("Leads consultados com sucesso.")

            # Exibir leads na tela
            leads_text = "\n".join([f"ID: {lead.get('id')}, Nome: {lead.get('lead', 'N/A')}, Município: {lead.get('city', 'N/A')}" for lead in leads_list])
            text_area.delete(1.0, tk.END)
            text_area.insert(tk.END, leads_text)

            # Salvar leads em uma planilha
            df_leads = pd.DataFrame(leads_list)
            caminho_saida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
            if caminho_saida:
                df_leads.to_excel(caminho_saida, index=False)
                messagebox.showinfo("Sucesso", f"Leads salvos em {caminho_saida}")
                logging.info(f"Leads salvos em {caminho_saida}")
        except json.JSONDecodeError:
            messagebox.showerror("Erro", "A resposta da API não está no formato JSON esperado.")
            logging.error("A resposta da API não está no formato JSON esperado.")
    else:
        messagebox.showerror("Erro", f"Erro ao consultar leads: {response.status_code} - {response.text}")
        logging.error(f"Erro ao consultar leads: {response.status_code} - {response.text}")

# Função para atualizar o campo personalizado "Região" com base no município do lead
def atualizar_regiao():
    headers = {
        "token_exact": os.getenv("TOKEN_EXACT"),
        "Content-Type": "application/json"
    }
    
    if 'leads_list' not in globals():
        messagebox.showerror("Erro", "Por favor, consulte os leads primeiro.")
        logging.error("Tentativa de atualizar leads sem consultar a lista de leads primeiro.")
        return

    for lead in leads_list:
        lead_id = lead.get('id')
        lead_city = lead.get('city')
        lead_name = lead.get('lead', 'N/A')

        if lead_city is None:
            logging.warning(f"Lead ID {lead_id} não possui cidade definida. Ignorando.")
            continue

        # Verifica se a cidade do lead está presente na planilha de RFs
        if 'df_rf' in globals() and lead_city in df_rf['MUNICIPIO'].values:
            regiao_value = df_rf[df_rf['MUNICIPIO'] == lead_city]['RF'].values[0]
            regiao_rd = df_rf[df_rf['MUNICIPIO'] == lead_city]['RD'].values[0]

            # Configura a atualização do campo "Região" com base no valor correspondente
            update_data = {
                "duplicityValidation": "false",
                "lead": {
                    "customFields": [
                        {
                            "id": "_regiao",
                            "options": [
                                {
                                    "id": int(regiao_value),
                                    "value": regiao_rd
                                }
                            ]
                        }
                    ]
                }
            }
            try:
                api_update_url = f"https://api.exactspotter.com/v3/LeadsUpdate/{lead_id}"
                update_response = requests.put(api_update_url, headers=headers, json=update_data, timeout=30)
                
                if update_response.status_code == 201:
                    print(f"Lead atualizado com sucesso! ID: {lead_id}, Nome: {lead_name}, Município: {lead_city}, Região: {regiao_value}")
                    logging.info(f"Lead atualizado com sucesso! ID: {lead_id}, Nome: {lead_name}, Município: {lead_city}, Região: {regiao_value}")
                else:
                    print(f"Erro ao atualizar lead {lead_id}: {update_response.status_code} - {update_response.text}")
                    print("Payload enviado:", json.dumps(update_data, indent=4, ensure_ascii=False))
                    logging.error(f"Erro ao atualizar lead {lead_id}: {update_response.status_code} - {update_response.text}")
                    logging.debug(f"Payload enviado: {json.dumps(update_data, indent=4, ensure_ascii=False)}")
            except requests.RequestException as e:
                print(f"Erro ao tentar atualizar lead {lead_id}: {e}")
                logging.error(f"Erro ao tentar atualizar lead {lead_id}: {e}")
        else:
            print(f"Região não encontrada para o município: {lead_city}")
            logging.warning(f"Região não encontrada para o município: {lead_city}")

# TKINTER
janela = tk.Tk()
janela.title("Atualização de Leads - ExactSales")
janela.geometry("600x600")
janela.configure(bg="#ffffff")

# Botão para carregar a planilha de RFs
btn_carregar_rf = tk.Button(janela, text="Carregar Planilha de RFs", command=carregar_planilha_rf)
btn_carregar_rf.pack(pady=10)

# Botão para listar leads
btn_listar_leads = tk.Button(janela, text="Listar Leads", command=listar_leads)
btn_listar_leads.pack(pady=10)

# Área de texto para exibir os leads
text_area = scrolledtext.ScrolledText(janela, width=70, height=20)
text_area.pack(pady=10)

# Botão para atualizar a região dos leads
btn_atualizar_regiao = tk.Button(janela, text="Atualizar Região dos Leads", command=atualizar_regiao)
btn_atualizar_regiao.pack(pady=10)

# Inicia a interface gráfica do Tkinter
janela.mainloop()
