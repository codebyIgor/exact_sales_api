import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import requests
import json
import os

# Função para carregar a planilha de RFs 
# Esta função abre uma janela para o usuário selecionar um arquivo Excel, 
# e armazena os dados em um DataFrame do Pandas para uso posterior.
def carregar_planilha_rf():
    global df_rf  # Variável global para armazenar a planilha
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione a planilha de Leads",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    
    if caminho_arquivo:
        try:
            df_rf = pd.read_excel(caminho_arquivo)
            messagebox.showinfo("Sucesso", "Planilha carregada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar a planilha: {e}")

#Função para carregar planilha de Leads e ID 
def carregar_planilha_leads():
    global df_leads  # Variável global para armazenar a planilha
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione a planilha de Leads",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    
    if caminho_arquivo:
        try:
            df_leads = pd.read_excel(caminho_arquivo)
            messagebox.showinfo("Sucesso", "Planilha carregada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar a planilha: {e}")

# Função para consultar dados dos leads na API e exibir os resultados
# Esta função faz uma requisição GET à API para listar os leads e mostra os resultados em uma nova janela pop-up.
def listar_leads():
    api_base_url = "https://api.exactspotter.com/v3/Leads"
    headers = {
        "token_exact": os.getenv("TOKEN_EXACT"),
        "Content-Type": "application/json"
    }
    
    response = requests.get(api_base_url, headers=headers)
    
    if response.status_code == 200:
        try:
            leads_data = response.json()
            # Verificar se a resposta possui uma estrutura com chave 'value', comum em OData
            if isinstance(leads_data, dict) and 'value' in leads_data:
                leads_data = leads_data['value']

            leads_info = []
            for lead in leads_data:
                if isinstance(lead, dict):  # Garantir que lead é um dicionário
                    leads_info.append({
                        "Lead_ID": lead.get('id', 'N/A'),
                        "Funil_ID": lead.get('funilId', 'N/A'),
                        "Nome": lead.get('lead', 'N/A'),
                        "Cidade": lead.get('city', 'N/A'),
                        "UF": lead.get('state', 'N/A'),
                        "País": lead.get('country', 'N/A'),
                        "Etapa": lead.get('stage', 'N/A')
                    })
            
            # Mostrar os resultados na interface
            resultado_texto = ""
            for lead in leads_info:
                resultado_texto += f"Lead_ID: {lead['Lead_ID']},Funil_ID: {lead['Funil_ID']}, Nome: {lead['Nome']}, Cidade: {lead['Cidade']}, UF: {lead['UF']}, País: {lead['País']}, Etapa: {lead['Etapa']}\n"
            mostrar_resultado(resultado_texto)
            
            # Salvar os resultados em um arquivo Excel
            salvar_caminho = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
            if salvar_caminho:
                df_leads = pd.DataFrame(leads_info)
                df_leads.to_excel(salvar_caminho, index=False)
                messagebox.showinfo("Sucesso", "Leads salvos com sucesso no arquivo Excel!")
        except json.JSONDecodeError:
            messagebox.showerror("Erro", "A resposta da API não está no formato JSON esperado.")
    else:
        messagebox.showerror("Erro", f"Erro ao consultar leads: {response.status_code} - {response.text}")

#Função para importar planilha de RF 
def importar_rf(caminho_arquivo):
    if 'df_rf' not in globals():
        messagebox.showerror("Erro", "Por favor, carregue uma planilha primeiro.")
        return

        try: 
            df_rf = pd.read_excel(caminho_arquivo)

            if 'MUNICIPIO' not in df_rf.columns or 'RF' not in df_rf.columns:
                messagebox.showerror("Erro", "As colunas 'MUNICIPIO' e 'RF' não foram encontradas na planilha.")
                return

            dados_extraidos = df_rf[['MUNICIPIO', 'RF']]

            return dados_extraidos
        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo não encontrado. Por favor, verifique o camingo do arquivo")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao importar a planilha: {str(e)}")


# Função para enviar os dados para a API da ExactSales
# Esta função percorre as linhas da planilha carregada e envia os dados de cada lead para a API.
def enviar_dados():
    if 'df_leads' not in globals():
        messagebox.showerror("Erro", "Por favor, carregue uma planilha primeiro.")
        return

    # API URL base e Token de Autenticação (atualize conforme necessário)
    api_base_url = "https://api.exactsales.com/v1/leads/"
    headers = {
        "Authorization": "Bearer seu_token_aqui",
        "Content-Type": "application/json"
    }

    # Loop para percorrer os dados da planilha e enviar a atualização de cada lead
    for index, row in df.iterrows():
        lead_id = row['Lead ID']  # Certifique-se de que a planilha tem o ID do lead

        # Dicionário com os dados do lead que serão enviados para a API
        lead_data = {
            "name": row['Nome da Empresa'],
            "industry": row['Mercado'],
            "source": row['Origem'],
            "subSource": row['Sub-Origem'],
            "organizationId": row['Organização ID'],
            "sdrEmail": row['Email SDR'],
            "group": row['Grupo'],
            "mktLink": row['Link de Marketing'],
            "ddiPhone": row['DDI Telefone'],
            "phone": row['Telefone'],
            "ddiPhone2": row['DDI Telefone 2'],
            "phone2": row['Telefone 2'],
            "website": row['Website'],
            "leadProduct": row['Produto do Lead'],
            "address": row['Endereço'],
            "addressNumber": row['Número Endereço'],
            "addressComplement": row['Complemento Endereço'],
            "neighborhood": row['Bairro'],
            "zipcode": row['CEP'],
            "city": row['Cidade'],
            "state": row['Estado'],
            "country": row['País'],
            "cpfcnpj": row['CPF/CNPJ'],
            "description": row['Descrição'],
            "customFields": row['Campos Personalizados'],  # Pode precisar de ajustes
            "duplicityValidation": False  # Ou True, dependendo da necessidade
        }

        # Envio dos dados usando o método PUT da API
        response = requests.put(f"{api_base_url}{lead_id}", headers=headers, data=json.dumps(lead_data))
        
        # Verifica se a requisição foi bem-sucedida
        if response.status_code == 200:
            print(f"Lead {row['Nome da Empresa']} atualizado com sucesso!")
        else:
            print(f"Erro ao atualizar lead {row['Nome da Empresa']}: {response.status_code} - {response.text}")
    
    messagebox.showinfo("Finalizado", "Atualização dos leads concluída!")

# Função para exibir os resultados em uma nova janela pop-up
def mostrar_resultado(resultado):
    resultado_janela = tk.Toplevel(janela)
    resultado_janela.title("Resultado da Consulta de Leads")
    resultado_janela.geometry("600x400")
    resultado_texto = scrolledtext.ScrolledText(resultado_janela, wrap=tk.WORD, width=70, height=20)
    resultado_texto.pack(padx=10, pady=10)
    resultado_texto.insert(tk.INSERT, resultado)
    resultado_texto.config(state=tk.DISABLED)

# TKINTER 
# Configurações dos estilos de janelas pelo Tkinter
# Criação da janela principal do aplicativo usando Tkinter
janela = tk.Tk()
janela.title("Atualização de Leads - ExactSales")
janela.geometry("500x400")
janela.configure(bg="#ffffff")

# Função para mudar a aparência do botão quando o mouse está sobre ele
def on_enter(e):
    e.widget["fg"] = "blue"
    e.widget["cursor"] = "hand2"

def on_leave(e):
    e.widget["fg"] = "#000000"
    e.widget["cursor"] = "arrow"

# Estilos personalizados para os botões
estilo_botao = {
    "font": ("Arial", 12),
    "bg": "#ffffff",
    "fg": "#000000",
    "activebackground": "#d9d9d9",
    "activeforeground": "#000000",
    "relief": tk.FLAT,
    "bd": 8,
    "width": 20,
    "height": 2,
    "highlightbackground": "#000000",
    "highlightthickness": 2
}

# Estilos personalizados para as etiquetas
estilo_label = {
    "font": ("Arial", 14),
    "bg": "#ffffff",
    "fg": "#333333"
}

# Label de título do aplicativo
label_titulo = tk.Label(janela, text="Atualização de Leads", **estilo_label)
label_titulo.pack(pady=20)

# Botão para consultar leads na API
btn_consultar = tk.Button(janela, text="Consultar Leads", command=listar_leads, **estilo_botao)
btn_consultar.pack(pady=10)
btn_consultar.bind("<Enter>", on_enter)
btn_consultar.bind("<Leave>", on_leave)

# Botão para importar a planilha
btn_importar = tk.Button(janela, text="Importar Planilha", command=importar_leads, **estilo_botao)
btn_importar.pack(pady=10)
btn_importar.bind("<Enter>", on_enter)
btn_importar.bind("<Leave>", on_leave)

# Botão para enviar dados para a API
btn_enviar = tk.Button(janela, text="Enviar Atualização", command=enviar_dados, **estilo_botao)
btn_enviar.pack(pady=10)
btn_enviar.bind("<Enter>", on_enter)
btn_enviar.bind("<Leave>", on_leave)

#Botão para importar planilha de RF
btn_importar_rf = tk.Button(janela, text="Importar RF", command=importar_rf, **estilo_botao)
btn_enviar.pack(pady=10)
btn_enviar.bind("<Enter>", on_enter)
btn_enviar.bind("<Leave>", on_leave)

# Inicia a interface gráfica do Tkinter
janela.mainloop()
