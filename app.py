import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import Calendar
from datetime import datetime

def abrir_planilha():
    global caminho_planilha, clientes
    caminho_planilha = filedialog.askopenfilename(
        title="Selecione a planilha de clientes",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if caminho_planilha:
        lbl_planilha.config(text=f"Planilha selecionada: {os.path.basename(caminho_planilha)}")
        exibir_planilha()

def abrir_arquivo_excel():
    global caminho_planilha
    if caminho_planilha:
        os.startfile(caminho_planilha)
    else:
        messagebox.showerror("Erro", "Por favor, selecione uma planilha primeiro!")

def exibir_planilha():
    global clientes
    if not caminho_planilha:
        return
    
    try:
        workbook = openpyxl.load_workbook(caminho_planilha)
        pagina_clientes = workbook.active
        clientes = []

        for linha in pagina_clientes.iter_rows(min_row=2, max_col=3):
            nome = linha[0].value
            telefone = linha[1].value
            vencimento = linha[2].value
            if nome and telefone and vencimento:
                clientes.append((nome, telefone, vencimento))

        tree.delete(*tree.get_children())
        for nome, telefone, vencimento in clientes:
            tree.insert('', 'end', values=(nome, telefone, vencimento.strftime('%d/%m/%Y') if isinstance(vencimento, datetime) else vencimento))
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")

def enviar_mensagens():
    global clientes
    if not caminho_planilha:
        messagebox.showerror("Erro", "Por favor, selecione uma planilha primeiro!")
        return
    
    try:
        mensagem_personalizada = campo_mensagem.get("1.0", "end-1c").strip()
        if not mensagem_personalizada:
            messagebox.showerror("Erro", "Por favor, escreva uma mensagem antes de enviar!")
            return
        
        webbrowser.open('https://web.whatsapp.com/')
        sleep(10)  # Aguardar tempo para carregar o WhatsApp Web

        for nome, telefone, vencimento in clientes:
            try:
                mensagem = mensagem_personalizada
                link_mensagem_whatsapp = f"https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}"
                
                webbrowser.open(link_mensagem_whatsapp)
                sleep(10)  # Aguardar tempo para carregar a mensagem
                # Pressionar Enter para enviar a mensagem
                pyautogui.press('enter')
                sleep(5)  # Tempo para garantir que a mensagem seja enviada
                # Fechar a aba após enviar a mensagem
                pyautogui.hotkey('ctrl', 'w')
                sleep(2)
                # Marcar o dia no calendário
                calendario.selection_set(datetime.today().date())
            except Exception as e:
                print(f"Erro ao enviar mensagem para {nome}: {e}")
                with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                    arquivo.write(f"{nome},{telefone}{os.linesep}")
        
        messagebox.showinfo("Sucesso", "Mensagens enviadas com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

root = tk.Tk()
root.title("Envio de Mensagens Automáticas")
root.state('zoomed')  # Janela sempre maximizada

root.configure(bg="#2e2e2e")

root.option_add('*Font', 'Arial 14')
root.option_add('*Background', '#2e2e2e')
root.option_add('*Foreground', '#ffffff')

frame_principal = tk.Frame(root, bg="#2e2e2e", bd=2, relief="solid", highlightbackground="#00FF00", highlightcolor="#00FF00", highlightthickness=2)
frame_principal.pack(fill='both', expand=True, padx=10, pady=10)

frame_esquerdo = tk.Frame(frame_principal, bg="#2e2e2e", bd=2, relief="solid", highlightbackground="#00FF00", highlightcolor="#00FF00", highlightthickness=2)
frame_esquerdo.pack(side='left', fill='both', expand=True, padx=10)

frame_direito = tk.Frame(frame_principal, bg="#2e2e2e", bd=2, relief="solid", highlightbackground="#00FF00", highlightcolor="#00FF00", highlightthickness=2)
frame_direito.pack(side='right', fill='both', expand=True, padx=10)

# Adicionando o calendário e ajustando o tamanho
calendario = Calendar(frame_esquerdo, selectmode='day', font=("Arial", 12), background="#2e2e2e", foreground="#ffffff")
calendario.pack(pady=20, expand=True, fill="both")

frame_planilha = tk.Frame(frame_direito, bg="#333333", bd=2, relief="solid", highlightbackground="#00FF00", highlightcolor="#00FF00", highlightthickness=2)
frame_planilha.pack(fill="x", padx=10, pady=5)

lbl_planilha = tk.Label(frame_planilha, text="Nenhuma planilha selecionada", bg="#333333", fg="#ffffff")
lbl_planilha.pack()

tree_frame = tk.Frame(frame_direito, bg="#2e2e2e", bd=2, relief="solid", highlightbackground="#00FF00", highlightcolor="#00FF00", highlightthickness=2)
tree_frame.pack(pady=10, fill='both', expand=True)

tree = ttk.Treeview(tree_frame, columns=("Nome", "Telefone", "Vencimento"), show="headings", height=10)
tree.pack(fill='both', expand=True)

for col in tree["columns"]:
    tree.heading(col, text=col)
    tree.column(col, anchor='center')

scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
scrollbar.pack(side='right', fill='y')
tree.configure(yscrollcommand=scrollbar.set)

frame_mensagem = tk.Frame(frame_direito, bg="#333333", bd=2, relief="solid", highlightbackground="#00FF00", highlightcolor="#00FF00", highlightthickness=2)
frame_mensagem.pack(padx=10, pady=5, fill="x")

lbl_mensagem = tk.Label(frame_mensagem, text="Mensagem personalizada:", bg="#333333", fg="#ffffff")
lbl_mensagem.pack()

campo_mensagem = tk.Text(frame_mensagem, height=6, wrap="word", font=("Arial", 14))
campo_mensagem.pack(fill="x", padx=10)

btn_frame = tk.Frame(frame_direito, bg="#2e2e2e", bd=2, relief="solid", highlightbackground="#00FF00", highlightcolor="#00FF00", highlightthickness=2)
btn_frame.pack(pady=20)

btn_abrir = tk.Button(btn_frame, text="Abrir Planilha", command=abrir_planilha, bg="#444444", fg="#ffffff", font=("Arial", 14), width=20, height=2)
btn_abrir.grid(row=0, column=0, padx=10)

btn_enviar = tk.Button(btn_frame, text="Enviar Mensagens", command=enviar_mensagens, bg="#444444", fg="#ffffff", font=("Arial", 14), width=20, height=2)
btn_enviar.grid(row=0, column=1, padx=10)

# Adicionando um botão de atalho para abrir o arquivo Excel diretamente
btn_atalho = tk.Button(btn_frame, text="Arquivo Excel", command=abrir_arquivo_excel, bg="#444444", fg="#ffffff", font=("Arial", 14), width=20, height=2)
btn_atalho.grid(row=0, column=2, padx=10)

lbl_observacao = tk.Label(
    root,
    text="Certifique-se de que o WhatsApp Web esteja configurado.\nAguarde o carregamento completo antes de iniciar.",
    fg="#ff0000",
    bg="#2e2e2e",
    wraplength=650,
    font=("Arial", 14)
)
lbl_observacao.pack(pady=10)

root.mainloop()
