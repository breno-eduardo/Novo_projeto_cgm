import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess

# Verifica se o programa está rodando como um .exe ou como um script Python
if getattr(sys, 'frozen', False):
    dir_atual = sys._MEIPASS  # Diretório temporário do PyInstaller
else:
    dir_atual = os.path.dirname(os.path.abspath(__file__))

# Caminhos corretos para os scripts dentro do .exe ou na pasta do projeto
caminho_api = os.path.join(dir_atual, "chamada_APInova.py")
caminho_processamento = os.path.join(dir_atual, "D1_convertido_otimizado.py")

# Variável global para armazenar o caminho do arquivo Excel
arquivo_excel = ""
arquivo_processado = ""

# Função para selecionar o arquivo
def selecionar_arquivo():
    global arquivo_excel
    arquivo_excel = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx;*.xlsm")])

    if arquivo_excel:
        lbl_status.config(text=f"📄 Arquivo Selecionado:\n{os.path.basename(arquivo_excel)}")
    else:
        lbl_status.config(text="🔹 Nenhum arquivo selecionado!")

# Função para executar um script sem abrir novas janelas
def executar_script(caminho_script, argumentos=[]):
    try:
        if not os.path.exists(caminho_script):
            print(f"❌ ERRO: Arquivo não encontrado -> {caminho_script}")
            return f"Erro: Arquivo não encontrado -> {caminho_script}"
        
        resultado = subprocess.run(
            [sys.executable, caminho_script] + argumentos,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        return resultado.stdout.strip()
    
    except subprocess.CalledProcessError as e:
        return f"Erro: {e.stderr.strip()}"

# Função para processar o arquivo selecionado
def processar_arquivo():
    global arquivo_excel, arquivo_processado

    if not arquivo_excel:
        messagebox.showwarning("Aviso", "Por favor, selecione um arquivo XLSX ou XLSM primeiro!")
        return

    try:
        # Chamar API e gerar CSV/XLSX automaticamente
        lbl_status.config(text="🔄 Chamando API...")
        root.update()
        resposta_api = executar_script(caminho_api)

        # Processar o arquivo Excel selecionado
        lbl_status.config(text="⚙️ Processando Arquivo...")
        root.update()
        resposta_processamento = executar_script(caminho_processamento, [os.path.abspath(arquivo_excel)])

        # Verifica se o script retornou o nome do arquivo processado
        if "Arquivo processado salvo em:" in resposta_processamento:
            arquivo_processado = resposta_processamento.split("Arquivo processado salvo em:")[1].strip()
        else:
            arquivo_processado = "Arquivo processado (nome desconhecido)"

        messagebox.showinfo("Sucesso", f"Processamento concluído com sucesso!\n📄 {arquivo_processado}")
        lbl_status.config(text=f"✅ Processamento Finalizado!\n📄 {arquivo_processado}")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro durante o processamento:\n{e}")
        lbl_status.config(text="❌ Erro no processamento!")

# Criar interface
root = tk.Tk()
root.title("Processador de Arquivos Excel")
root.geometry("400x300")

# Botão para selecionar o arquivo
btn_selecionar = tk.Button(root, text="Selecionar Arquivo Excel", command=selecionar_arquivo, width=30, height=2, bg="blue", fg="white")
btn_selecionar.pack(pady=10)

# Botão para processar o arquivo selecionado
btn_processar = tk.Button(root, text="Processar Arquivo Selecionado", command=processar_arquivo, width=30, height=2, bg="green", fg="white")
btn_processar.pack(pady=10)

# Rótulo de status
lbl_status = tk.Label(root, text="🔹 Selecione um arquivo para começar", wraplength=350, justify="center")
lbl_status.pack(pady=10)

# Rodar interface
root.mainloop()
