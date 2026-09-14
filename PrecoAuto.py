import tkinter as tk
from tkinter import messagebox, scrolledtext
import pyautogui
import win32com.client
import pythoncom
import keyboard
import time
import threading
import random

# ===========================
# CONFIGURAÇÕES
# ===========================

ABA_PLANILHA = "Planilha1"   # nome da aba
COLUNA_CODIGO = 1            # Coluna Produto
COLUNA_PRECO = 3             # Coluna Custo
TEMPO_PREPARO = 5            # tempo para se posicionar
DELAY_CARREGAMENTO = 0.2     # tempo de espera após digitar código
DELAY_ENTRE_ITENS = 0.4      # tempo entre linhas
PAUSA_HUMANA = (0, 0)        # pausas aleatórias
parar = False


# ===========================
# FUNÇÕES DE AUTOMAÇÃO
# ===========================

def obter_planilha_aberta():
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        workbook = excel.ActiveWorkbook
        if not workbook:
            messagebox.showerror("Erro", "Nenhum arquivo arquivo encontrado.")
            return None
        try:
            planilha = workbook.Sheets(ABA_PLANILHA)
        except:
            planilha = workbook.ActiveSheet
        return planilha
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível acessar o Excel.\n\nDetalhes: {e}")
        return None
    
def digitar(texto):
    for char in str(texto):
        pyautogui.write(char)
        
def executar_automacao():
    global parar
    parar = False
    pythoncom.CoInitialize()
    
    try:
        planilha = obter_planilha_aberta()
        if planilha is None:
            return
        
        # Contagem regressiva / preparo
        linhas = planilha.UsedRange.Rows.Count
        log_text.insert(tk.END, f"Planilha detectada com {linhas - 1} linhas de dados.\n")
        log_text.insert(tk.END, f"Você tem {TEMPO_PREPARO} segundos para se posicionar no campo CÓDIGO do sistema...\n")
        log_text.yview(tk.END)
        
        for t in range(TEMPO_PREPARO, 0, -1):
            log_text.insert(tk.END, f"Iniciando em {t} segundos...\r")
            log_text.yview(tk.END)
            time.sleep(1)

        log_text.insert(tk.END, "\n🚀 Iniciando automação!\n\n")
        log_text.yview(tk.END)
        
        for i in range(2, linhas + 1):
            if planilha.Rows(i).Hidden:
                continue
            
            codigo = planilha.Cells(i, COLUNA_CODIGO).Value
            custo = planilha.Cells(i, COLUNA_PRECO).Value
            
            codigo_txt = (
                str(int(codigo)).strip()
                if isinstance(codigo, float)
                else str(codigo).strip()
            )

            custo_txt = (
                str(custo).replace(".", ",").strip()
                if custo is not None else ""
            )

            import pygetwindow as gw
            
            digitar(codigo_txt)
            pyautogui.press('tab')
            digitar(custo_txt)
            pyautogui.press('tab')
            pyautogui.press('insert')
            time.sleep(DELAY_ENTRE_ITENS)
            
    except Exception as e:
                messagebox.showerror("Erro", f"Ocorreu um erro durante a automação:\n{e}")
    finally:
        pythoncom.CoUninitialize()
        botao_iniciar.config(state=tk.NORMAL)
        
def iniciar_automacao():
    botao_iniciar.config(state=tk.DISABLED)
    threading.Thread(target=executar_automacao, daemon=True).start()
    
# ===========================
# INTERFACE GRÁFICA
# ===========================
janela = tk.Tk()
janela.title("Automação de preços")
janela.geometry("600x400")
bg = "gray27"

tk.Label(janela, text="Para iniciar clique no botão de iniciar automação", font=("Arial", 12, "bold")).pack(pady=10)
botao_iniciar = tk.Button(janela, text="Iniciar Automação", bg="#4CAF50", fg="white", width=20, command=iniciar_automacao)
botao_iniciar.pack(pady=10)

log_text = scrolledtext.ScrolledText(janela, width=70, height=15)
log_text.pack(padx=10, pady=10)

janela.mainloop()