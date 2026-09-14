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
ABA_PLANILHA = "Pedido"      # nome da aba
COLUNA_CODIGO = 1            # Coluna Produto
COLUNA_QTD = 2               # Coluna Quantidade
COLUNA_PRECO = 3             # Coluna Custo
COLUNA_FORNECEDOR = 4        # Coluna fornecedor
COLUNA_EMPRESA = 5           # Coluna empresa
COLUNA_PRAZO = 6             # Coluna prazo
TEMPO_PREPARO = 5            # tempo para se posicionar
DELAY_CARREGAMENTO = 0.2     # tempo de espera após digitar código
DELAY_ENTRE_ITENS = 0.3      # tempo entre linhas
PAUSA_HUMANA = (0, 0)        # pausas aleatórias
parar = False


# ===========================
# FUNÇÕES DE AUTOMAÇÃO
# ===========================
def monitorar_esc():
    """Monitora a tecla ESC globalmente para parar a automação."""
    global parar
    keyboard.wait('esc')
    parar = True


def obter_planilha_aberta():
    """Obtém a planilha ativa do Excel."""
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        workbook = excel.ActiveWorkbook
        if not workbook:
            messagebox.showerror("Erro", "Nenhum arquivo Excel aberto foi encontrado.")
            return None
        try:
            planilha = workbook.Sheets(ABA_PLANILHA)
        except:
            planilha = workbook.ActiveSheet
        return planilha
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível acessar o Excel.\n\nDetalhes: {e}")
        return None


def digitar_lento(texto):
    """Simula digitação humana com pausas."""
    for char in str(texto):
        pyautogui.write(char)
        time.sleep(random.uniform(*PAUSA_HUMANA))


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

        # controle se já estamos dentro de um pedido aberto no sistema
        pedido_ativo = False

        # percorre TODAS as linhas com dados da planilha (a partir da linha 2 pra ignorar cabeçalho de tabela)
        for i in range(2, linhas + 1):

            # Se o usuário apertou ESC, aborta
            if parar:
                log_text.insert(tk.END, "\n⛔ Automação interrompida (ESC pressionado)\n")
                log_text.yview(tk.END)
                break

            # Se a linha estiver oculta por filtro no Excel, pula
            if planilha.Rows(i).Hidden:
                continue

            # lê os valores dessa linha
            codigo      = planilha.Cells(i, COLUNA_CODIGO).Value
            qtd         = planilha.Cells(i, COLUNA_QTD).Value
            custo       = planilha.Cells(i, COLUNA_PRECO).Value
            fornecedor  = planilha.Cells(i, COLUNA_FORNECEDOR).Value
            empresa     = planilha.Cells(i, COLUNA_EMPRESA).Value
            prazo       = planilha.Cells(i, COLUNA_PRAZO).Value

            # detecta se ESSA linha é o INÍCIO DE UM NOVO PEDIDO
            tem_fornecedor = fornecedor not in (None, "", " ")
            tem_empresa    = empresa    not in (None, "", " ")
            tem_prazo      = prazo      not in (None, "", " ")

            if tem_fornecedor and tem_empresa and tem_prazo:
                # tentar interpretar empresa como número (quantas vezes descer com seta ↓)
                empresa_str = str(empresa).strip().replace(","," ")
                try:
                    empresa_pos = int(float(empresa_str))
                except (ValueError, TypeError):
                    log_text.insert(tk.END, f"⚠️ Linhas {i} ignorada como cabeçalho: Empresa inválida ({empresa})\n")
                    log_text.yview(tk.END)
                    continue

                # normaliza fornecedor e prazo pra strings limpas
                fornecedor_txt = (
                    str(int(fornecedor)).strip()
                    if isinstance(fornecedor, float)
                    else str(fornecedor).strip()
                )

                if isinstance(prazo, float):
                    prazo_txt = str(int(prazo)).strip()
                else:
                    prazo_txt = str(prazo).strip() if prazo is not None else ""

                log_text.insert(tk.END, f"\n📦 Novo pedido: Fornecedor {fornecedor_txt} | Empresa {empresa_pos} | Prazo {prazo_txt}\n")
                log_text.yview(tk.END)

                # === DIGITA O CABEÇALHO NO SISTEMA ===
                # foco deve estar no campo "Fornecedor" do sistema ANTES da automação começar

                # digita fornecedor
                digitar_lento(fornecedor_txt)

                # vai pro campo "Empresa"
                pyautogui.press('tab')

                # escolhe empresa apertando seta pra baixo N vezes
                for _ in range(empresa_pos):
                    pyautogui.press('down')
                    time.sleep(0.05)

                # agora navega até o campo "Prazo"
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')

                # digita prazo
                digitar_lento(prazo_txt)

                # depois navega até a grade de itens (seu monte de TABs)
                for _ in range(21):
                    pyautogui.press('tab')

                # confirma entrada na grade
                pyautogui.press('enter')

                pedido_ativo = True
                # essa linha era só cabeçalho -> não tenta tratar como item
                continue

            # Se NÃO era cabeçalho: agora pode ser linha de ITEM ou linha de FIM DE PEDIDO

            # Se não tem código de produto, isso significa que terminou esse pedido
            if codigo in (None, "", " "):
                if pedido_ativo:
                    pyautogui.press('f3')  # emite pedido
                    log_text.insert(tk.END, "🛒 Pedido emitido (F3)\n")
                    log_text.yview(tk.END)
                    pedido_ativo = False
                # e segue pra próxima linha pra talvez abrir outro pedido abaixo
                continue

            # A partir daqui, é uma linha de ITEM do pedido
            # Normaliza os dados pra digitação
            codigo_txt = (
                str(int(codigo)).strip()
                if isinstance(codigo, float)
                else str(codigo).strip()
            )

            qtd_txt = (
                str(int(qtd)).strip()
                if isinstance(qtd, float)
                else str(qtd).replace(".0", "").strip()
                if qtd is not None else ""
            )

            custo_txt = (
                str(custo).replace(".", ",").strip()
                if custo is not None else ""
            )

            # === DIGITA ITEM NO GRID ===
            # sequência: Código → Tab → Tab → Qtde → Tab → (custo se marcado) → Tab → Tab → Insert

            import pygetwindow as gw

            # digita código
            digitar_lento(codigo_txt)

            # depois do código: primeiro TAB (vai pra próximo campo)
            pyautogui.press('tab')

            # se aparecer alerta tipo "Atenção", paramos tudo
            janelas = gw.getAllTitles()
            if any("Atenção" in j for j in janelas):
                log_text.insert(tk.END, f"\n⛔ Automação interrompida (Alerta na Consinco após código {codigo_txt})\n")
                log_text.yview(tk.END)
                break

            # segundo TAB (campo quantidade)
            pyautogui.press('tab')
            time.sleep(DELAY_CARREGAMENTO)

            # digita quantidade
            digitar_lento(qtd_txt)

            # checa alerta depois de qtd
            janelas = gw.getAllTitles()
            if any("Atenção" in j for j in janelas):
                log_text.insert(tk.END, f"\n⛔ Automação interrompida (Alerta na Consinco após QTD do código {codigo_txt})\n")
                log_text.yview(tk.END)
                break

            # sai do campo quantidade
            pyautogui.press('tab')
            time.sleep(DELAY_CARREGAMENTO)

            # se a opção "Digitar custo" estiver marcada na interface, digita o custo
            if incluir_preco_var.get() and custo_txt != "":
                digitar_lento(custo_txt)

                # checa alerta depois do custo
                janelas = gw.getAllTitles()
                if any("Atenção" in j for j in janelas):
                    log_text.insert(tk.END, f"\n⛔ Automação interrompida (Alerta na Consinco após custo do código {codigo_txt})\n")
                    log_text.yview(tk.END)
                    break

            # avança mais duas vezes tab (ajusta a posição conforme seu sistema)
            pyautogui.press('tab')
            pyautogui.press('tab')
            time.sleep(0.2)

            # cria nova linha de item dentro do mesmo pedido
            pyautogui.press('insert')
            

            # log
            time.sleep(DELAY_ENTRE_ITENS)
            log_text.insert(
                tk.END,
                f"Linha {i}: Código {codigo_txt} | QTD {qtd_txt} | Custo {custo_txt}\n"
            )
            log_text.yview(tk.END)

        # acabou de percorrer a planilha:
        # se terminamos ainda com um pedido ativo (não teve linha vazia pra fechar), fecha agora
        if pedido_ativo:
            pyautogui.press('f3')
            log_text.insert(tk.END, "🛒 Pedido emitido no final da planilha.\n")
            log_text.yview(tk.END)

        log_text.insert(tk.END, "\n⏳ Automação finalizada.\n")
        log_text.yview(tk.END)

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro durante a automação:\n{e}")
    finally:
        pythoncom.CoUninitialize()
        botao_iniciar.config(state=tk.NORMAL)


def iniciar_automacao():
    """Inicia automação em uma thread separada."""
    botao_iniciar.config(state=tk.DISABLED)
    threading.Thread(target=monitorar_esc, daemon=True).start()
    threading.Thread(target=executar_automacao, daemon=True).start()


# ===========================
# INTERFACE GRÁFICA
# ===========================
janela = tk.Tk()
janela.title("Automação de pedidos")
janela.geometry("600x400")

tk.Label(janela, text="Para iniciar clique no botão de iniciar automação", font=("Arial", 12, "bold")).pack(pady=10)
botao_iniciar = tk.Button(janela, text="Iniciar Automação", bg="#4CAF50", fg="white", width=20, command=iniciar_automacao)
botao_iniciar.pack(pady=10)

incluir_preco_var = tk.BooleanVar() 
tk.Checkbutton(janela, text="Digitar custo", variable=incluir_preco_var).pack()

log_text = scrolledtext.ScrolledText(janela, width=70, height=15)
log_text.pack(padx=10, pady=10)

tk.Label(janela, text="Pressione ESC para parar a automação a qualquer momento.", fg="gray").pack()

janela.mainloop()