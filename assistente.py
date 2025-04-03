import time
import os
import pandas as pd
import pyautogui
import webbrowser
from openpyxl import Workbook

#Autor(a): Emilly Lourenço da Silva
#Este código é uma automação para o envio de mensagens. Ele solicita ao usuário o número do destinatário, abre o WhatsApp Web, seleciona o campo de mensagem, digita o conteúdo e pressiona a tecla Enter para enviá-la.


wb = Workbook()
ws = wb.active
ws.title = "Relatório de Execução"
ws.append(["Tarefa", "Status", "Horário", "Tempo_Execução"])  

def registrar_tarefa_excel(tarefa, status, tempo_execucao):
    horario = time.strftime("%H:%M:%S")
    ws.append([tarefa, status, horario, round(tempo_execucao, 2)])  
    print(f"> {tarefa}: {status} (Tempo de Execução: {tempo_execucao:.2f}s)")

def abrir_whatsapp(contato):
    inicio = time.time()
    url_whatsapp = f"https://web.whatsapp.com/send?phone=55{contato}"
    webbrowser.open(url_whatsapp)
    time.sleep(10)  
    fim = time.time()
    registrar_tarefa_excel("Abrir WhatsApp", "Executado", fim - inicio)

def esperar(segundos):
    inicio = time.time()
    time.sleep(int(segundos))
    fim = time.time()
    registrar_tarefa_excel(f"Esperar {segundos} segundos", "Executado", fim - inicio)

def clicar_na_tela(x, y):
    inicio = time.time()
    pyautogui.click(x, y)
    fim = time.time()
    registrar_tarefa_excel(f"Clique em ({x}, {y})", "Executado", fim - inicio)

def digitar_texto(texto):
    inicio = time.time()
    pyautogui.write(texto)
    fim = time.time()
    registrar_tarefa_excel("Digitar Mensagem", "Executado", fim - inicio)

def pressionar_tecla(tecla):
    inicio = time.time()
    pyautogui.press(tecla)
    fim = time.time()
    registrar_tarefa_excel(f"Pressionar {tecla}", "Executado", fim - inicio)

def executar_tarefas(csv_arquivo, contato):
    tarefas = pd.read_csv(csv_arquivo)

    for _, linha in tarefas.iterrows():
        if linha["Tipo"] == "abrir":
            abrir_whatsapp(contato)
        elif linha["Tipo"] == "espera":
            esperar(linha["Dado"])
        elif linha["Tipo"] == "digitar":
            digitar_texto(linha["Dado"])
        elif linha["Tipo"] == "tecla":
            pressionar_tecla(linha["Dado"])
        elif linha["Tipo"] == "clique":
            x, y = map(int, linha["Dado"].strip('"').split(","))
            clicar_na_tela(x, y)


def salvar_relatorio():
    nome_relatorio = f"relatorio_execucao_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
    relatorios = "relatorio_execucao"
    salvar_rela = os.path.join(relatorios, nome_relatorio)  
    wb.save(salvar_rela)
    print(f"\n >Tarefa executada com sucesso! Relatório salvo em: {salvar_rela}")

if __name__ == "__main__":
    print("\n------------------------- Boot inicial de mensagens -------------------------")
    contato = input("Digite o número da pessoa que quer mandar uma mensagem (somente números e tudo junto, sem +55): ").strip()
    
    executar_tarefas("tarefas.csv", contato)
    salvar_relatorio()
