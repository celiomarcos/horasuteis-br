import csv
import datetime
from datetime import date
import businesstimedelta
import holidays

import tkinter as tk
from tkinter import TclError, filedialog, simpledialog, messagebox

# tela usada na gui, inicia em None, pq precisa ser criada no ambiente
win = None
try:
    win = tk.Tk()
except TclError:
    # se tentou criar, e deu problema (nao tem tela por ex), vai ficar win valendo None, pra saber q nao pode usar win
    print('nao vai rolar tela, usar nome fixo de arquivo local')
    # usar um nome de arquivo local
    path_selecionado_pela_gui = 'dataspracomparar.csv'
else:
    win.withdraw()
    path_selecionado_pela_gui = filedialog.askopenfilename(filetypes=[("Formato de planilhas", ".csv .xls .xlsx")])

d1= datetime.datetime.now()
arquivo_nome_com_data = d1.strftime("%Y%m%d_%H%M%S")

# definir um dia de trawbalho
diadetrabalho = businesstimedelta.WorkDayRule(
    start_time=datetime.time(8),
    end_time=datetime.time(18),
    working_days=[0, 1, 2, 3, 4])

# horario de almoco
lunchbreak = businesstimedelta.LunchTimeRule(
    start_time=datetime.time(12),
    end_time=datetime.time(13),
    working_days=[0, 1, 2, 3, 4])

# combinar os dois
#horas_uteis = businesstimedelta.Rules([diadetrabalho, lunchbreak])
horas_uteis = businesstimedelta.Rules([diadetrabalho])

#cabecalho do arquivo de entrada
cabecalho_esperado = ['INICIO', 'FINAL', 'TEMPO', 'UF']
cabecalho_output = cabecalho_esperado + ['TIME_REAL', 'TIME_OK']

quantidade_de_registros_gravados = 0

with open(path_selecionado_pela_gui,'r') as data_input:
    # alguns testes basicos com o arquivo de entrada

    #para isso le os primeiros 1024 bytes, aumentar se achar q n for suficiente, mas geralmente eh
    inicio_do_arquivo = data_input.read(1024)
    
    tem_cabecalho = csv.Sniffer().has_header(inicio_do_arquivo)
    if not tem_cabecalho:
        _mensagem_de_erro = "Sem cabecalho no arquivo, tem q ter um assim:\n{}".format(cabecalho_esperado)
        if win is None:
            print(_mensagem_de_erro)
        else:
            messagebox.showerror("CSV inválido", _mensagem_de_erro)
        exit

    # analisar se descobre o separador
    dialect = csv.Sniffer().sniff(inicio_do_arquivo)
    if dialect is None or dialect == '':
        # usa um separador padrao
        separador = ','

        #ou poderia perguntar qual separador usar, dai fica por conta do usuario ter certeza disso
        if win is None:
            separador = input("qual separador usar: ")
        else:
            separador = simpledialog.askstring("Questão", "Qual separador usar?")
        
        # reverifdica pq esse if de cima pode ser cancelado pelo usuario
        if separador is None or separador == '':
            separador = ','

    else:
        separador = dialect.delimiter

    #volta pro inicio do arquivo
    data_input.seek(0)

    # agora assim abrir o arquvo para ler as linhas
    reader = csv.DictReader(data_input, delimiter=separador)

    #get fieldnames from DictReader object and store in list
    headers = reader.fieldnames
    print("Cabecalhos encontrados:" + str(headers))

    #if row != cabecalho_esperado:
    #    if win is None:
    #        print("Ops! Arquivo inválido! ☹")
    #    else:
    #        messagebox.showerror("Cabeçalho Inválido ☹", "Use estes separados por ponto e vírgula: {}".format(validado))
    #    #de qquer forma encerra
    #    exit
        
    #mudar essa validacao simples para

    #headers = []
    #for row in reader:
    #    headers = [x.lower() for x in list(row.keys())]
    #    break

    #if 'minha ccoluna' not in headers or 'id_nome' not in headers:
    #    print('Arquivo CSV precisa ter as colunas "Minha Coluna" e a coluna "ID_Nome"')

    # confirma onde salvar o arquivo destino
    if win is None:
        nome_arquivo_saida = f'{arquivo_nome_com_data}.csv'
    else:
        f = filedialog.asksaveasfile(initialfile = f'{arquivo_nome_com_data}.csv', defaultextension=".csv",filetypes=[("Planilha csv","*.csv"),("Documentos texto","*.txt")])
        if f is None:
            # provavel q cancelou a selecao de arquivo, encerrar tudo entao
            exit
        # extrair o filenamen do objeto retornado pela tela
        nome_arquivo_saida = f.name

    #abre arquivo de saida
    with open(nome_arquivo_saida, 'w') as arquivo_output:
        writer = csv.DictWriter(arquivo_output, cabecalho_output, lineterminator='\n')
        
        #escreve cabecalho novo
        writer.writeheader()
        
        # variavel que contera todas as linhas do arquivo original, eh incrementada 
        # enquanto fica lendo no loop abaixo, e será gravada de uma só vez no final 
        all_rows = []

        for row in reader:
            row_saida = {}
            
            inicio = datetime.datetime.strptime(row['INICIO'], '%d/%m/%Y %H:%M')
            end = datetime.datetime.strptime(row['FINAL'], '%d/%m/%Y %H:%M')
        
            # adiciona na variavel da saida
            #row.append(businesshrs.difference(inicio, end))
            bdiff = horas_uteis.difference(inicio, end)
            row['TIME_REAL'] = "{}:{}:00".format(bdiff.hours, f"{int(bdiff.seconds/60):02d}")

            estado = row['UF']
            if estado == "":
                estado = 'SP'
            feriados_uf = holidays.BR(state=estado)
            feriados = businesstimedelta.HolidayRule(feriados_uf)

            #horas_uteis = businesstimedelta.Rules([diadetrabalho, lunchbreak, feriados])
            businesshrs = businesstimedelta.Rules([diadetrabalho, feriados])
            bdiff = businesshrs.difference(inicio, end)
            row['TIME_OK'] = "{}:{}:00".format(bdiff.hours, f"{int(bdiff.seconds/60):02d}")

            all_rows.append(row)

        #escreve toda variavel par3a arquivo
        writer.writerows(all_rows)
        quantidade_de_registros_gravados = len(all_rows)

#finaliza com alguma msg pro usuario
msg = 'Processou {} registros'.format(quantidade_de_registros_gravados)
if writer is None:
    print(msg)
else:
    messagebox.showinfo("Encerrou", msg)
