import csv
import numpy as np
import pandas
import datetime
from datetime import date
import businesstimedelta
import holidays

import tkinter as tk
from tkinter import TclError, filedialog, simpledialog, messagebox

# tela usada na gui, inicia em None, porque precisa ser criada no ambiente
win = None
try:
    win = tk.Tk()
except TclError:
    # se tentou criar, e deu problema (nao tem tela por ex), vai ficar win validando None, pra saber que nao pode usar win
    print('Ops... Não foi possível mostrar tela, use nome fixo de arquivo local')
    # usar um nome de arquivo local
    path_selecionado_pela_gui = 'dataspracomparar.csv'
else:
    win.withdraw()
    path_selecionado_pela_gui = filedialog.askopenfilename(filetypes=[("Formato de planilhas", ".csv .xls .xlsx")])
    #TODO se cancelar esta tela, va gerar erro grave ao tentar abrir arquivo

d1 = datetime.datetime.now()
arquivo_nome_com_data = d1.strftime("%Y%m%d_%H%M%S")

# definir um dia de e horário de trabalho
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


###############################
# carregar feriados regionais
df_regionais = pandas.read_csv("holidays_city.csv", quoting=csv.QUOTE_NONE)


#cabecalho do arquivo de entrada
cabecalho_esperado = ['INICIO', 'FINAL', 'TEMPO', 'UF']
cabecalho_output = cabecalho_esperado + ['TIME_REAL', 'TIME_OK', 'H_DECIMAL', 'CITY']

quantidade_de_registros_gravados = 0

with open(path_selecionado_pela_gui,'r') as data_input:
    # alguns testes basicos com o arquivo de entrada

    #para isso le os primeiros 1024 bytes, aumentar se achar que nao for suficiente, mas geralmente supre.
    inicio_do_arquivo = data_input.read(1024)
    
    tem_cabecalho = csv.Sniffer().has_header(inicio_do_arquivo)
    if not tem_cabecalho:
        _mensagem_de_erro = "Sem cabeçalho neste arquivo, é necessário que possua estes:\n{}".format(cabecalho_esperado)
        if win is None:
            print(_mensagem_de_erro)
        else:
            messagebox.showerror("Arquivo inválido", _mensagem_de_erro)
        exit

    # Analisar se existe algum separador, isso pode funcionar para todos os tipos.
    dialect = csv.Sniffer().sniff(inicio_do_arquivo)
    if dialect is None or dialect.delimiter == ' ' or dialect.delimiter == '':
        #poderia perguntar qual delimitador usar, logo fica a mercê do usuário.
        #se possuir tela (win), então chama dialog
        if win is None:
            separador = input("Qual delimitador usar: ")
        else:
            separador = simpledialog.askstring("Questão", "Qual separador usar?")
        
        # Reverifica caso a condição seja cancelada pelo usuário (if win...)
        if separador is None or separador == ' ' or separador == '':
            separador = ','

    else:
        separador = dialect.delimiter

    #volta para o inicio do arquivo
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
        f = filedialog.asksaveasfile(initialfile = f'{arquivo_nome_com_data}.csv', defaultextension=".csv",filetypes=[("Tabela csv","*.csv"),("Documento texto","*.txt")])
        if f is None:
            # Se a seleção do local é cancelada, então é encerrado.
            exit
        # extrair o filename do objeto retornado pela tela
        nome_arquivo_saida = f.name

    #abre arquivo de saida
    with open(nome_arquivo_saida, 'w') as arquivo_output:
        with open(nome_arquivo_saida+'.err', 'a') as arquivo_erros:
            writer = csv.DictWriter(arquivo_output, cabecalho_output, lineterminator='\n')
            
            #escreve cabecalho novo
            writer.writeheader()
            
            # variavel que contera todas as linhas do arquivo original, eh incrementada 
            # enquanto fica lendo no loop abaixo, e será gravada de uma só vez no final 
            all_rows = []

            try:
                for row in reader:
                    row_saida = {}
                    
                    inicio = datetime.datetime.strptime(row['INICIO'], '%d/%m/%Y %H:%M')
                    end = datetime.datetime.strptime(row['FINAL'], '%d/%m/%Y %H:%M')

                    if inicio > end:
                        msg = "Datas invertidas linha {}".format(reader.line_num-1)
                        arquivo_erros.write("{}\n".format(msg))
                        print(msg)

                    # adiciona na variavel da saida
                    #row.append(businesshrs.difference(inicio, end))
                    bdiff = horas_uteis.difference(inicio, end)
                    row['TIME_REAL'] = "{}:{}:00".format(bdiff.hours, f"{int(bdiff.seconds/60):02d}")

                    estado = row['UF']
                    if estado == "":
                        estado = 'SP'
                    feriados = holidays.BR(state=estado)

                    #adicionar os regionais
                    city = row['CITY']
                    if city != '':
                        requer_df = df_regionais[df_regionais['cod']==int(city)]
                        if requer_df is not None:
                            for r in requer_df['dt'].to_list():
                                print("cidade: {} tem feriado em {}".format(city, r.replace('"', "").replace("'", "")))
                            feriados.append(requer_df['dt'].to_list())
                        else:
                            print('n achou cidade')

                    regras_feriados = businesstimedelta.HolidayRule(feriados)

                    #horas_uteis = businesstimedelta.Rules([diadetrabalho, lunchbreak, regras_feriados])
                    businesshrs = businesstimedelta.Rules([diadetrabalho, regras_feriados])
                    bdiff = businesshrs.difference(inicio, end)
                    row['TIME_OK'] = "{}:{}:00".format(bdiff.hours, f"{int(bdiff.seconds/60):02d}")

                    _segs_por_dia = 24*60*60 # horas x minutos x segundos
                    #row['H_DECIMAL'] = "{:.2f}".format(bdiff.hours+(bdiff.seconds/60/60)).replace(".", ",") # formatar em float 0.00
                    row['H_DECIMAL'] = "{}".format(bdiff.hours+(bdiff.seconds/60/60)).replace(".", ",")

                    all_rows.append(row)

            except csv.Error as e:
                print('erro lendo {}, linha {}: {}'.format(nome_arquivo_saida,reader.line_num-1, e))
            
            #escreve toda variavel par3a arquivo
            writer.writerows(all_rows)
            quantidade_de_registros_gravados = len(all_rows)

#finaliza com alguma msg pro usuario
msg = 'Processou {} registros'.format(quantidade_de_registros_gravados)
if writer is None:
    print(msg)
else:
    messagebox.showinfo("Encerrou", msg)
