import csv
import datetime
import os
import tkinter as tk
from datetime import date
from tkinter import TclError, filedialog, messagebox, simpledialog
from progress.bar import FillingCirclesBar

import businesstimedelta
import holidays
import numpy as np
import pandas as pd
# essa lib carrega variaveis de ambiente e se n houver ele considerars as variaveis declaradas no arquivo .env
import dotenv


class Calculadora:
    def __init__(self):
        ###############################
        # definir um dia de e horário de trabalho
        self.diadetrabalho = businesstimedelta.WorkDayRule(
            start_time=datetime.time(8),
            end_time=datetime.time(18),
            working_days=[0, 1, 2, 3, 4])

        # horario de almoco
        self.lunchbreak = businesstimedelta.LunchTimeRule(
            start_time=datetime.time(12),
            end_time=datetime.time(13),
            working_days=[0, 1, 2, 3, 4])

        # combinar os dois
        #horas_uteis = businesstimedelta.Rules([diadetrabalho, lunchbreak])
        self.horas_uteis = businesstimedelta.Rules([self.diadetrabalho])

    def horas_real(self, ini, fim):
        bdiff = self.horas_uteis.difference(ini, fim)
        return bdiff
    
    def horas_com_feriados(self, ini, fim, feriados):
        #montar businesstimedelta com os feriados
        regras_feriados = businesstimedelta.HolidayRule(feriados)

        #horas_uteis = businesstimedelta.Rules([diadetrabalho, lunchbreak, regras_feriados])
        businesshrs = businesstimedelta.Rules([self.diadetrabalho, regras_feriados])
        bdiff = businesshrs.difference(ini, fim)
        return bdiff

def main():
    # carregar as variaveis de ambiente
    dotenv_file = dotenv.find_dotenv() # soh q agora procu5ra o arquivo .env no OS, pq vamos reutilizar ele depois
    dotenv.load_dotenv(dotenv_file)

    PATH_ARQUIVO_BASE = os.getenv('ARQUIVO_BASE', "")
    PATH_ARQUIVO_CIDADES = os.getenv('ARQUIVO_CIDADES', "")
    FORMAT_DT_INICIO = os.getenv('FORMAT_DT_INI', '%d/%m/%Y %H:%M')
    FORMAT_DT_FINAL = os.getenv('FORMAT_DT_FIM', '%d/%m/%Y %H:%M')
    UF_DEFAULT = os.getenv('UF_DEFAULT', 'SP')
    FORMAT_DT_CIDADES = os.getenv('FORMAT_DT_CIDADES', "%d/%m/%Y")
    DELIMITADOR = os.getenv('DELIMITADOR', 'auto')

    COL_DATA_A = os.getenv('COL_DATA_A', 'INICIO')
    COL_DATA_B = os.getenv('COL_DATA_B', 'FINAL')
    COL_COD = os.getenv('COL__COD', 'SR') 
    COL_MUNICIPIO = os.getenv('COL_MUNICIO', 'CODIGO')
    COL_UF = os.getenv('COL_UF', 'UF')

    # cabecalho do arquivo de entrada
    CABECALHO_ESPERADO = [COL_DATA_A, COL_DATA_B, COL_COD, COL_MUNICIPIO, COL_UF]
    CABECALHO_OUTPUT = CABECALHO_ESPERADO + ['TIME_REAL', 'TIME_OK', 'H_DECIMAL', 'HREAL_DECIMAL']

    #TODO montar os cabecalhos por aqui logo dps de definiar quais sao

    # tela usada na gui, inicia em None, porque precisa ser criada no ambiente
    win = None
    try:
        win = tk.Tk()
    except TclError:
        # se tentou criar, e deu problema (nao tem tela por ex), vai ficar win validando None, pra saber que nao pode usar win
        print('Ops... Não foi possível criar tela, sera usado apenas terminal')
    else:
        # como nao queremos uma GUI inteira e sim apenas mostrar alguns Dialogs, se previne a janela root de aparecer
        win.withdraw()

    ####################################
    # arquivo BASE precisa existir
    while os.path.exists(PATH_ARQUIVO_BASE) is False:

        # nao existe, entao pergunta onde arquivo esta abre Dialog perguntando onde o arquivo  esta
        if win is None:
            # pergunta no terminal
            PATH_ARQUIVO_BASE = input(
                "Informe o arquivo base para processar (deixe em branco para cancelar): ")
        else:
            # ou pergunta em Dialog
            PATH_ARQUIVO_BASE = filedialog.askopenfilename(
                title="Informe o arquivo CSV para processar",
                filetypes=[("Comma Separated Values - CSV", ".csv")])

        # se o usuario cancelar essa seleccao, retornara um None, entao interrompe programa pois esse arquivo eh fundamental
        if PATH_ARQUIVO_BASE is None or PATH_ARQUIVO_BASE == '':
            msg = "Você cancelou a seleção do arquivo ❌."
            if win is None:
                print(msg)
            else:
                messagebox.showwarning('Seleção cancelada', msg)
            #encerra tudo
            return

        # sobe ali pro inicio desse while pra testar novamente se arquivo existe

    ###############################
    # carregar arquivo de feriados regionais
    df_regionais = None
    if PATH_ARQUIVO_CIDADES != "" and os.path.exists(PATH_ARQUIVO_CIDADES) is False:
        if win is None:
            PATH_ARQUIVO_CIDADES = input(
                "Selecionar arquivo de feriados municipais com cod. IBGE: ")
        else:
            PATH_ARQUIVO_CIDADES = filedialog.askopenfilename(
                title="Informe o arquivo com feriados de cidades",
                filetypes=[("Formato de planilhas", ".csv")])

        if PATH_ARQUIVO_CIDADES is None or PATH_ARQUIVO_CIDADES == '':
            print("Usuário cancelou a seleção dos feriados municipais, continuando o calculo sem.")
            PATH_ARQUIVO_CIDADES = None
        else:
            # sep=None faz o pandas testar os separador ideal automaticamente
            print("Abrindo calendário de cidades {}".format(PATH_ARQUIVO_CIDADES))
            try:
                df_regionais = pd.read_csv(
                    PATH_ARQUIVO_CIDADES, sep=None if DELIMITADOR == 'auto' else DELIMITADOR, quoting=csv.QUOTE_NONE)
            except pd.errors.EmptyDataError:
                print('Arquivo está vazio: {}'.format(PATH_ARQUIVO_CIDADES))
            except:
                print('Erro ao abrir {}'.format(PATH_ARQUIVO_CIDADES))
            else:
                # já q abriu, rememoriza esse nome de arquivo pra facilitar na proxima execução
                dotenv.set_key(dotenv_file, 'ARQUIVO_CIDADES', PATH_ARQUIVO_CIDADES)

                # TODO devia testar aqui se esta ok com esse arquivo..

    ##################################################
    # arquivo de saída
    d1 = datetime.datetime.now()
    arquivo_nome_com_data = d1.strftime("%Y%m%d_%H%M%S")
    PATH_ARQUIVO_SAIDA = f'.\{arquivo_nome_com_data}.csv'
    # aqui eh o inverso, SE o arquivo de saida existir, dai pergunta um outro nome
    if os.path.exists(PATH_ARQUIVO_SAIDA):
        # confirma onde salvar o arquivo destino
        if win is None:
            f = input("Salvar arquivo de saída: ")
        else:
            f = filedialog.asksaveasfile(
                    title="Informe Arquivo de saída",
                    initialfile=f'{arquivo_nome_com_data}.csv', 
                    defaultextension=".csv", 
                    filetypes=[("Tabela csv", "*.csv"), ("Documento texto", "*.txt")])
            if f is None:
                # Se a seleção do local é cancelada, então é encerrado.
                return
            # extrair o filename do objeto retornado pela tela
            PATH_ARQUIVO_SAIDA = f.name

    ###############################
    # inicializa a 
    calculadora = Calculadora()

    quantidade_de_registros_gravados = 0

    print("Abrindo arquivo {}".format(PATH_ARQUIVO_BASE))
    with open(PATH_ARQUIVO_BASE, 'r') as data_input:

        # já q abriu, memoriza esse nome de arquivo pra facilitar na proxima execução
        dotenv.set_key(dotenv_file, 'ARQUIVO_BASE', PATH_ARQUIVO_BASE)
        
        # alguns testes basicos com o arquivo de entrada

        # para isso le os primeiros 1024 bytes, aumentar se achar que nao for suficiente, mas geralmente supre.
        inicio_do_arquivo = data_input.read(1024)

        tem_cabecalho = csv.Sniffer().has_header(inicio_do_arquivo)
        if not tem_cabecalho:
            _msg = "Sem cabeçalho neste arquivo, é necessário que possua estes:\n{}".format(CABECALHO_ESPERADO)
            if win is None:
                print(_msg)
            else:
                messagebox.showerror("Arquivo inválido ☹", _msg)
            return

        # Se deve ou nao usar um analisador de separador
        if DELIMITADOR.lower() == 'auto':
            # Tentar detectar automaticamente
            dialect = csv.Sniffer().sniff(inicio_do_arquivo)
            if dialect is None or dialect.delimiter == ' ' or dialect.delimiter == '':
                # poderia perguntar qual delimitador usar, logo fica a mercê do usuário.
                # se possuir tela (win), então chama dialog
                if win is None:
                    separador = input("Qual delimitador usar: ")
                else:
                    separador = simpledialog.askstring("Questão", "Qual separador usar?")

                # Reverifica caso a condição seja cancelada pelo usuário
                if separador is None or separador == ' ' or separador == '':
                    # utilizar o do arquivo .env
                    separador = DELIMITADOR
            else:
                separador = dialect.delimiter
        else:
            separador = DELIMITADOR

        # volta para o inicio do arquivo
        data_input.seek(0)

        # agora assim abrir o arquvo para ler as linhas
        reader = csv.DictReader(data_input, delimiter=separador)

        # get fieldnames from DictReader object and store in list
        headers = reader.fieldnames
        print("Cabeçalhos encontrados:" + str(headers))

        
        # verificando se a primeira lista, contem na segunda para ser considerado verdadeiro

        # essa comparação eh um pouco complexa de entender pois são dois objetos lists, 
        if set(CABECALHO_ESPERADO).issubset(set(headers)) is False:
            # sobre esse metodo de comparação, eh necessario observar que eh a mesma coisa que
            # if CABECALHO_ESPERADO <= headers
            # tanto um quanto o outro método, converte as listas para sets, o que significa que nao
            # considera valores duplicados, não afetará neste caso, mas eh uma caracteristica interessante de observar

            if win is None:
               print("Ops! Cabecalho de Arquivo inválido! ☹")
            else:
               messagebox.showerror("Cabeçalho Inválido ☹", "Use estes separados por ponto e vírgula: {}".format(CABECALHO_ESPERADO))
            #de qquer forma encerra
            return

        # mudar essa validacao simples para

        #headers = []
        # for row in reader:
        #    headers = [x.lower() for x in list(row.keys())]
        #    break

        # if 'minha ccoluna' not in headers or 'id_nome' not in headers:
        #    print('Arquivo CSV precisa ter as colunas "Minha Coluna" e a coluna "ID_Nome"')
        
        #maneira mais rapida que sei pra contar quantas linhas tem eh assim:
        qt_rows = sum(1 for _ in reader)
        data_input.seek(0)
        reader.__init__(data_input, delimiter=separador)
        print("Quantidade de linhas do arquivo: {} rows".format(qt_rows))

        # abre arquivo de saida
        print("Abrindo arquivo de saida {}".format(PATH_ARQUIVO_SAIDA))
        with open(PATH_ARQUIVO_SAIDA, 'w') as arquivo_output:
            with open(PATH_ARQUIVO_SAIDA+'.err', 'a') as arquivo_erros:
                writer = csv.DictWriter(arquivo_output, CABECALHO_OUTPUT, lineterminator='\n')

                # escreve cabecalho novo
                writer.writeheader()

                # variavel que contera todas as linhas do arquivo original, eh incrementada
                # enquanto fica lendo no loop abaixo, e será gravada de uma só vez no final
                all_rows = []

                try:
                    bar = FillingCirclesBar('Calculando', max=qt_rows)
                    bar.width = 50 # tamanho da barra

                    for row in reader:
                        row_saida = {}

                        inicio = datetime.datetime.strptime(row[COL_DATA_A], FORMAT_DT_INICIO)
                        end = datetime.datetime.strptime(row[COL_DATA_B], FORMAT_DT_FINAL)

                        if inicio > end:
                            msg = "Data retroativa linha {} SR: {}".format(reader.line_num-1, row[COL_COD])
                            arquivo_erros.write("{}\n".format(msg))
                            print('-- Gravou log: {}'.format(msg))

                        # adiciona na variavel da saida
                        bdiff = calculadora.horas_real(inicio, end)
                        row_saida['TIME_REAL'] = "{}:{}:00".format(bdiff.hours, f"{int(bdiff.seconds/60):02d}")
                        row_saida['HREAL_DECIMAL'] = "{:.2f}".format(bdiff.hours+(bdiff.seconds/60/60)).replace(".", ",")

                        estado = row[COL_UF]
                        if estado == "":
                            estado = UF_DEFAULT
                        feriados = holidays.BR(state=estado)

                        # armazenar aqui quantos feriados municipais encontrou
                        quantos_feriados_municipio = 0

                        # adicionar os regionais SE conseguiu usar o arquivo de CIDADES
                        if df_regionais is not None:
                            city = row[COL_MUNICIPIO]
                            if city != '':
                                requer_df = df_regionais[df_regionais['CODIGO_MUNICIPIO'] == int(city)]

                                # CODIGO_MUNICIPIO,DATE,UF,NOME_MUNICIPIO
                                if requer_df is not None:
                                    for r in requer_df['DATE'].to_list():

                                        # limpar as aspas, trocar traço por barra
                                        r = r.replace('"', "").replace("'", "").replace("-", "/")

                                        # converter a data em formato PT-BR para date de python
                                        dateObj = datetime.datetime.strptime(r, FORMAT_DT_CIDADES).date()

                                        # adicionar aos feriados a serem considerados
                                        feriados.append(dateObj)

                                        quantos_feriados_municipio += 1
                                        print("cidade: {} tem feriado em {}".format(city, r))
                                    print('feriados no municipio {}: {}'.format(city, quantos_feriados_municipio))
                                else:
                                    print('Cidade não encontrada {}'.format(city))
                        
                        # contar quantidade de feriados no periodo
                        # quantos_feriados_total = 0
                        # df = pd.DataFrame()
                        # df['Datas'] = pd.date_range(inicio, end)
                        # for val in df['Datas']:
                        #     if str(val).split()[0] in feriados:
                        #         quantos_feriados_total += 1
                        # if quantos_feriados_total > 0:
                        #     print('feriados total: {}'.format(quantos_feriados_total))

                        bdiff = calculadora.horas_com_feriados(inicio, end, feriados)
                        row_saida['TIME_OK'] = "{}:{}:00".format(bdiff.hours, f"{int(bdiff.seconds/60):02d}")
                        
                        _segs_por_dia = 24*60*60  # horas x minutos x segundos
                        #row_saida['H_DECIMAL'] = "{}".format(bdiff.hours+(bdiff.seconds/60/60)).replace(".", ",")
                        row_saida['H_DECIMAL'] = "{:.2f}".format(bdiff.hours+(bdiff.seconds/60/60)).replace(".", ",") # formatar em float 0.00
                        
                        row_saida[COL_DATA_A] = row[COL_DATA_A]
                        row_saida[COL_DATA_B] = row[COL_DATA_B]
                        row_saida[COL_COD] = row[COL_COD]
                        row_saida[COL_MUNICIPIO] = row[COL_MUNICIPIO]
                        row_saida[COL_UF] = row[COL_UF]

                        all_rows.append(row_saida)

                        #
                        #print('pressando linha: {}'.format(reader.line_num-1))
                        #print(".", end =" ")
                        bar.next()
                    
                    bar.finish()

                except csv.Error as e:
                    msg = 'Erro ao ler {}, linha {}: {}'.format(
                        PATH_ARQUIVO_SAIDA, reader.line_num-1, e)
                    if win is None:
                        print(msg)
                    else:
                        messagebox.showerror(msg)

                # escreve toda variavel par3a arquivo
                writer.writerows(all_rows)
                quantidade_de_registros_gravados = len(all_rows)

    # finaliza com alguma msg pro usuario
    msg = 'Foram processados {} registros, verifique os logs no arquivo ".err" '.format(quantidade_de_registros_gravados)
    if win:
        messagebox.showinfo("Encerrado", msg)
    print(msg)

if __name__ == "__main__":
    main()
