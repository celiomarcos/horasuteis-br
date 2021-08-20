import csv
import datetime
from datetime import date
import businesstimedelta
import holidays

import tkinter as tk
from tkinter import TclError, filedialog, messagebox

file_path = 'dataspracomparar.csv'
root = None
try:
    root = tk.Tk()
except TclError:
    print('nao vai rolar tela')
else:
    root.withdraw()
    file_path = filedialog.askopenfilename()

d1= datetime.datetime.now() # time object

print("now =", d1)
print("type(now) =", type(d1))	

filename = d1.strftime("%Y%m%d_%H%M%S")
# dd/mm/YY

#print("d1 =", d1)



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

with open(file_path,'r') as data_input:
    with open(f'{filename}.csv', 'w') as data_output:
        reader = csv.reader(data_input, delimiter=',')
        writer = csv.writer(data_output, lineterminator='\n')
        
        all = []
        row = next(reader)
        print("HEADER:", row)
        validado = ['DATA ABERTURA', 'DATA CHECK-IN', 'TIME', 'UF']
        if row != validado:
            if root is None:
                print("CSV inválido!! vaca")
            else:
                messagebox.showerror("cbecqlho invalido", "use {}".format(validado))
            exit
        for column in row:
            # Add validation here
            print(column)

        row.append('calculou_btd')
        row.append('calculou_hol')
        all.append(row)

        for row in reader:
            
            inicio = datetime.datetime.strptime(row[0], '%d/%m/%Y %H:%M:%S')
            end = datetime.datetime.strptime(row[1], '%d/%m/%Y %H:%M:%S')
        
            # adiciona na variavel da saida
            #row.append(businesshrs.difference(inicio, end))
            bdiff = horas_uteis.difference(inicio, end)
            row.append("{}:{}:00".format(bdiff.hours, f"{int(bdiff.seconds/60):02d}"))

            estado = row[3]
            if estado == "":
                estado = 'SP'
            feriados_uf = holidays.BR(state=estado)
            feriados = businesstimedelta.HolidayRule(feriados_uf)

            #horas_uteis = businesstimedelta.Rules([diadetrabalho, lunchbreak, feriados])
            businesshrs = businesstimedelta.Rules([diadetrabalho, feriados])
            bdiff = businesshrs.difference(inicio, end)
            row.append("{}:{}:00".format(bdiff.hours, f"{int(bdiff.seconds/60):02d}"))

            all.append(row)

        #escreve toda variavel para arquivo
        writer.writerows(all)
    

