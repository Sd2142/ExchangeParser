from exchangelib import DELEGATE, Credentials, Account, EWSDateTime, EWSTimeZone, Configuration
import re
import os
import csv
import subprocess

datelisted = []
TaskInc = []

##

PWD = str(input("Enter password\n"))
print('Начинаем процедуру настройки!\nСначала необходимо ввести дату формата: ОТ: ГОД:Месяц:Число. Затем ДО: ГОД:Месяц:Число! \n Если число месяца меньше 10, то необходимо вводить без Ноля')
OtDate = str(input("Введите дату начала формирования отчета. Пример::2019, 12, 5\n"))
DODate = str(input("Введите дату окончания формирования отчета. Пример::2019, 12, 7\n"))
OTDateSplits = (OtDate.split(", "))
DoDateSplits = (DODate.split(", "))
OTDateYears = int(OTDateSplits[0])
OTDateMounth = int(OTDateSplits[1])
OTDateday = int(OTDateSplits[2])
DoDateYears = int(DoDateSplits[0])
DoDateMounth = int(DoDateSplits[1])
DoDateday = int(DoDateSplits[2])
print('Выбрана сортировка обращения ОТ: ',OtDate, ' ДО: ',DODate)
tz = EWSTimeZone.localzone()
creds = Credentials(username="User.Name", password=PWD)
config = Configuration(server='Exchangeserver.@domain.com', credentials=creds)
account = Account(primary_smtp_address="User.Name@domain.com", autodiscover=False, config=config, access_type=DELEGATE)
my_folder = account.inbox / 'Registred_inc'#NameOfFolder
for i in my_folder.filter(datetime_sent__range=(tz.localize(EWSDateTime(OTDateYears, OTDateMounth, OTDateday)), tz.localize(EWSDateTime(DoDateYears, DoDateMounth, DoDateday)))):
    DataPolucheniyaPisma = str(i.datetime_sent)
    datelisted.append(DataPolucheniyaPisma)#DateOfSentmail
##########################
    BodyWhithNumber = str(i.body)
    RazbivkaPoSimvoluBR = (BodyWhithNumber.split("br"))
    NoSplitBODY = RazbivkaPoSimvoluBR[1]
    splitBODY2 = (NoSplitBODY.split(" "))
    NeGotovyINC = (splitBODY2[-1])
    NeGotovyINC2 = (NeGotovyINC.split("."))#
    CompliteToString = ''.join(str(t) for t in NeGotovyINC2[0])
    RealNumberOfTask = (CompliteToString)
    TaskInc.append(RealNumberOfTask)#NumberOfinc
		
countofstroki = int(len(datelisted))

with open("Export.xls", "w", newline='') as csv_file:
    writer = csv.writer(csv_file, delimiter='\t')
    level_counter = 0
    max_levels = countofstroki
    while level_counter < max_levels:
        writer.writerow((TaskInc[level_counter], datelisted[level_counter]))
        level_counter = level_counter +1

print ('Count inc of time: ', countofstroki)
print ('Count inc: ', len(TaskInc))

subprocess.call("Export.xls", shell=True)

input('Press ENTER to exit')
