
# coding: utf-8

# In[1]:


# добавление библиотек для работы с экселем
import pandas as pd
import numpy as np
# добавление библиотек для работы с операционной системой
import os
import sys
# добавление библиотеки для открытия Outlook
import win32com.client
from win32com.client import Dispatch, constants
# добавление библиотеки для работы со временем
import time
from time import sleep
# добавление библиотеки работы с папками
import glob
# добавление библиотеки для всплывающих окон и выбора файлов
from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilename

# проверка и cчитывание SSO посредством ввода пользователем
while True:
    vaw_SSO = input("Введите ваш SSO: ")
    if os.path.isdir('C:\\Users\\{}'.format(vaw_SSO)):
        break
    else:
        print("\nЭто не ваш SSO!\n")
        
# текст показывающий данный этап
print('\n'+'='*8+' Создание соответствующих папок... '+'='*8)

# проверка выполняется ли данная операция для Healthcare посредством ввода пользователем
health_or_not = input("Выполняете ли Вы данную операцию для GE Healthcare? (Введите Y если да): ")

# создание пути для создания необходимых папок
# проверка выполняется ли данная операция для Healthcare
# если да, то
if health_or_not.lower()=='y':
    # для Healthcare файлы сохранятся в папке Рабочий стол->Vacations->Сегодняшняя дата->(Vacations by teams ИЛИ Champions)
    result_path = 'C:\\Users\\{}\\Desktop\\Vacations\\{}\\Vacations by teams'.format(vaw_SSO,time.strftime("%d")+" "+time.strftime("%B"))
    champions_path = 'C:\\Users\\{}\\Desktop\\Vacations\\{}\\Champions'.format(vaw_SSO,time.strftime("%d")+" "+time.strftime("%B"))
# если нет, то
else:
    # запросить у пользователя название юр.лица (как написано на вкладке файла backlog)
    entity = input("Введите название юр.лица для которого выполняется данная операция (как написано во вкладке):")
    # для других юр.лиц файлы сохранятся в папке Рабочий стол->Vacations->Название юр.лица->Сегодняшняя дата->(Vacations by teams ИЛИ Champions)
    result_path = 'C:\\Users\\{}\\Desktop\\Vacations\\{}\\{}\\Vacations by teams'.format(vaw_SSO,time.strftime("%d")+" "+time.strftime("%B"),entity)
    champions_path = 'C:\\Users\\{}\\Desktop\\Vacations\\{}\\{}\\Champions'.format(vaw_SSO,time.strftime("%d")+" "+time.strftime("%B"),entity)

# создание папок упомянутых выше (если их не существует)
if not os.path.exists(result_path):
    os.makedirs(result_path)
if not os.path.exists(champions_path):
    os.makedirs(champions_path)    
time.sleep(1.5)

# текст показывающий данный этап
print('\n'+'='*8+' Папки созданы! '+'='*8+'\n')
time.sleep(1)

# инициализация для всплывающих окон
root = Tk()
root.attributes('-topmost', 1)
root.withdraw()

# создание всплывающего окна с нотификацией
messagebox.showinfo("Check #1", "Пожалуйста, укажите отчет об отпусках (.xlsx формат)")
# отображение панели диалога для выбора файла backlog
backlog_filename = askopenfilename() 
# создание всплывающего окна с нотификацией
messagebox.showinfo("Check #2", "Пожалуйста, укажите файл headcount из Oracle (.xlsx format)")
# отображение панели диалога для выбора файла headcount
headcount_filename = askopenfilename()

# проверка выполняется ли данная операция для Healthcare
# если да, то
if health_or_not.lower()=='y':
    # cчитывание вкладки HEALTHCARE из эксель файла backlog
    df_backlog_HC = pd.read_excel(backlog_filename,'HEALTHCARE')
    # добавление строки юр.лицо (GE Healthcare) для сотрудников GE Healthcare
    df_backlog_HC["Юр.лицо"] = "GE Healthcare"
     # cчитывание вкладки NYCOMED из эксель файла backlog
    df_backlog_NC = pd.read_excel(backlog_filename,"NYCOMED")
    # добавление строки юр.лицо (Nycomed) для сотрудников Nycomed
    df_backlog_NC["Юр.лицо"] = "Nycomed"

    # соединение данных в один дата фрейм
    frames =[df_backlog_HC,df_backlog_NC]
    df_backlog=pd.concat(frames)

# если нет, то
else:
    # считывание вкладки,название которой было введено пользователем
    df_backlog = pd.read_excel(backlog_filename,'{}'.format(entity))
    # добавление вкладки юр.лицо
    df_backlog["Юр.лицо"] = "{}".format(entity)
    
# cчитывание вкладки Details из отчета headcount
df_headcount = pd.read_excel(headcount_filename, 'Details')
# переименовывание название колонки 'SSO ID' в 'Табельный номер'
df_headcount=df_headcount.rename(columns={'SSO ID':'Табельный номер'})

# извлечение необходимых колонок из отчета backlog
backlog_required_columns = ["Табельный номер", "Ф.И.О. сотудника (англ.)","Должность (англ.)","Дата приема","Юр.лицо",'Осталось всего дней  отпуска','Основной ежегодный отпуск', 'Доп отпуск НРД', 'Доп отпуск МКС']
df_backlog = df_backlog[backlog_required_columns]
# конвертирование содержимого колонки 'Табельный номер' в текст
df_backlog['Табельный номер']=df_backlog['Табельный номер'].astype(str)

# извлечение необходимых колонок из отчета headcount
headcount_required_columns = ["Табельный номер",'HRM Name','Manager Name',"Manager SSO"]
# считывание всех рядов, кроме последних двух - они не нужны
df_headcount = df_headcount.iloc[:df_headcount.shape[0]-2]
# конвертирование содержимого колонки 'Manager SSO' в целое число
df_headcount["Manager SSO"] = df_headcount["Manager SSO"].astype(int)
# конвертирование содержимого всех колонок в текст
df_headcount= df_headcount[headcount_required_columns].astype(str)

# cоединение данных из двух отчетов в один (аналогично функции VLOOKUP в экселе)
df_result=df_backlog.merge(df_headcount,left_on='Табельный номер',right_on='Табельный номер',how='left')

# удаление всех запятых в конечном дата фрейме
for cols in headcount_required_columns:
    df_result[cols]=df_result[cols].str.replace(',','')
    
# упорядочивание колонок в определенный формат
df_result = df_result[["Табельный номер","Ф.И.О. сотудника (англ.)","Должность (англ.)",'HRM Name','Manager Name',"Manager SSO","Дата приема","Юр.лицо",'Осталось всего дней  отпуска','Основной ежегодный отпуск', 'Доп отпуск НРД', 'Доп отпуск МКС']]
# замена всех ячеек со значением NaN на пустой знак
df_result = df_result.replace(np.nan, '', regex=True)

# получение Количества дней неиспользованного отпуска для чемпионов
while True:
    champion_days = input("\nКоличество дней неиспользованного отпуска для чемпионов (число): ")
    try:
        ch_int = int(champion_days)
    except ValueError:
        print("\nПожалуйста введите число...")
        continue
    else:
        break
        
# текст показывающий данный этап
print('\n'+'='*8+' Создание файла для чемпионов... '+'='*8)

# сохранение элементов дата фрейма, в которых значение в колонке 'Осталось всего дней  отпуска' больше чем число
# введенное пользователем в эксель файл
ch_writer = pd.ExcelWriter(os.path.join(champions_path,'Чемпионы +{}.xlsx'.format(str(ch_int))), engine = 'xlsxwriter')
df_result.loc[df_result['Осталось всего дней  отпуска']>int(champion_days)].sort_values("Ф.И.О. сотудника (англ.)").to_excel(ch_writer,sheet_name='Sheet1')

# приведение файла в нужный формат - ширина колонок и т.д.
workbook1 = ch_writer.book
worksheet1 = ch_writer.sheets['Sheet1']
      
header_format1 = workbook1.add_format({
    'bold': True,
    'text_wrap': True,
    'align': 'center',
    'valign': 'vcenter'})
        
worksheet1.set_zoom(90)
worksheet1.set_column('A:A',0)
worksheet1.set_column('B:B',11.5)
worksheet1.set_column('C:C',30)
worksheet1.set_column('D:D',33)
worksheet1.set_column('E:E',20)
worksheet1.set_column('F:F',24)
worksheet1.set_column('G:G',0)
worksheet1.set_column('H:H',13)
worksheet1.set_column('I:I',12.7)
worksheet1.set_column('J:J',9)
worksheet1.set_column('K:K',9)
worksheet1.set_column('L:L',9)
worksheet1.set_column('M:M',9)
worksheet1.set_row(0,None,header_format1)

for colx, value in enumerate(df_result.columns.values):
    worksheet1.write(0, colx+1, value,header_format1)        
ch_writer.save()
ch_writer.close()

# текст показывающий данный этап
print('\n'+'='*8+' Файл для чемпионов создан! '+'='*8+'\n')
time.sleep(1) 

# текст показывающий данный этап
print('\n'+'='*8+' Создание файлов для менеджеров... '+'='*8)

# создание файлов для менеджеров
# проведение операции для каждого уникального значения в колонке 'Manager SSO'
for manager_SSO in set(df_result['Manager SSO'].values):
    # Если у сотрудника есть менеджер
    if not manager_SSO == '':
        # сохранить всех сотрудников с определенным менеджеров в файл с названием имени менеджера
        writer = pd.ExcelWriter(os.path.join(result_path,"{}.xlsx".format(df_result.loc[(df_result['Manager SSO'] == manager_SSO)]['Manager Name'].values[0])), engine = 'xlsxwriter')
        df_result.loc[(df_result['Manager SSO'] == manager_SSO) | 
          (df_result['Табельный номер'] == manager_SSO)].sort_values("Manager Name").to_excel(writer,sheet_name='Sheet1')
    # если у сотрудник нет менеджера
    else:
        # для каждого значения в колонке 'Ф.И.О. сотудника (англ.)'
        for employee in set(df_result.loc[df_result["Manager SSO"]==manager_SSO]['Ф.И.О. сотудника (англ.)']):
            # сохранить всех сотрудников без менеджеров в файл с названием имени сотрудника
            writer = pd.ExcelWriter(os.path.join(result_path,"{}.xlsx".format(employee)), engine = 'xlsxwriter')
            df_result.loc[df_result['Ф.И.О. сотудника (англ.)']==employee].to_excel(writer,sheet_name = 'Sheet1')
    
    # приведение файла в нужный формат - ширина колонок и т.д.      
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'align': 'center',
    'valign': 'vcenter'})
    
    worksheet.set_zoom(90)
    worksheet.set_column('A:A',0)
    worksheet.set_column('B:B',11.5)
    worksheet.set_column('C:C',30)
    worksheet.set_column('D:D',33)
    worksheet.set_column('E:E',20)
    worksheet.set_column('F:F',24)
    worksheet.set_column('G:G',0)
    worksheet.set_column('H:H',13)
    worksheet.set_column('I:I',12.7)
    worksheet.set_column('J:J',9)
    worksheet.set_column('K:K',9)
    worksheet.set_column('L:L',9)
    worksheet.set_column('M:M',9)
    worksheet.set_row(0,None,header_format)

    for colx, value in enumerate(df_result.columns.values):
        worksheet.write(0, colx+1, value,header_format)
        
    writer.save()        

# текст показывающий данный этап
print('\n'+'='*8+' Файлы для менеджеров созданы! '+'='*8+'\n')

# переменная для сохранения пути созданных файлов для менеджеров
path = os.path.abspath(result_path+"//*.xlsx")
const=win32com.client.constants

# текст показывающий данный этап
print('\n'+'='*8+' Cоздание шаблона письма... '+'='*8+'\n')

# cчитывание данных от пользователя для внесения в шаблон письма
subject = input("\nВведите тему письма: ")
info_date = input("\nОт какого числа данные в файле? (формат число месяц): ")
specialist_name = input("\nВведите имя HR Ops специалиста: ")
specialist_email = input("\nВведите корпоративный e-mail этого специалиста: ")
days_unused = input("\nСколько дней не должен превышать неиспользованный отпуск на конец года? (число):")

extra = input("\nВведите дополнительный пункт при обсуждении графиков отпусков: ")
mesto = input("\nВведите департамент, блок и этаж где сидит HR Ops специалист: ")

ot_kogo = input("\nВведите имя человека, которое должно быть в подписи: ")
extension = input("\nВведите Ваш  extension номер (4 цифры): ")
mobile = input("\nВведите ваш мобильный номер (в формате +7 XXX-XXX-XX-XX): ")
sign = input("\nВведите блок и этаж для подписи (на англ.):")

# текст показывающий данный этап
print('\n'+'='*8+' Данные для шаблона внесены! '+'='*8+'\n')
time.sleep(2)

# текст показывающий данный этап
print('\n'+'='*8+' Генерируем письма... '+'='*8+'\n')

# создание писем для каждого файла с названием менеджера или сотрудника
for fname in glob.glob(path):
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    # cоздание темы письма
    newMail.Subject = "{}".format(subject)
    newMail.BodyFormat = 2
    # cоздание шаблона с данными введеными пользователем
    newMail.HTMLBody = ("<HTML><BODY>Добрый день!     <br> <br>    Направляю актуальную информацию по неиспользованным отпускам по вашей команде. В файле данные от {}.     <br> <br>    Сотрудникам и менеджеру необходимо убедиться, что ранее использованные отпуска были правильно оформлены и количество неиспользованного отпуска отображается корректно. Если отпуск был использован, но не был оформлен и количество дней отображается неправильно, вам необходимо обратиться к <strong>{}</strong> <a href="r"maito:{}> {}</a>.    <br> <br>    <strong>При обсуждении графиков отпусков с командой прошу учесть данные пункты:</strong>    <ul>      <li>{}</li>      <li>При планировании отпуска неиспользованный отпуск на конец года не должен превышать <strong>{} дней</strong>;</li>      <li>Планировать отпуск необходимо в календарных днях с учетом выходных;</li>      <li>Можно взять денежную компенсация за ненормированный рабочий день (НРД), написав заявление;</li>      <li>Можно спланировать отпуск с пятницы по понедельник, включая выходные (в этом случае выходные дни будут оплачиваться);</li>      <li>Важно спланировать основной отпуск, который должен быть не менее 14 дней.</li>    </ul>    <br>    <strong>После согласования графиков отпусков сотрудники должны передать заявления на отпуск на имя {} в {}.</strong>    <br> <br>    <strong>{}</strong>    <br> <br>    HR Advisory    <br> <br>    GE Healthcare Russia and CIS    <br> <br>    <small>T +7 (495) 739-69-19 ext. {} | M  {}</small>    <br><br>     <a href="r"www.gehealthcare.com>www.gehealthcare.com</a>    <br><br>    Presnenskaya nab., 10, {} | Moscow, 123317, Russia    </BODY></HTML>").format(info_date,specialist_name,specialist_email,specialist_email,extra,days_unused,specialist_name, mesto, ot_kogo,
                           extension,mobile,sign)

    # добавление нужных файлов как приложение к письму
    attachment1 = str(fname).replace("\\","\\\\")
    newMail.Attachments.Add(Source=attachment1)
    newMail.display(True)

time.sleep(1)
# текст показывающий данный этап
print('\n'+'='*8+' Операция завершена! '+'='*8+'\n')
# стирание всех переменных
sys.modules[__name__].__dict__.clear()

