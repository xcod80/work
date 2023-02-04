from sys import argv
import os
import platform
import win32com.client

#Выбираем команду пинг в зависимости от ОС
def SelectPing():
    oc = platform.system()
    if (oc == "Windows"):
        ping_com = "ping -n 1 "
    else:
        ping_com = "ping -c 1 "
    return ping_com

#Пинугуем адрес
def PingAddr(comm):
    response = os.popen(comm).readlines()
    for line in response:
        if 'ttl' in line.lower():
            return True
    else:
        return False

#Пингуем выбранные устройства
def PingDevices(DevType):
    i = 3
    value = sheet.Cells(i, 10).value
    valuedevice = sheet.Cells(i, 6)
    while value != None:
        if str(valuedevice) == DevType:
            if not PingAddr(pingcmd + str(value)):
                print(f'Строка:{i}\t{DevType}\t{value}\t{sheet.Cells(i, 2)}\t{sheet.Cells(i, 3)}\t{sheet.Cells(i, 4)}\t{sheet.Cells(i, 14)}')
        i += 1
        value = sheet.Cells(i, 10).value
        valuedevice = sheet.Cells(i, 6)

#Начало
choice = 999999
scriptname, argument = argv[0], argv[1]
types = []
pingcmd = SelectPing()

#Открываем книгу эксель
Excel = win32com.client.Dispatch("Excel.Application")
mb = Excel.Workbooks.Open(argument)
#Открываем нужные таблицы
sheettypes = mb.Sheets["Таблицы"]
sheet = mb.Sheets["Оборудование"]

#Считываем типы оборудования
i = 2
value = sheettypes.Cells(i,1).value
while value != None:
    types.append(str(value))
    i += 1
    value = sheettypes.Cells(i,1).value

#Конец подготовки и начало работы меню программы.
while True:
    try:
        try:
            choice = int(argv[2])
        except Exception:
            #Рисуем меню
            for number in range(0, len(types)):
                print(f'{number}\t- {str(types[number])}')
            print(f'99\t- выход')
            #Ждем ввод пользователя
            choice = int(input("Выберите пункт меню:"))
        else:
            if choice <= len(types):
                print("\n", types[choice], "\n")
                PingDevices(types[choice])
            exit()
    except Exception:
        print()
    if choice < len(types):
        print("\n", types[choice], "\n")
        PingDevices(types[choice])
    elif choice == 99:
        exit()


