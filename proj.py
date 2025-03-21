#Подключаем модули для работ : 1.с операционной системой, 2.с файлами формата json(для сохранения и загрузки пути файла), 3.с фаилами формата Excel, 4.создание и работа с треем, 5.планирования действий функций, 6.многопоточном режиме, 7.работы с времинем
import os, json, openpyxl, pystray, schedule, threading, time 
#Модуль для создания окна
import tkinter as tk
from tkinter import filedialog
#Модуль работы с датой и временем
from datetime import datetime, timedelta
from PIL import Image

# Глобальные переменные
file_path = None
todays = datetime.now() #Переменная сегодняшнего дня
raz_year1 = timedelta(days=365) #Переменная года
raz_year2 = timedelta(days=365*2) #Переменная 2 лет
raz_year3 = timedelta(days=365*3) #Переменная 3 лет

#Переменнные для сравнения остаточных дней
ost = timedelta(days=0) 
ost_15 = timedelta(days=15)
ost_30 = timedelta(days=30)

# Списки для хранения результатов
prosroch_ot = []
d15_ot = []
d30_ot = []
prosroch_fnpi = []
d15_fnpi = []
d30_fnpi = []
prosroch_fnptb = []
d15_fnptb = []
d30_fnptb = []
prosroch_ppb = []
d15_ppb = []
d30_ppb = []
prosroch_opeas = []
d15_opeas = []
d30_opeas = []


# Очистка списков перед формированием нового отчета
def clear_spiskov():
    
    prosroch_ot.clear()
    d15_ot.clear()
    d30_ot.clear()
    prosroch_fnpi.clear()
    d15_fnpi.clear()
    d30_fnpi.clear()
    prosroch_fnptb.clear()
    d15_fnptb.clear()
    d30_fnptb.clear()
    prosroch_ppb.clear()
    d15_ppb.clear()
    d30_ppb.clear()
    prosroch_opeas.clear()
    d15_opeas.clear()
    d30_opeas.clear()

# Функция для сохранения отчета в файл
def save_report(report, filename="last_report.txt"):
    with open(filename, "w", encoding="utf-8") as file:
        file.write(report)

# Функция для загрузки отчета из файла
def load_report(filename="last_report.txt"):
    global message1
    if os.path.exists(filename):
        with open(filename, "r", encoding="utf-8") as file:
             message1=file.read()
             return message1
    return None

# Основная логика загрузки последнего отчета
def get_last_report(message):
    # Пытаемся загрузить последний отчет
    last_report = load_report()

    if last_report is None:
        # Если отчета нет, создаем новый
        new_report = generate_report(message)
        save_report(new_report)
        return new_report
    else:
        # Если отчет есть, возвращаем его
        return last_report


#Функция для формирования отчета для каждого списка
def prosrok(sheet):
    global prosroch_ot, d15_ot, d30_ot, prosroch_fnpi, d15_fnpi, d30_fnpi, prosroch_fnptb, d15_fnptb, d30_fnptb, prosroch_ppb, d15_ppb, d30_ppb, prosroch_opeas, d15_opeas, d30_opeas
    # ОТ и ЭБ
    for row in range(2, sheet.max_row):
        name = str(sheet.cell(row=row, column=1).value)
        dolsh = str(sheet.cell(row=row, column=2).value)
        date_ot = str(sheet.cell(row=row, column=3).value)
        if date_ot == "-" or date_ot == "None":
            continue
        else:
            date_ot = datetime.strptime(date_ot, "%Y-%m-%d %H:%M:%S")
            if dolsh == "НС ТЦ-3" or dolsh == "ВИУТ":
                ostatok_ot = (date_ot + raz_year1) - todays
            elif dolsh in ["Инженер", "Инженер 3кат.", "Инженер ГТС", "Инженер 1кат.", "ВИГТС", "ИРиМ", "Инженер(рег)", "ВИППР", "ВИР", "ВИЭ(ТМО)", "ВИЭ", "ЗНТЦ-3Р", "ЗНТЦ-3Э", "НТЦ-3"]:
                ostatok_ot = (date_ot + raz_year3) - todays
            else:
                ostatok_ot = (date_ot + raz_year2) - todays

            if ostatok_ot <= ost:
                prosroch_ot.append(name)
            elif ost < ostatok_ot <= ost_15:
                d15_ot.append(name)
            elif ost_15 < ostatok_ot <= ost_30:
                d30_ot.append(name)

    # ФНП(ЯБ)
    for row in range(2, sheet.max_row):
        name = str(sheet.cell(row=row, column=1).value)
        dolsh = str(sheet.cell(row=row, column=2).value)
        date_fnpi = str(sheet.cell(row=row, column=4).value)
        if date_fnpi == "-" or date_fnpi == "None":
            continue
        else:
            date_fnpi = datetime.strptime(date_fnpi, "%Y-%m-%d %H:%M:%S")
            if dolsh == "НС ТЦ-3" or dolsh == "ВИУТ":
                ostatok_fnpi = (date_fnpi + raz_year1) - todays
            elif dolsh in ["Инженер", "Инженер 3кат.", "Инженер ГТС", "Инженер 1кат.", "ВИГТС", "ИРиМ", "Инженер(рег)", "ВИППР", "ВИР", "ВИЭ(ТМО)", "ВИЭ", "ЗНТЦ-3Р", "ЗНТЦ-3Э", "НТЦ-3"]:
                ostatok_fnpi = (date_fnpi + raz_year3) - todays
            else:
                ostatok_fnpi = (date_fnpi + raz_year2) - todays

            if ostatok_fnpi <= ost:
                prosroch_fnpi.append(name)
            elif ost < ostatok_fnpi <= ost_15:
                d15_fnpi.append(name)
            elif ost_15 < ostatok_fnpi <= ost_30:
                d30_fnpi.append(name)

    # ФНП(ТБ)
    for row in range(2, sheet.max_row):
        name = str(sheet.cell(row=row, column=1).value)
        dolsh = str(sheet.cell(row=row, column=2).value)
        date_fnptb = str(sheet.cell(row=row, column=5).value)
        if date_fnptb == "-" or date_fnptb == "None":
            continue
        else:
            date_fnptb = datetime.strptime(date_fnptb, "%Y-%m-%d %H:%M:%S")
            if dolsh == "НС ТЦ-3" or dolsh == "ВИУТ":
                ostatok_fnptb = (date_fnptb + raz_year1) - todays
            elif dolsh in ["Инженер", "Инженер 3кат.", "Инженер ГТС", "Инженер 1кат.", "ВИГТС", "ИРиМ", "Инженер(рег)", "ВИППР", "ВИР", "ВИЭ(ТМО)", "ВИЭ", "ЗНТЦ-3Р", "ЗНТЦ-3Э", "НТЦ-3"]:
                ostatok_fnptb = (date_fnptb + raz_year3) - todays
            else:
                ostatok_fnptb = (date_fnptb + raz_year2) - todays

            if ostatok_fnptb <= ost:
                prosroch_fnptb.append(name)
            elif ost < ostatok_fnptb <= ost_15:
                d15_fnptb.append(name)
            elif ost_15 < ostatok_fnptb <= ost_30:
                d30_fnptb.append(name)

    # ППБ
    for row in range(2, sheet.max_row):
        name = str(sheet.cell(row=row, column=1).value)
        dolsh = str(sheet.cell(row=row, column=2).value)
        date_ppb = str(sheet.cell(row=row, column=6).value)
        if date_ppb == "-" or date_ppb == "None":
            continue
        else:
            date_ppb = datetime.strptime(date_ppb, "%Y-%m-%d %H:%M:%S")
            if dolsh == "НС ТЦ-3" or dolsh == "ВИУТ":
                ostatok_ppb = (date_ppb + raz_year1) - todays
            elif dolsh in ["Инженер", "Инженер 3кат.", "Инженер ГТС", "Инженер 1кат.", "ВИГТС", "ИРиМ", "Инженер(рег)", "ВИППР", "ВИР", "ВИЭ(ТМО)", "ВИЭ", "ЗНТЦ-3Р", "ЗНТЦ-3Э", "НТЦ-3"]:
                ostatok_ppb = (date_ppb + raz_year3) - todays
            else:
                ostatok_ppb = (date_ppb + raz_year2) - todays

            if ostatok_ppb <= ost:
                prosroch_ppb.append(name)
            elif ost < ostatok_ppb <= ost_15:
                d15_ppb.append(name)
            elif ost_15 < ostatok_ppb <= ost_30:
                d30_ppb.append(name)

    # ОПЭ АС - оперативный персонал
    for row in range(2, sheet.max_row):
        name = str(sheet.cell(row=row, column=1).value)
        dolsh = str(sheet.cell(row=row, column=2).value)
        date_opeas = str(sheet.cell(row=row, column=7).value)
        if date_opeas == "-" or date_opeas == "None":
            continue
        else:
            date_opeas = datetime.strptime(date_opeas, "%Y-%m-%d %H:%M:%S")
            if dolsh == "НС ТЦ-3" or dolsh == "ВИУТ":
                ostatok_opeas = (date_opeas + raz_year1) - todays
            elif dolsh in ["Инженер", "Инженер 3кат.", "Инженер ГТС", "Инженер 1кат.", "ВИГТС", "ИРиМ", "Инженер(рег)", "ВИППР", "ВИР", "ВИЭ(ТМО)", "ВИЭ", "ЗНТЦ-3Р", "ЗНТЦ-3Э", "НТЦ-3"]:
                ostatok_opeas = (date_opeas + raz_year3) - todays
            else:
                ostatok_opeas = (date_opeas + raz_year2) - todays

            if ostatok_opeas <= ost:
                prosroch_opeas.append(name)
            elif ost < ostatok_opeas <= ost_15:
                d15_opeas.append(name)
            elif ost_15 < ostatok_opeas <= ost_30:
                d30_opeas.append(name)

#Формирование основного отчета
def generate_report():
    global message
    if file_path:
        wb = openpyxl.load_workbook(filename=file_path)
        sheet1 = wb["Оперативный персонал"]  # Открыть лист "Оперативный персонал"
        sheet2 = wb["АТП"]  # Открыть лист "АТП"
        sheet3 = wb["Спец.резерв"]  # Открыть лист "Спец.резерв"
        # Очистка списков перед формированием нового отчета
        clear_spiskov()
        #Формируем отчет
        prosrok(sheet1)
        prosrok(sheet2)
        prosrok(sheet3)
        message = "Сроки экзаменов по ОТ И ЭБ:\nПросроченные: {}\nОсталось 15 дней: {}\nОсталось 30 дней: {}\n---------------\nСроки экзаменов по ФНП(ЯБ):\nПросроченные: {}\nОсталось 15 дней: {}\nОсталось 30 дней: {}\n---------------\nСроки экзаменов по ФНП(ТБ):\nПросроченные: {}\nОсталось 15 дней: {}\nОсталось 30 дней: {}\n---------------\nСроки экзаменов по ППБ:\nПросроченные: {}\nОсталось 15 дней: {}\nОсталось 30 дней: {}\n---------------\nСроки экзаменов по ОПЭ АС:\nПросроченные: {}\nОсталось 15 дней: {}\nОсталось 30 дней: {}\n---------------".format(
    (", ".join(prosroch_ot)), (", ".join(d15_ot)), (", ".join(d30_ot)),
    (", ".join(prosroch_fnpi)), (", ".join(d15_fnpi)), (", ".join(d30_fnpi)),
    (", ".join(prosroch_fnptb)), (", ".join(d15_fnptb)), (", ".join(d30_fnptb)),
    (", ".join(prosroch_ppb)), (", ".join(d15_ppb)), (", ".join(d30_ppb)),
    (", ".join(prosroch_opeas)), (", ".join(d15_opeas)), (", ".join(d30_opeas))
)
        save_report(message)
    else:
        load_report()
        message=message1

def okno():
    try:
        generate_report()
    except:
        load_report()
    # Основное окно
    root = tk.Tk()
    root.title("Информация сроках экзамена персонала ТЦ-3")
    # Размеры окна
    root.resizable(False, False)
    # Кнопка для вызова всплывающего окна
    text1 = tk.Label(root, text=message,  wraplength=280, justify="left")
    text1.pack(fill="both", expand=True, padx=10, pady=10)

    # Запуск основного цикла
    root.mainloop()

# Функция для выбора файла
def choose_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        print(f"Выбран файл: {file_path}")
        return file_path
    return None

# Функция для сохранения пути файла
def save_file_path(file_path):
    with open("file_path.json", "w") as file:
        json.dump({"file_path": file_path}, file)
    print(f"Путь сохранен: {file_path}")

# Функция для загрузки сохраненного пути
def load_file_path():
    if os.path.exists("file_path.json"):
        with open("file_path.json", "r") as file:
            data = json.load(file)
            file_path = data["file_path"]
            if os.path.exists(file_path):  # Проверяем, существует ли файл
                print(f"Загружен путь: {file_path}")
                return file_path
            else:
                print("Файл по сохраненному пути не найден.")
    else:
        print("Сохраненный путь не найден.")
    return None

# Основная функция для проверки и выбора файла
def get_file_path():
    # Пытаемся загрузить сохраненный путь
    file_path = load_file_path()

    # Если путь не загружен или файл недоступен, выбираем новый файл
    if not file_path:
        file_path = choose_file()
        if file_path:
            save_file_path(file_path)  # Сохраняем новый путь
        else:
            print("Файл не выбран.")

    return file_path

#Функция выхода из программы
def on_exit(icon, item):
    icon.stop()
    root.quit()

#Функция для загрузки изображение для иконки трея
def create_tray_icon():
    image = Image.open("icon.ico")  # Загрузите иконку для трея
    menu = pystray.Menu(pystray.MenuItem("Выбрать файл", choose_file), pystray.MenuItem("Сформировать отчет", okno), pystray.MenuItem("Выход", on_exit))
    icon = pystray.Icon("name", image, "Отчеты", menu)
    icon.run()

#Функция для запуска по времени
def schedule_report():
    # Остальные задачи для формирования отчета
    schedule.every().day.at("00:00").do(okno)
    schedule.every().day.at("04:00").do(okno)
    schedule.every().day.at("08:00").do(okno)
    schedule.every().day.at("12:00").do(okno)
    schedule.every().day.at("16:00").do(okno)
    schedule.every().day.at("20:00").do(okno)
    #schedule.every(5).seconds.do(okno)   #пробный запуск

    # Обновление данных и формирование отчета в 23:50
    schedule.every().day.at("23:50").do(clear_spiskov)
    while True:
        schedule.run_pending()
        time.sleep(60)

#Функция для запуска по времени в многопоточном режиме
def run_scheduler():
    threading.Thread(target=schedule_report, daemon=True).start()

#Запуск программы на прямую
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Скрываем основное окно tkinter

    # Получаем путь к файлу
    file_path = get_file_path()

    if file_path:
        print(f"Используемый файл: {file_path}")
    else:
        print("Файл не выбран или недоступен.")
    
    run_scheduler()

    # Запуск иконки в трее
    create_tray_icon()
    
    
    
