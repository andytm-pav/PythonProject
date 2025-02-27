import PySimpleGUI as sg

# Кортеж с названиями полей
field_names = ("Название 1", "Название 2", "Название 3", "Название 4",
               "Название 5", "Название 6", "Название 7")

# Определяем макет окна (только поля ввода)
layout = [
    [sg.Text(name + ":"), sg.InputText(key=f"-INPUT{i}-")] for i, name in enumerate(field_names)
]

# Создаем окно
window = sg.Window("Моя программа", layout, size=(400, 300))

while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED:
        break

window.close()
