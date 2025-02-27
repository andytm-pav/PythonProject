import PySimpleGUI as sg
import json

# Кортеж с названиями полей
field_names = ("Название 1", "Название 2", "Название 3", "Название 4",
               "Название 5", "Название 6", "Название 7")

# Определяем макет окна
layout = [
             [sg.Text(name + ":"), sg.InputText(key=f"-INPUT{i}-")] for i, name in enumerate(field_names)
         ] + [
             [sg.Button("Открыть файл", key="-OPEN-", pad=((50, 10), (20, 20))),
              sg.Button("Сохранить", key="-SAVE-", pad=((10, 50), (20, 20)))]
         ]

# Создаем окно
window = sg.Window("Моя программа", layout, size=(400, 400))

while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED:
        break
    elif event == "-OPEN-":
        filename = sg.popup_get_file("Выберите файл",
                                     file_types=(("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")))
        if filename:
            sg.popup(f"Вы выбрали: {filename}")
    elif event == "-SAVE-":
        # Создаем словарь из введенных данных
        data_dict = {name: values[f"-INPUT{i}-"] for i, name in enumerate(field_names)}

        # Открываем диалоговое окно для выбора пути сохранения JSON
        save_path = sg.popup_get_file("Сохранить как", save_as=True, file_types=(("JSON файлы", "*.json"),))

        if save_path:
            # Если путь не содержит расширения, добавляем .json
            if not save_path.lower().endswith(".json"):
                save_path += ".json"

            # Записываем данные в JSON-файл
            with open(save_path, "w", encoding="utf-8") as f:
                json.dump(data_dict, f, ensure_ascii=False, indent=4)

            sg.popup(f"Данные сохранены в {save_path}")

window.close()
