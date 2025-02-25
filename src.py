import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import json

def select_file():
    """Функция для выбора xls/xlsx-файла"""
    filepath = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
    )
    if filepath:
        input_file_label.config(text=filepath)
        global input_file_path
        input_file_path = filepath

def process_file():
    """Функция для обработки данных"""
    try:
        # Проверяем, выбран ли файл
        if not input_file_path:
            messagebox.showerror("Ошибка", "Выберите входной файл!")
            return

        # Получение имени столбца из интерфейса
        column_name = column_entry.get()

        if not column_name:
            messagebox.showerror("Ошибка", "Введите имя столбца!")
            return

        # Чтение Excel файла
        df = pd.read_excel(input_file_path)

        if column_name not in df.columns:
            messagebox.showerror("Ошибка", f"Столбец '{column_name}' не найден!")
            return

        # Функция для извлечения данных из строки JSON
        def extract_value(json_string, key):
            try:
                data = json.loads(json_string)  # Преобразуем строку в объект JSON
                value = data.get(key, "")  # Получаем значение по ключу

                # Если выбран ключ "id", удаляем префикс "acc_", если он есть
                if key == "id" and value.startswith("acc_"):
                    return value.replace("acc_", "")

                return value
            except json.JSONDecodeError:
                return ""  # Если строка невалидна, вернуть пустую строку

        # Извлечение значения из столбца
        key_to_extract = key_var.get()
        df["Результат"] = df[column_name].astype(str).apply(lambda x: extract_value(x, key_to_extract))

        # Удаление всех других столбцов, кроме результата
        df = df[["Результат"]]

        # Сохранение обработанных данных в новый Excel файл
        output_file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Сохранить обработанный файл как",
        )
        if output_file_path:
            df.to_excel(output_file_path, index=False)
            messagebox.showinfo("Готово", f"Файл успешно сохранен: {output_file_path}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")

# Создаем окно приложения
app = tk.Tk()
app.title("Обработка JSON из Excel")

# Интерфейс выбора файла
tk.Label(app, text="Выберите входной Excel файл:").pack(pady=5)
input_file_label = tk.Label(app, text="Файл не выбран", fg="red")
input_file_label.pack()
tk.Button(app, text="Выбрать файл", command=select_file).pack(pady=5)

# Поле для ввода имени столбца
tk.Label(app, text="Введите имя столбца с данными:").pack(pady=5)
column_entry = tk.Entry(app)
column_entry.pack(pady=5)

# Опция для выбора ключа
key_var = tk.StringVar(value="sn")
tk.Radiobutton(app, text="Извлечь sn", variable=key_var, value="sn").pack()
tk.Radiobutton(app, text="Извлечь id", variable=key_var, value="id").pack()

# Кнопка обработки
tk.Button(app, text="Обработать данные", command=process_file).pack(pady=20)

# Запуск приложения
input_file_path = None
app.mainloop()

