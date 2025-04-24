import tkinter as tk
from tkinter import ttk, filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import numpy as np
from scipy.interpolate import interp1d
from scipy.optimize import fsolve
import pandas as pd  # Для работы с Excel

class CurveIntersectionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("График кривых")
        self.curves = []  # Список для хранения данных кривых
        self.colors = plt.cm.tab20.colors  # Используем цветовую карту для большого количества кривых

        # Установка иконки
        try:
            self.icon = tk.PhotoImage(file="iconCat.png")  # Замените "icon.png" на имя вашего файла
            self.root.iconphoto(False, self.icon)  # Устанавливаем иконку
        except Exception as e:
            print(f"Ошибка загрузки иконки: {e}")

        # Главное окно
        self.create_main_window()

    def filter_unique_intersections(self, intersections, precision=4):
        """
        Функция для удаления дублирующихся точек пересечения.

        :param intersections: Список кортежей (x, y), представляющих точки пересечения.
        :param precision: Количество знаков после запятой для округления координат.
        :return: Множество уникальных точек пересечения.
        """
        unique_points = set()
        for point in intersections:
            x, y = point
            rounded_x = round(float(x), precision)
            rounded_y = round(float(y), precision)
            unique_points.add((rounded_x, rounded_y))
        return unique_points

    def create_main_window(self):
        # Фрейм для ввода данных
        self.input_frame = ttk.Frame(self.root)
        self.input_frame.pack(pady=10)

        # Добавить кривую
        ttk.Button(self.input_frame, text="Добавить кривую", command=self.add_curve).grid(row=0, column=0, padx=5)
        ttk.Button(self.input_frame, text="Очистить все", command=self.clear_all).grid(row=0, column=1, padx=5)

        # Кнопка для скачивания точек пересечения
        ttk.Button(self.input_frame, text="Скачать точки пересечения", command=self.download_intersections).grid(row=0, column=2, padx=5)

        # Кнопка для импорта данных из Excel
        ttk.Button(self.input_frame, text="Импортировать данные из Excel", command=self.import_from_excel).grid(row=0, column=3, padx=5)

        # Фрейм для отображения введенных данных
        self.data_frame = ttk.Frame(self.root)
        self.data_frame.pack(pady=10)

        # Кнопка для построения графика
        ttk.Button(self.root, text="Построить график", command=self.plot_graph).pack(pady=10)

    def add_curve(self):
        curve_window = tk.Toplevel(self.root)
        curve_window.title("Добавить кривую")

        # Установка иконки
        try:
            icon = tk.PhotoImage(file="iconCat.png")
            curve_window.iconphoto(False, self.icon)
        except Exception as e:
            print(f"Ошибка загрузки иконки: {e}")

        ttk.Label(curve_window, text="X (через пробел):").grid(row=0, column=0, padx=5, pady=5)
        x_entry = ttk.Entry(curve_window)
        x_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(curve_window, text="Y (через пробел):").grid(row=1, column=0, padx=5, pady=5)
        y_entry = ttk.Entry(curve_window)
        y_entry.grid(row=1, column=1, padx=5, pady=5)

        def save_curve():
            try:
                x = list(map(float, x_entry.get().split()))
                y = list(map(float, y_entry.get().split()))
                if len(x) != len(y):
                    raise ValueError("Количество точек X и Y должно совпадать.")
                self.curves.append({"x": x, "y": y})
                self.update_data_display()
                curve_window.destroy()
            except Exception as e:
                tk.messagebox.showerror("Ошибка", str(e))

        ttk.Button(curve_window, text="Сохранить", command=save_curve).grid(row=2, column=0, columnspan=2, pady=10)

    def clear_all(self):
        self.curves = []
        self.update_data_display()

    def update_data_display(self):
        for widget in self.data_frame.winfo_children():
            widget.destroy()

        for i, curve in enumerate(self.curves):
            ttk.Label(self.data_frame, text=f"Кривая {i + 1}:").grid(row=i, column=0, sticky="w")
            ttk.Label(self.data_frame, text=f"X: {curve['x']}").grid(row=i, column=1, sticky="w")
            ttk.Label(self.data_frame, text=f"Y: {curve['y']}").grid(row=i, column=2, sticky="w")

    def plot_graph(self):
        if len(self.curves) < 2:
            tk.messagebox.showwarning("Предупреждение", "Необходимо добавить хотя бы две кривые.")
            return

        # Создаем окно для графика
        graph_window = tk.Toplevel(self.root)
        graph_window.title("График кривых")
        
        progress_label = ttk.Label(graph_window, text="Загрузка...")
        progress_label.pack(pady=10)

        try:
            icon = tk.PhotoImage(file="iconCat.png")
            graph_window.iconphoto(False, self.icon)
        except Exception as e:
            print(f"Ошибка загрузки иконки: {e}")

        fig, ax = plt.subplots(figsize=(8, 6))
        intersections = []

        for i, curve in enumerate(self.curves):
            x, y = curve["x"], curve["y"]
            ax.plot(x, y, label=f"Деформация {i + 1}", color=self.colors[i % len(self.colors)])

            for j in range(i + 1, len(self.curves)):
                other_curve = self.curves[j]
                intersection_points = self.find_intersections(curve, other_curve)
                intersections.extend(intersection_points)

        unique_intersections = self.filter_unique_intersections(intersections)

        valid_points = []
        for point in unique_intersections:
            x, y = point
            if any(abs(np.interp(x, curve["x"], curve["y"]) - y) < 1e-6 for curve in self.curves):
                valid_points.append(point)

        for point in unique_intersections:
            ax.plot(point[0], point[1], 'ro')

        ax.legend()
        ax.set_xlabel("X (Сцепление по трещинам, т/м^2)")
        ax.set_ylabel("Y (Угол трения по трещинам, градусы)")
        ax.set_title("График зависимости сцепления от угла трения")
        ax.grid(True)

        canvas = FigureCanvasTkAgg(fig, master=graph_window)
        canvas.draw()
        canvas.get_tk_widget().pack()

        # Удаление лейбла после завершения
        progress_label.destroy()

        if unique_intersections:
            scroll_frame = ttk.Frame(graph_window)
            scroll_frame.pack(fill=tk.BOTH, expand=True)

            ttk.Label(scroll_frame, text="Точки пересечения", font=("Arial", 12, "bold")).pack(anchor="w", pady=5)

            scrollbar = ttk.Scrollbar(scroll_frame, orient=tk.VERTICAL)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            text_area = tk.Text(scroll_frame, yscrollcommand=scrollbar.set, wrap=tk.NONE, height=10)
            text_area.insert(tk.END, "\n".join([f"({x:.2f}, {y:.2f})" for x, y in sorted(unique_intersections)]))
            text_area.pack(fill=tk.BOTH, expand=True)

            scrollbar.config(command=text_area.yview)
        else:
            tk.messagebox.showinfo("Информация", "Точки пересечения не найдены.")

    def find_intersections(self, curve1, curve2):
        x1, y1 = curve1["x"], curve1["y"]
        x2, y2 = curve2["x"], curve2["y"]

        f1 = interp1d(x1, y1, kind='quadratic', fill_value="extrapolate")
        f2 = interp1d(x2, y2, kind='quadratic', fill_value="extrapolate")

        intersections = []
        for x in np.linspace(min(min(x1), min(x2)), max(max(x1), max(x2)), 1000):
            try:
                root = fsolve(lambda x: f1(x) - f2(x), x)[0]
                if min(x1) <= root <= max(x1) and min(x2) <= root <= max(x2):
                    tolerance = 1e-6  # Допустимая погрешность
                    if abs(f1(root) - f2(root)) < tolerance:
                        intersections.append((root, f1(root)))
            except:
                continue

        return intersections

    def download_intersections(self):
        """
        Функция для скачивания точек пересечения в Excel файл.
        """
        if not self.curves or len(self.curves) < 2:
            tk.messagebox.showwarning("Предупреждение", "Нет точек пересечения для скачивания.")
            return

        # Вычисляем все точки пересечения
        intersections = []
        for i, curve in enumerate(self.curves):
            for j in range(i + 1, len(self.curves)):
                other_curve = self.curves[j]
                intersection_points = self.find_intersections(curve, other_curve)
                intersections.extend(intersection_points)

        # Фильтруем дубликаты
        unique_intersections = self.filter_unique_intersections(intersections)

        if not unique_intersections:
            tk.messagebox.showinfo("Информация", "Точки пересечения не найдены.")
            return

        # Создаем DataFrame
        data = {"X": [point[0] for point in unique_intersections],
                "Y": [point[1] for point in unique_intersections]}
        df = pd.DataFrame(data)

        # Запрашиваем у пользователя путь для сохранения файла
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Сохранить точки пересечения"
        )

        if file_path:
            try:
                df.to_excel(file_path, index=False)
                tk.messagebox.showinfo("Успех", f"Точки пересечения успешно сохранены в:\n{file_path}")
            except Exception as e:
                tk.messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{str(e)}")

    def import_from_excel(self):
        """
        Функция для импорта данных из Excel файла.
        """
        # Открываем диалоговое окно для выбора файла Excel
        file_path = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Выберите файл Excel"
        )

        if not file_path:
            return  # Если пользователь отменил выбор файла

        try:
            # Читаем данные из Excel файла
            df = pd.read_excel(file_path)

            # Проверяем количество столбцов
            num_columns = len(df.columns)
            if num_columns % 2 != 0:
                tk.messagebox.showerror("Ошибка", "Файл должен содержать четное количество столбцов (пары X и Y).")
                return

            # Очищаем текущие кривые
            self.curves = []

            # Обрабатываем столбцы парами (X, Y)
            for i in range(0, num_columns, 2):
                x_column = df.iloc[:, i].dropna().tolist()  # Столбец X
                y_column = df.iloc[:, i + 1].dropna().tolist()  # Столбец Y

                # Округляем значения до двух знаков после запятой
                x_column_rounded = [round(float(x), 2) for x in x_column]
                y_column_rounded = [round(float(y), 2) for y in y_column]

                if len(x_column_rounded) != len(y_column_rounded):
                    tk.messagebox.showerror("Ошибка", f"Количество значений X и Y в столбцах {i + 1} и {i + 2} не совпадает.")
                    return

                # Добавляем кривую
                self.curves.append({"x": x_column_rounded, "y": y_column_rounded})

            # Обновляем отображение данных
            self.update_data_display()

            tk.messagebox.showinfo("Успех", f"Данные успешно импортированы из:\n{file_path}")
        except Exception as e:
            tk.messagebox.showerror("Ошибка", f"Не удалось импортировать данные:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = CurveIntersectionApp(root)
    root.mainloop()