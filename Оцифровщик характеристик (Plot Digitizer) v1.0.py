import sys
import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from PIL import Image, ImageDraw, ImageTk
import matplotlib
import numpy as np
import pandas as pd

# Set matplotlib backend before other imports
matplotlib.use('TkAgg')
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt


def resource_path(relative_path):
    """Get absolute path to resource for PyInstaller"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


class GraphDigitizer:
    def __init__(self, root):
        self.root = root
        self.root.title("Оцифровщик характеристик. ©ГЦЭ-энерго. Надежность")

        # Load application icon
        self.load_application_icon()

        # Initialize matplotlib components
        self.fig, self.ax = plt.subplots()
        self.ax.axis('off')
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.root)
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

        # Create scope area (right sidebar)
        self.scope_frame = tk.Frame(self.root)
        self.scope_label = tk.Label(self.scope_frame)
        self.scope_label.pack(pady=10)
        self.scope_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10)

        # Data storage
        self.reset_program()

        # UI Setup
        self.create_widgets()
        self.fig.canvas.mpl_connect('button_press_event', self.on_click)
        self.root.geometry("1280x1024")
        self.canvas.mpl_connect('motion_notify_event', self.show_scope)

    def reset_program(self):
        """Reset all variables and UI to initial state"""
        self.image_path = None
        self.image = None
        self.points = []
        self.axis_points = []
        self.data_points = []
        self.datasets = []
        self.current_graph_number = 1
        self.x_axis_xp = np.array([])
        self.x_axis_fp = np.array([])
        self.y_axis_xp = np.array([])
        self.y_axis_fp = np.array([])

        # Reset matplotlib display
        self.ax.clear()
        self.ax.axis('off')
        self.ax.set_title("Для выбора картинки нажмите 'Загрузить изображение'")
        self.canvas.draw()

        # Reset UI buttons
        if hasattr(self, 'save_button'):
            self.save_button.config(state=tk.DISABLED)
            self.finish_button.config(state=tk.DISABLED)

    def load_application_icon(self):
        """Handle application icon loading with fallback"""
        try:
            icon_path = resource_path('sova.ico')
            self.root.iconbitmap(icon_path)
        except tk.TclError:
            try:
                png_path = resource_path('sova.png')
                img = Image.open(png_path)
                photo = ImageTk.PhotoImage(img)
                self.root.iconphoto(True, photo)
            except Exception as e:
                messagebox.showwarning("Предупреждение", f"Иконки не найдены: {str(e)}")

    def create_widgets(self):
        """Initialize UI components"""
        control_frame = tk.Frame(self.root)
        control_frame.pack(side=tk.BOTTOM, fill=tk.X)

        self.new_image_button = tk.Button(control_frame, text="Новое изображение",
                                          command=self.new_image_session)
        self.new_image_button.pack(side=tk.LEFT, padx=5)

        self.load_button = tk.Button(control_frame, text="Загрузить изображение",
                                     command=self.load_image)
        self.load_button.pack(side=tk.LEFT, padx=5)

        self.save_button = tk.Button(control_frame, text="Сохранить в Excel",
                                     command=self.save_data, state=tk.DISABLED)
        self.save_button.pack(side=tk.LEFT, padx=5)

        self.finish_button = tk.Button(control_frame, text="Точки выбраны",
                                       command=self.finish_digitizing, state=tk.DISABLED)
        self.finish_button.pack(side=tk.LEFT, padx=5)

        self.cancel_button = tk.Button(control_frame, text="Отменить выбор",
                                       command=self.cancel_last_selection)
        self.cancel_button.pack(side=tk.LEFT, padx=5)

    def new_image_session(self):
        """Start new session with confirmation"""
        if self.datasets or self.data_points:
            response = messagebox.askyesno(
                "Новая сессия",
                "Вы уверены что хотите начать новую сессию? Текущие данные будут потеряны!"
            )
            if not response:
                return
        self.reset_program()

    def load_image(self):
        """Load and display the selected image."""
        self.reset_program()
        self.image_path = filedialog.askopenfilename(
            filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.webp")]
        )
        if self.image_path:
            self.image = Image.open(self.image_path)
            self.refresh_plot()
            self.ax.set_title(f"Установите оси для графика {self.current_graph_number}")

    def refresh_plot(self):
        self.ax.clear()
        self.ax.axis('off')
        if self.image:
            self.ax.imshow(self.image)
        for point in self.axis_points:
            self.ax.plot(point[0], point[1], 'ro')
        self.canvas.draw()
        self.show_scope(None)

    def show_scope(self, event):
        if self.image and (event is None or event.xdata and event.ydata):
            x = int(event.xdata) if event else self.image.width // 2
            y = int(event.ydata) if event else self.image.height // 2

            start_x = max(0, x - 50)
            start_y = max(0, y - 50)
            end_x = min(self.image.width, x + 50)
            end_y = min(self.image.height, y + 50)

            scope_image = self.image.crop((start_x, start_y, end_x, end_y))
            scope_image = scope_image.resize((200, 200), Image.BILINEAR)

            draw = ImageDraw.Draw(scope_image)
            draw.line((100, 0, 100, 200), fill='red')
            draw.line((0, 100, 200, 100), fill='red')

            photo = ImageTk.PhotoImage(scope_image)
            self.scope_label.config(image=photo)
            self.scope_label.image = photo

    def on_click(self, event):
        if event.inaxes != self.ax:
            return

        if event.button == 3:  # Right-click
            self.cancel_last_selection()
            return

        if len(self.axis_points) < 4:
            self.handle_axis_selection(event)
        else:
            self.handle_data_selection(event)

    def handle_axis_selection(self, event):
        self.axis_points.append((event.xdata, event.ydata))
        self.ax.plot(event.xdata, event.ydata, 'ro')
        self.update_axis_stage()
        self.canvas.draw()

    def update_axis_stage(self):
        if len(self.axis_points) == 2:
            self.ax.set_title("Теперь Y1, Y2")
        elif len(self.axis_points) == 4:
            if self.define_axes():
                self.ax.set_title(f"Определите точки для графика {self.current_graph_number}")
                self.finish_button.config(state=tk.NORMAL)
                self.plot_axis_lines()
            else:
                self.axis_points = []
                self.refresh_plot()

    def plot_axis_lines(self):
        self.ax.plot(
            [self.axis_points[0][0], self.axis_points[1][0]],
            [self.axis_points[0][1], self.axis_points[1][1]], 'r-'
        )
        self.ax.plot(
            [self.axis_points[2][0], self.axis_points[3][0]],
            [self.axis_points[2][1], self.axis_points[3][1]], 'r-'
        )
        self.canvas.draw()

    def handle_data_selection(self, event):
        self.data_points.append((event.xdata, event.ydata))
        self.ax.plot(event.xdata, event.ydata, 'bo')
        self.canvas.draw()

    def define_axes(self):
        try:
            # X-axis calibration
            x_pairs = [
                (self.axis_points[0][0], self.get_axis_value("X1", self.axis_points[0][0], True)),
                (self.axis_points[1][0], self.get_axis_value("X2", self.axis_points[1][0], True))
            ]
            x_sorted = sorted(x_pairs, key=lambda x: x[0])
            self.x_axis_xp = np.array([x[0] for x in x_sorted])
            self.x_axis_fp = np.array([x[1] for x in x_sorted])

            # Y-axis calibration
            y_pairs = [
                (self.axis_points[2][1], self.get_axis_value("Y1", self.axis_points[2][1], False)),
                (self.axis_points[3][1], self.get_axis_value("Y2", self.axis_points[3][1], False))
            ]
            y_sorted = sorted(y_pairs, key=lambda y: y[0])
            self.y_axis_xp = np.array([y[0] for y in y_sorted])
            self.y_axis_fp = np.array([y[1] for y in y_sorted])
            return True
        except (TypeError, ValueError):
            return False

    def get_axis_value(self, label, pixel_value, is_x):
        while True:
            prompt = f"Введите величину {label} (Пиксель {'X' if is_x else 'Y'}: {pixel_value:.2f})"
            value = simpledialog.askfloat("Калибровка осей", prompt, parent=self.root)
            if value is not None:
                return value
            messagebox.showwarning("Ошибка", "Введите корректное число!")

    def pixel_to_actual(self, pixel_x, pixel_y):
        try:
            x = np.interp(pixel_x, self.x_axis_xp, self.x_axis_fp)
            y = np.interp(pixel_y, self.y_axis_xp, self.y_axis_fp)
            return round(x, 5), round(y, 5)
        except ValueError as e:
            messagebox.showerror("Ошибка", f"Ошибка преобразования: {str(e)}")
            return None, None

    def cancel_last_selection(self):
        if len(self.axis_points) < 4:
            if self.axis_points:
                self.axis_points.pop()
                self.refresh_plot()
        else:
            if self.data_points:
                self.data_points.pop()
                self.refresh_plot()
        self.refresh_buttons_state()

    def refresh_buttons_state(self):
        self.save_button.config(state=tk.NORMAL if self.datasets else tk.DISABLED)
        self.finish_button.config(state=tk.NORMAL if len(self.axis_points) == 4 else tk.DISABLED)

    def finish_digitizing(self):
        try:
            # Convert and sort points
            raw_points = [self.pixel_to_actual(x, y) for x, y in self.data_points]
            clean_points = [(x, y) for x, y in raw_points if x is not None and y is not None]

            # Add explicit sorting before DataFrame creation
            clean_points = sorted(clean_points, key=lambda point: point[0])  # <-- NEW LINE

            if not clean_points:
                messagebox.showwarning("Пустые данные", "Нет корректных точек для обработки!")
                return

            # Create DataFrame and ensure sorted ascending X values
            current_data = pd.DataFrame(
                clean_points,
                columns=[f'График {self.current_graph_number} X',
                         f'График {self.current_graph_number} Y']
            ).sort_values(by=f'График {self.current_graph_number} X', ascending=True)

            # Remove duplicates and ensure strict increase
            current_data = current_data.drop_duplicates(
                subset=[f'График {self.current_graph_number} X'],
                keep='first'
            )
            current_data = current_data[
                current_data[f'График {self.current_graph_number} X'].diff().fillna(1) > 0
                ]

            if len(current_data) < 2:
                messagebox.showwarning("Мало точек", "Нужно как минимум 2 уникальные точки по X!")
                return

            x_orig = current_data[f'График {self.current_graph_number} X'].values
            y_orig = current_data[f'График {self.current_graph_number} Y'].values

            # Polynomial regression (degree 2-5 based on points)
            max_degree = min(5, len(x_orig) - 1)
            degree = max(2, max_degree)  # Minimum 2nd degree polynomial

            coeffs = np.polyfit(x_orig, y_orig, deg=degree)
            poly = np.poly1d(coeffs)

            # Generate smooth curve
            x_interp = np.linspace(x_orig.min(), x_orig.max(), 100)
            y_interp = poly(x_interp)

            # Create output DataFrames
            interp_data = pd.DataFrame({
                f'График {self.current_graph_number} Интерп. X': x_interp,
                f'График {self.current_graph_number} Интерп. Y': y_interp
            })

            # Combine data with original first
            combined = pd.concat([current_data, interp_data], axis=1)
            self.datasets.append(combined)

            # Update UI
            # Update UI
            self.current_graph_number += 1
            self.ax.set_title("Оцифровано! Сохраните или добавьте еще график")
            self.save_button.config(state=tk.NORMAL)

            # New code: Ask about axis retention
            keep_axes = messagebox.askyesno(
                "Сохранение осей",
                "Хотите использовать те же оси для следующего графика?\n"
                "(Да - оставить текущие оси, Нет - сбросить настройки осей)"
            )

            if keep_axes:
                # Keep axis calibration, clear only data points
                self.data_points = []
                # Redraw existing axis lines
                self.plot_axis_lines()
                self.ax.set_title(f"Определите точки для графика {self.current_graph_number} (те же оси)")
            else:
                # Full reset
                self.axis_points = []
                self.x_axis_xp = np.array([])
                self.x_axis_fp = np.array([])
                self.y_axis_xp = np.array([])
                self.y_axis_fp = np.array([])
                self.data_points = []
                self.refresh_plot()
                self.ax.set_title(f"Установите оси для графика {self.current_graph_number}")

            self.refresh_plot()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка обработки: {str(e)}")

    def save_data(self):
        if not self.datasets:
            messagebox.showwarning("Пустые данные", "Нет данных для сохранения!")
            return

        combined_data = pd.concat(self.datasets, axis=1)

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not save_path:
            return

        with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
            combined_data.to_excel(writer, sheet_name='Данные', index=False)

            workbook = writer.book
            worksheet = writer.sheets['Данные']

            # Formatting
            header_format = workbook.add_format({
                'font_name': 'Times New Roman',
                'font_size': 10,
                'bold': True,
                'align': 'center'
            })
            cell_format = workbook.add_format({
                'font_name': 'Times New Roman',
                'font_size': 10
            })

            for col in range(combined_data.shape[1]):
                worksheet.write(0, col, combined_data.columns[col], header_format)
                worksheet.set_column(col, col, None, cell_format)

            # Create two separate charts
            raw_chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
            interp_chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})

            # Get actual data length
            max_row = len(combined_data)

            # Plot all raw data points with proper ranges
            for i in range(0, combined_data.shape[1], 4):
                # Raw data series (original points)
                raw_chart.add_series({
                    'name': f'График {i // 4 + 1} (исходный)',
                    'categories': ['Данные', 1, i, max_row, i],  # X-values
                    'values': ['Данные', 1, i + 1, max_row, i + 1],  # Y-values
                    'marker': {
                        'type': 'circle',
                        'size': 5,
                        'border': {'color': 'black'},
                        'fill': {'color': '#4F81BD'}  # Blue color for visibility
                    },
                    'line': {'none': True}
                })

            # Plot all interpolated curves with separate ranges
            for i in range(2, combined_data.shape[1], 4):
                # Interpolated data series
                interp_chart.add_series({
                    'name': f'График {i // 4 + 1} (интерполяция)',
                    'categories': ['Данные', 1, i, max_row, i],  # X-values
                    'values': ['Данные', 1, i + 1, max_row, i + 1],  # Y-values
                    'marker': {'type': 'none'},
                    'line': {'width': 1.5, 'color': '#C0504D'}  # Red color for contrast
                })

            # Configure charts independently
            raw_chart.set_title({'name': 'Исходные данные', 'name_font': {'name': 'Times New Roman', 'size': 14}})
            raw_chart.set_x_axis({'name': 'X', 'name_font': {'name': 'Times New Roman', 'size': 12}})
            raw_chart.set_y_axis({'name': 'Y', 'name_font': {'name': 'Times New Roman', 'size': 12}})
            raw_chart.set_legend({'position': 'bottom', 'font': {'name': 'Times New Roman'}})

            interp_chart.set_title(
                {'name': 'Интерполированные данные', 'name_font': {'name': 'Times New Roman', 'size': 14}})
            interp_chart.set_x_axis({'name': 'X', 'name_font': {'name': 'Times New Roman', 'size': 12}})
            interp_chart.set_y_axis({'name': 'Y', 'name_font': {'name': 'Times New Roman', 'size': 12}})
            interp_chart.set_legend({'position': 'bottom', 'font': {'name': 'Times New Roman'}})

            # Insert charts with proper spacing
            worksheet.insert_chart('D2', raw_chart, {'x_offset': 25, 'y_offset': 10, 'x_scale': 1.5, 'y_scale': 1.5})
            worksheet.insert_chart('D20', interp_chart,
                                   {'x_offset': 25, 'y_offset': 10, 'x_scale': 1.5, 'y_scale': 1.5})

        messagebox.showinfo("Успех", f"Данные сохранены в:\n{save_path}")

        # Post-save options
        continue_response = messagebox.askyesno(
            "Продолжить работу",
            "Хотите продолжить работу с этим изображением?"
        )
        if continue_response:
            keep_axes = messagebox.askyesno(
                "Настройки осей",
                "Использовать текущие оси для следующего графика?\n"
                "(Да - использовать те же оси, Нет - задать новые)"
            )
            if keep_axes:
                self.data_points = []
                self.plot_axis_lines()
                self.ax.set_title(f"Определите точки для графика {self.current_graph_number} (те же оси)")
            else:
                self.axis_points = []
                self.x_axis_xp = np.array([])
                self.x_axis_fp = np.array([])
                self.y_axis_xp = np.array([])
                self.y_axis_fp = np.array([])
                self.data_points = []
                self.refresh_plot()
                self.ax.set_title(f"Установите оси для графика {self.current_graph_number}")
        else:
            self.new_image_session()

if __name__ == "__main__":
    root = tk.Tk()
    app = GraphDigitizer(root)
    root.mainloop()