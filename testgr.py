import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import AutoMinorLocator
import openpyxl
from openpyxl.styles import PatternFill
import random

st.title("График из Excel с несколькими характеристиками и цветными точками")

# Функция для определения цвета ячейки
def get_cell_color(workbook, sheet_name, row, col):
    """Извлекает цвет заливки ячейки из Excel-файла."""
    try:
        worksheet = workbook[sheet_name]
        cell = worksheet.cell(row=row, column=col)
        fill = cell.fill
        if isinstance(fill, PatternFill) and fill.fill_type == 'solid':
            rgb = fill.fgColor.rgb
            if rgb:
                # Удаляем префикс 'FF' (если есть) и преобразуем в HEX
                hex_color = f'#{rgb[2:] if rgb.startswith("FF") else rgb}'
                # Проверяем основные цвета (красный, жёлтый)
                if hex_color.lower() in ['#ff0000', '#ff0000ff']:  # Красный
                    return 'red'
                elif hex_color.lower() in ['#ffff00', '#ffff00ff']:  # Жёлтый
                    return 'yellow'
                else:
                    return hex_color  # Возвращаем HEX, если цвет не красный/жёлтый
        return 'blue'  # Значение по умолчанию, если цвет не определён
    except Exception as e:
        st.error(f"Ошибка при чтении цвета ячейки: {str(e)}")
        return 'blue'

# Загрузка файла
uploaded_file = st.file_uploader("Загрузите Excel файл", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Читаем данные с помощью pandas
        df = pd.read_excel(uploaded_file)
        # Устанавливаем первый столбец как индекс
        df.set_index(df.columns[0], inplace=True)
        
        # Открываем файл с помощью openpyxl для чтения цветов
        wb = openpyxl.load_workbook(uploaded_file)
        sheet_name = wb.sheetnames[0]  # Берем первый лист
        
        # Выбор нескольких характеристик
        params = st.multiselect("Выберите характеристики", df.index.tolist())
        
        if params:  # Если выбраны хотя бы одна характеристика
            # Построение графика
            fig, ax = plt.subplots(figsize=(12, 6))  # Увеличиваем размер для нескольких линий
            
            # Список цветов для линий (если цвет ячейки не определён)
            line_colors = ['b', 'r', 'g', 'c', 'm', 'y', 'k']
            if len(params) > len(line_colors):
                random_colors = [f'#{random.randint(0, 255):02x}{random.randint(0, 255):02x}{random.randint(0, 255):02x}' 
                               for _ in range(len(params) - len(line_colors))]
                line_colors.extend(random_colors)
            
            # Построение графиков для каждой выбранной характеристики с цветными точками
            for i, param in enumerate(params):
                color = line_colors[i % len(line_colors)]  # Цвет линии по умолчанию
                point_colors = []  # Список цветов точек для этой характеристики
                
                # Извлекаем цвета для каждой ячейки в строке характеристики
                for col in range(2, len(df.columns) + 2):  # Начинаем со второго столбца (индекс 2)
                    row = df.index.get_loc(param) + 2  # Номер строки (начиная с 2, т.к. первая строка — заголовки)
                    point_color = get_cell_color(wb, sheet_name, row, col)
                    point_colors.append(point_color)
                
                # Построение линии с точками, где цвет точек зависит от цвета ячейки
                x_data = df.columns
                y_data = df.loc[param]
                
                # Рисуем линию
                ax.plot(x_data, y_data, color=color, label=param, linewidth=2)
                
                # Рисуем точки с разными цветами
                for x, y, point_color in zip(x_data, y_data, point_colors):
                    ax.scatter(x, y, color=point_color, s=50, edgecolor='black', zorder=5)
            
            # Настройка меток осей
            ax.set_xlabel("Время")
            ax.set_ylabel("Значение")
            ax.set_title("Графики выбранных характеристик")
            
            # Ротация меток на оси X и настройка интервала
            plt.xticks(rotation=45, ha="right")
            
            # Автоматическая подгонка меток, чтобы избежать наложения
            fig.tight_layout()
            
            # Ограничение количества меток, если их слишком много
            if len(df.columns) > 10:
                ax.xaxis.set_major_locator(plt.MaxNLocator(nbins=10))
            
            # Добавляем легенду и сетку
            ax.legend()
            ax.grid(True, linestyle='--', alpha=0.7)
            
            # Отображение графика
            st.pyplot(fig)
        else:
            st.write("Пожалуйста, выберите хотя бы одну характеристику.")
    except Exception as e:
        st.error(f"Ошибка: {str(e)}")
