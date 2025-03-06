import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import AutoMinorLocator
import openpyxl
from openpyxl.styles import PatternFill
import random

st.title("График из Excel с несколькими характеристиками и точками для красных ячеек")

# Функция для определения цвета ячейки
def get_cell_color(workbook, sheet_name, row, col):
    """Извлекает цвет заливки ячейки из Excel-файла с учётом StyleProxy и формата ARGB."""
    try:
        worksheet = workbook[sheet_name]
        cell = worksheet.cell(row=row, column=col)
        fill = cell.fill
        
        # Пытаемся получить доступ к внутреннему объекту fill
        if hasattr(fill, 'fill'):
            actual_fill = fill.fill
        else:
            actual_fill = fill
        
        # Прямой доступ к rgb
        if hasattr(actual_fill, 'fgColor') and hasattr(actual_fill.fgColor, 'rgb'):
            rgb = actual_fill.fgColor.rgb
            if rgb:
                # Преобразуем в нижний регистр для надёжности
                rgb_lower = rgb.lower()
                
                # Проверяем, красный ли цвет (FFFF0000, 00FF0000, FF0000)
                if rgb_lower in ['ffff0000', '00ff0000']:  # ARGB форматы красного
                    return 'red'
                elif rgb_lower in ['ff0000', 'ff0000ff']:  # HEX форматы красного
                    return 'red'
        return None  # Возвращаем None, если цвет не красный
    except Exception:
        return None

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
            
            # Список цветов для линий
            line_colors = ['b', 'r', 'g', 'c', 'm', 'y', 'k']
            if len(params) > len(line_colors):
                # Если характеристик больше, чем цветов, генерируем случайные цвета
                random_colors = [f'#{random.randint(0, 255):02x}{random.randint(0, 255):02x}{random.randint(0, 255):02x}' 
                               for _ in range(len(params) - len(line_colors))]
                line_colors.extend(random_colors)
            
            # Построение графиков для каждой выбранной характеристики
            for i, param in enumerate(params):
                color = line_colors[i % len(line_colors)]  # Цвет линии
                x_data = df.columns
                y_data = df.loc[param]
                
                # Рисуем линию
                ax.plot(x_data, y_data, color=color, label=param, linewidth=2)
                
                # Извлекаем цвета для каждой ячейки в строке характеристики
                point_colors = []
                for col in range(2, len(df.columns) + 2):  # Начинаем со второго столбца (индекс 2)
                    row = df.index.get_loc(param) + 2  # Номер строки (начиная с 2, т.к. первая строка — заголовки)
                    point_color = get_cell_color(wb, sheet_name, row, col)
                    point_colors.append(point_color)
                
                # Рисуем точки только для красных ячеек
                for x, y, point_color in zip(x_data, y_data, point_colors):
                    if point_color == 'red':  # Рисуем точку только если ячейка красная
                        ax.scatter(x, y, color='red', s=50, edgecolor='black', zorder=5)
            
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
