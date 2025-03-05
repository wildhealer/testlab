import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import AutoMinorLocator
import openpyxl
from openpyxl.styles import PatternFill
import random

st.title("График из Excel с несколькими характеристиками и точками для красных ячеек")

# Функция для определения цвета ячейки с точной обработкой ARGB
def get_cell_color(workbook, sheet_name, row, col):
    """Извлекает цвет заливки ячейки из Excel-файла с учётом формата ARGB."""
    try:
        worksheet = workbook[sheet_name]
        cell = worksheet.cell(row=row, column=col)
        fill = cell.fill
        
        st.write(f"Row: {row}, Col: {col}, Fill type: {type(fill)}, Fill: {fill}")
        
       # if isinstance(fill, PatternFill) and fill.fill_type == 'solid':
        if isinstance(fill, openpyxl.styles.fills.PatternFill):    
            rgb = fill.fgColor.rgb
            if rgb:
                # Проверяем формат ARGB (например, 00FF0000) или HEX
                st.write(f"Raw RGB value: {rgb}")
                # Преобразуем в нижний регистр для надёжности
                rgb_lower = rgb.lower()
                
                # Проверяем, красный ли цвет (00FF0000 или FF0000)
                if rgb_lower == 'ff993300':  # ARGB формат
                    st.write("Detected red color (ARGB 00FF0000)")
                    return 'red'
                # Проверяем HEX без префикса (FF0000)
                elif rgb_lower == 'ff0000':  # HEX без префикса FF
                    st.write("Detected red color (HEX FF0000)")
                    return 'red'
                # Проверяем полный HEX с префиксом (FF0000FF)
                elif rgb_lower == 'ff0000ff':  # Полный HEX
                    st.write("Detected red color (HEX FF0000FF)")
                    return 'red'
            else:
                st.write("No RGB color found in fill.")
        else:
            st.write("Fill is not PatternFill or not solid.")
        return None  # Возвращаем None, если цвет не красный
    except Exception as e:
        st.error(f"Ошибка при чтении цвета ячейки: {str(e)}")
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
        
        # Выводим структуру DataFrame для отладки
        st.write("DataFrame structure:")
        st.write(df)
        
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
                
                # Проверяем конкретно для "Экзистенциальные приключения Пятачка"
                if param == "Экзистенциальные приключения Пятачка":
                    st.write(f"Processing param: {param}, X-data: {x_data}")
                
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
                        st.write(f"Red point at X: {x}, Y: {y}, Row: {row}, Col: {col}")
            
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
