import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import AutoMinorLocator
import openpyxl
from openpyxl.styles import PatternFill
import random
import os
from datetime import datetime
import base64  # Для создания ссылки на скачивание
import mpld3  # Для интерактивного зума
from streamlit.components.v1 import html  # Для отображения интерактивного графика

st.title("График из Excel с несколькими характеристиками и точками для красных ячеек")

# Функция для определения цвета ячейки
def get_cell_color(workbook, sheet_name, row, col):
    """Извлекает цвет заливки ячейки из Excel-файла с учётом StyleProxy и формата ARGB."""
    try:
        worksheet = workbook[sheet_name]
        cell = worksheet.cell(row=row, column=col)
        fill = cell.fill
        
        if hasattr(fill, 'fill'):
            actual_fill = fill.fill
        else:
            actual_fill = fill
        
        if hasattr(actual_fill, 'fgColor') and hasattr(actual_fill.fgColor, 'rgb'):
            rgb = actual_fill.fgColor.rgb
            if rgb:
                rgb_lower = rgb.lower()
                if rgb_lower in ['ffff0000', '00ff0000', 'ff0000', 'ff0000ff']:
                    return 'red'
        return None
    except Exception:
        return None

# Функция для создания ссылки на скачивание файла
def get_download_link(file_path, file_name):
    """Создаёт ссылку для скачивания файла."""
    with open(file_path, "rb") as f:
        bytes_data = f.read()
    b64 = base64.b64encode(bytes_data).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}">Скачать {file_name}</a>'
    return href

# Проверка наличия файла output_highlighted.xlsx и создание ссылки для скачивания
default_file = "output_highlighted.xlsx"
uploaded_file = None
if os.path.exists(default_file):
    st.markdown(f"Файл по умолчанию доступен: {get_download_link(default_file, default_file)}", unsafe_allow_html=True)
else:
    uploaded_file = st.file_uploader("Загрузите Excel файл", type=["xlsx", "xls"])

# Если файл загружен через uploader, используем его
if uploaded_file is not None or os.path.exists(default_file):
    try:
        # Читаем данные с помощью pandas
        if uploaded_file is None and os.path.exists(default_file):
            df = pd.read_excel(default_file)
            wb = openpyxl.load_workbook(default_file)
        else:
            df = pd.read_excel(uploaded_file)
            from io import BytesIO
            wb = openpyxl.load_workbook(BytesIO(uploaded_file.read()))
        
        df.set_index(df.columns[0], inplace=True)
        sheet_name = wb.sheetnames[0]
        
        # Выбор нескольких характеристик
        params = st.multiselect("Выберите характеристики", df.index.tolist())
        
        if params:
            # Построение графика
            fig, ax = plt.subplots(figsize=(12, 6))
            line_colors = ['b', 'r', 'g', 'c', 'm', 'y', 'k']
            if len(params) > len(line_colors):
                random_colors = [f'#{random.randint(0, 255):02x}{random.randint(0, 255):02x}{random.randint(0, 255):02x}' 
                               for _ in range(len(params) - len(line_colors))]
                line_colors.extend(random_colors)
            
            for i, param in enumerate(params):
                color = line_colors[i % len(line_colors)]
                x_data = df.columns
                y_data = df.loc[param]
                ax.plot(x_data, y_data, color=color, label=param, linewidth=2)
                
                point_colors = []
                for col in range(2, len(df.columns) + 2):
                    row = df.index.get_loc(param) + 2
                    point_color = get_cell_color(wb, sheet_name, row, col)
                    point_colors.append(point_color)
                
                for x, y, point_color in zip(x_data, y_data, point_colors):
                    if point_color == 'red':
                        ax.scatter(x, y, color='red', s=50, edgecolor='black', zorder=5)
            
            # Настройка осей и оформления
            ax.set_xlabel("Время")
            ax.set_ylabel("Значение")
            ax.set_title("Графики выбранных характеристик")
            plt.xticks(rotation=45, ha="right")
            fig.tight_layout()
            if len(df.columns) > 10:
                ax.xaxis.set_major_locator(plt.MaxNLocator(nbins=10))
            ax.legend(loc='lower center', bbox_to_anchor=(0.5, -0.4), ncol=3, fontsize='small')
            ax.minorticks_on()
            ax.grid(which='major', linestyle='--', alpha=0.7)
            ax.grid(which='minor', linestyle=':', alpha=0.2)
            
            # Добавляем вертикальные полосы для дней
            try:
                dates = [datetime.strptime(str(x), '%d.%m %H:%M') for x in x_data]
                days = [d.date() for d in dates]
                unique_days = sorted(set(days))
                for i, day in enumerate(unique_days):
                    day_indices = [j for j, d in enumerate(days) if d == day]
                    if day_indices:
                        start_idx = day_indices[0]
                        end_idx = day_indices[-1] + 1 if day_indices[-1] < len(x_data) - 1 else day_indices[-1]
                        fill_color = '#FFFFE0' if i % 2 == 0 else 'white'
                        ax.axvspan(start_idx, end_idx, facecolor=fill_color, alpha=0.9, zorder=0)
            except ValueError:
                for i in range(len(x_data)):
                    fill_color = '#FFFFE0' if i % 2 == 0 else 'white'
                    ax.axvspan(i, i + 1, facecolor=fill_color, alpha=0.3, zorder=0)
            
            # Преобразуем график в интерактивный HTML с помощью mpld3
            html_graph = mpld3.fig_to_html(fig, template_type="simple")
            st.write("Используйте колесо мыши или жесты для масштабирования графика:")
            html(html_graph, height=600, scrolling=True)
            
        else:
            st.write("Пожалуйста, выберите хотя бы одну характеристику.")
    except Exception as e:
        st.error(f"Ошибка: {str(e)}")
