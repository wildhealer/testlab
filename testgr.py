import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl
from openpyxl.styles import PatternFill
import random
import os
from datetime import datetime
import base64

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

# Функция для создания HTML-превью таблицы с цветами
def create_html_table(df, workbook, sheet_name):
    """Создаёт HTML-таблицу с учётом цвета ячеек."""
    html = "<table style='border-collapse: collapse; width: 100%;'>"
    # Заголовок
    html += "<tr style='background-color: #f2f2f2;'>"
    html += "<th style='border: 1px solid #ddd; padding: 8px;'></th>"  # Пустая ячейка для индекса
    for col in df.columns:
        html += f"<th style='border: 1px solid #ddd; padding: 8px;'>{col}</th>"
    html += "</tr>"
    
    # Данные
    for i, (index, row) in enumerate(df.iterrows()):
        html += "<tr>"
        html += f"<td style='border: 1px solid #ddd; padding: 8px; font-weight: bold;'>{index}</td>"
        for j, value in enumerate(row):
            color = get_cell_color(workbook, sheet_name, i + 2, j + 2)  # +2 из-за смещения (заголовок и индекс)
            style = "border: 1px solid #ddd; padding: 8px;"
            if color == 'red':
                style += "background-color: #ffcccc;"  # Светло-красный фон
            html += f"<td style='{style}'>{value}</td>"
        html += "</tr>"
    html += "</table>"
    return html

# Проверка наличия файла output_highlighted.xlsx и создание ссылки для скачивания
default_file = "output_highlighted.xlsx"
uploaded_file = None
if os.path.exists(default_file):
    st.markdown(f"Файл по умолчанию доступен: {get_download_link(default_file, default_file)}", unsafe_allow_html=True)
else:
    uploaded_file = st.file_uploader("Загрузите Excel файл", type=["xlsx", "xls"])

# Если файл загружен через uploader или существует файл по умолчанию
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
        
        # Выбор типа графика
        chart_type = st.selectbox("Выберите тип графика", ["Линейный", "Столбчатый", "Точечный", "Площадной"])
        
        if params:
            # Создаём объект Plotly
            fig = go.Figure()
            
            # Список цветов для линий/столбцов
            line_colors = ['blue', 'red', 'green', 'cyan', 'magenta', 'yellow', 'black']
            if len(params) > len(line_colors):
                random_colors = [f'#{random.randint(0, 255):02x}{random.randint(0, 255):02x}{random.randint(0, 255):02x}' 
                               for _ in range(len(params) - len(line_colors))]
                line_colors.extend(random_colors)
            
            # Добавляем данные для каждой характеристики
            for i, param in enumerate(params):
                color = line_colors[i % len(line_colors)]
                x_data = df.columns
                y_data = df.loc[param]
                
                if chart_type == "Линейный":
                    # Линейный график
                    fig.add_trace(go.Scatter(
                        x=x_data,
                        y=y_data,
                        mode='lines',
                        name=param,
                        line=dict(color=color, width=2)
                    ))
                    
                    # Извлекаем цвета для каждой ячейки (только для линейного и точечного)
                    point_colors = []
                    for col in range(2, len(df.columns) + 2):
                        row = df.index.get_loc(param) + 2
                        point_color = get_cell_color(wb, sheet_name, row, col)
                        point_colors.append(point_color)
                    
                    # Добавляем точки для красных ячеек
                    red_x = [x for x, pc in zip(x_data, point_colors) if pc == 'red']
                    red_y = [y for y, pc in zip(y_data, point_colors) if pc == 'red']
                    if red_x:
                        fig.add_trace(go.Scatter(
                            x=red_x,
                            y=red_y,
                            mode='markers',
                            name=f'{param} (red points)',
                            marker=dict(color='red', size=10, line=dict(color='black', width=1)),
                            showlegend=False
                        ))
                
                elif chart_type == "Столбчатый":
                    # Столбчатый график
                    fig.add_trace(go.Bar(
                        x=x_data,
                        y=y_data,
                        name=param,
                        marker_color=color,
                        width=0.8
                    ))
                
                elif chart_type == "Точечный":
                    # Точечный график
                    fig.add_trace(go.Scatter(
                        x=x_data,
                        y=y_data,
                        mode='markers',
                        name=param,
                        marker=dict(color=color, size=8, line=dict(color='black', width=1))
                    ))
                    
                    # Красные точки
                    point_colors = []
                    for col in range(2, len(df.columns) + 2):
                        row = df.index.get_loc(param) + 2
                        point_color = get_cell_color(wb, sheet_name, row, col)
                        point_colors.append(point_color)
                    
                    red_x = [x for x, pc in zip(x_data, point_colors) if pc == 'red']
                    red_y = [y for y, pc in zip(y_data, point_colors) if pc == 'red']
                    if red_x:
                        fig.add_trace(go.Scatter(
                            x=red_x,
                            y=red_y,
                            mode='markers',
                            name=f'{param} (red points)',
                            marker=dict(color='red', size=10, line=dict(color='black', width=1)),
                            showlegend=False
                        ))
                
                elif chart_type == "Площадной":
                    # Площадной график
                    fig.add_trace(go.Scatter(
                        x=x_data,
                        y=y_data,
                        mode='lines',
                        name=param,
                        fill='tozeroy',  # Заполнение до оси Y=0
                        line=dict(color=color, width=2),
                        fillcolor=f'rgba{tuple(int(color.lstrip("#")[i:i+2], 16) for i in (0, 2, 4)) + (0.3,)}'  # Полупрозрачная заливка
                    ))
            
            # Настройка осей и оформления
            fig.update_layout(
                xaxis_title="Время",
                yaxis_title="Значение",
                title="Графики выбранных характеристик",
                xaxis=dict(
                    tickangle=45,
                    tickmode='auto',
                    nticks=10 if len(df.columns) > 10 else None
                ),
                legend=dict(
                    yanchor="bottom",
                    y=-0.4,
                    xanchor="center",
                    x=0.5,
                    orientation="h",
                    font=dict(size=10)
                ),
                height=600,
                margin=dict(l=50, r=50, t=50, b=100),
                showlegend=True
            )
            
            # Добавляем основную сетку
            fig.update_layout(
                xaxis=dict(showgrid=True, gridcolor='rgba(200, 200, 200, 0.7)', griddash='dash'),
                yaxis=dict(showgrid=True, gridcolor='rgba(200, 200, 200, 0.7)', griddash='dash')
            )
            
            # Добавляем вертикальные полосы для дней
            try:
                dates = [datetime.strptime(str(x), '%d.%m %H:%M') for x in x_data]
                days = [d.date() for d in dates]
                unique_days = sorted(set(days))
                shapes = []
                for i, day in enumerate(unique_days):
                    day_indices = [j for j, d in enumerate(days) if d == day]
                    if day_indices:
                        start_idx = day_indices[0]
                        end_idx = day_indices[-1] + 1 if day_indices[-1] < len(x_data) - 1 else day_indices[-1]
                        fill_color = '#FFFFE0' if i % 2 == 0 else 'white'
                        shapes.append(dict(
                            type="rect",
                            x0=start_idx,
                            x1=end_idx,
                            y0=0,
                            y1=1,
                            yref="paper",
                            fillcolor=fill_color,
                            opacity=0.9,
                            layer="below",
                            line_width=0
                        ))
                fig.update_layout(shapes=shapes)
            except ValueError:
                shapes = []
                for i in range(len(x_data)):
                    fill_color = '#FFFFE0' if i % 2 == 0 else 'white'
                    shapes.append(dict(
                        type="rect",
                        x0=i,
                        x1=i + 1,
                        y0=0,
                        y1=1,
                        yref="paper",
                        fillcolor=fill_color,
                        opacity=0.3,
                        layer="below",
                        line_width=0
                    ))
                fig.update_layout(shapes=shapes)
            
            # Отображение графика в Streamlit
            st.plotly_chart(fig, use_container_width=True)
            
            # Превью Excel-файла
            st.subheader("Превью Excel-файла")
            html_table = create_html_table(df, wb, sheet_name)
            st.markdown(html_table, unsafe_allow_html=True)
            
        else:
            st.write("Пожалуйста, выберите хотя бы одну характеристику.")
    except Exception as e:
        st.error(f"Ошибка: {str(e)}")
