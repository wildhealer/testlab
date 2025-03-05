import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import AutoMinorLocator
import random

st.title("График из Excel с несколькими характеристиками")

# Загрузка файла
uploaded_file = st.file_uploader("Загрузите Excel файл", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Читаем файл
        df = pd.read_excel(uploaded_file)
        # Устанавливаем первый столбец как индекс
        df.set_index(df.columns[0], inplace=True)
        
        # Выбор нескольких характеристик
        params = st.multiselect("Выберите характеристики", df.index.tolist())
        
        if params:  # Если выбраны хотя бы одна характеристика
            # Построение графика
            fig, ax = plt.subplots(figsize=(12, 6))  # Увеличиваем размер для нескольких линий
            
            # Список цветов для разных линий
            colors = ['b', 'r', 'g', 'c', 'm', 'y', 'k']  # Базовые цвета
            if len(params) > len(colors):
                # Если характеристик больше, чем цветов, генерируем случайные цвета
                random_colors = [f'#{random.randint(0, 255):02x}{random.randint(0, 255):02x}{random.randint(0, 255):02x}' 
                               for _ in range(len(params) - len(colors))]
                colors.extend(random_colors)
            
            # Построение графиков для каждой выбранной характеристики
            for i, param in enumerate(params):
                ax.plot(df.columns, df.loc[param], color=colors[i % len(colors)], label=param, linewidth=2)
            
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
