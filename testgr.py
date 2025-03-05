import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import AutoMinorLocator

st.title("График из Excel")

# Загрузка файла
uploaded_file = st.file_uploader("Загрузите Excel файл", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Читаем файл
        df = pd.read_excel(uploaded_file)
        # Устанавливаем первый столбец как индекс
        df.set_index(df.columns[0], inplace=True)
        
        # Выбор характеристики
        param = st.selectbox("Выберите характеристику", df.index.tolist())
        
        # Построение графика
        fig, ax = plt.subplots(figsize=(10, 6))  # Увеличиваем размер фигуры для лучшей читаемости
        ax.plot(df.columns, df.loc[param], 'b-')
        
        # Настройка меток осей
        ax.set_xlabel("Время")
        ax.set_ylabel("Значение")
        ax.set_title(param)
        
        # Ротация меток на оси X и настройка интервала
        plt.xticks(rotation=45, ha="right")  # Поворачиваем метки на 45 градусов и выравниваем вправо
        
        # Автоматическая подгонка меток, чтобы избежать наложения
        fig.tight_layout()  # Убирает наложение меток и заголовков
        
        # Дополнительно: можно ограничить количество меток, если их слишком много
        if len(df.columns) > 10:  # Если меток больше 10, показываем только каждую пятую
            ax.xaxis.set_major_locator(plt.MaxNLocator(nbins=10))  # Ограничиваем количество меток
        
        # Добавляем сетку для лучшей читаемости
        ax.grid(True, linestyle='--', alpha=0.7)
        
        # Отображение графика
        st.pyplot(fig)
    except Exception as e:
        st.error(f"Ошибка: {str(e)}")
