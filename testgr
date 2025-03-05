import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.title("График из Excel")

# Загрузка файла
uploaded_file = st.file_uploader("Загрузите Excel файл", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Читаем файл
    df = pd.read_excel(uploaded_file)
    # Устанавливаем первый столбец как индекс
    df.set_index(df.columns[0], inplace=True)
    
    # Выбор характеристики
    param = st.selectbox("Выберите характеристику", df.index.tolist())
    
    # Построение графика
    fig, ax = plt.subplots()
    ax.plot(df.columns, df.loc[param], 'b-')
    ax.set_xlabel("Время")
    ax.set_ylabel("Значение")
    ax.set_title(param)
    plt.xticks(rotation=45)
    
    # Отображение графика
    st.pyplot(fig)
