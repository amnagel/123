import streamlit as st, pandas as pd, requests as r, json, io

# Заголовок приложения на странице
st.title("Практическая №10 (Colab)")

# Адрес  бэкенда на Render
u = "https://one23-4bfu.onrender.com/p"

# Поля для ввода параметров в боковой панели
dx = st.sidebar.number_input("dx", 0.0)
dy = st.sidebar.number_input("dy", 0.0)
dz = st.sidebar.number_input("dz", 0.0)
m = st.sidebar.number_input("m", 0.0)

# Окно для загрузки файла
f = st.file_uploader("Excel", type=['xlsx'])

if f:
    # Читаем файл и показываем первые несколько строк для проверки
    df = pd.read_excel(f)
    st.write(df.head())
    
    # Кнопка запуска расчетов
    if st.button("Расчет"):
        f.seek(0)  # Сбрасываем указатель файла в начало перед отправкой
        
        # Собираем параметры в словарь
        ps = {"dx": dx, "dy": dy, "dz": dz, "m": m}
        
        # Отправляем файл и параметры на сервер Render
        res = r.post(u, files={"file": f}, data={"params": json.dumps(ps)})
        
        # Если сервер ответил успешно (код 200), даем скачать готовый Word-файл
        if res.status_code == 200:
            st.download_button("Скачать DOCX", res.content, "report.docx")
        else:
            st.error("Ошибка при расчете на сервере")
