import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import hashlib
import json
import os
from datetime import datetime

# ------------------------------
# 1. КОНФИГУРАЦИЯ
# ------------------------------
st.set_page_config(page_title="Оқушылар рейтингі", layout="wide")

# GitHub RAW файлдардың сілтемелері (ДҰРЫС ЖОЛДАРМЕН АУЫСТЫРЫҢЫЗ)
REITING_URL = "https://raw.githubusercontent.com/aidarpavl/reitingUcheniki/refs/heads/main/reiting%20ucheniki.xlsx"
BALDAR_URL = "https://raw.githubusercontent.com/aidarpavl/reitingUcheniki/refs/heads/main/%D0%91%D0%B0%D0%BB%D0%B4%D0%B0%D1%80.xlsx"
PAROL_URL = "https://raw.githubusercontent.com/aidarpavl/reitingUcheniki/refs/heads/main/Parol%20reiting%20ucheniki.xlsx"

# Google Drive сақтау орыны (сервис аккаунт керек)
# Бұл мысалда жергілікті файлға сақтаймыз
DATA_FILE = "reiting_updated.xlsx"
UPLOAD_DIR = "diploms"

if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# ------------------------------
# 2. ФУНКЦИЯЛАР
# ------------------------------
def load_data_from_github(url):
    """GitHub RAW сілтемесінен Excel жүктеу"""
    try:
        response = requests.get(url)
        response.raise_for_status()
        return pd.read_excel(BytesIO(response.content))
    except Exception as e:
        st.error(f"Файл жүктеу қатесі {url}: {e}")
        return None

def save_data_to_github(df, filename="reiting_updated.xlsx"):
    """Жаңартылған деректерді GitHub-қа push жасау (токен керек)"""
    # GitHub API арқылы жазу үшін personal access token қажет
    # Бұл мысалда жергілікті сақтаймыз
    df.to_excel(DATA_FILE, index=False)
    st.success("Деректер жергілікті сақталды (GitHub-қа жазу үшін токен керек)")
    return True

def check_password(student_name, input_password):
    """Парольді тексеру"""
    df_parol = load_data_from_github(PAROL_URL)
    if df_parol is None:
        return False
    # Баған атаулары: Оқушы_аты, Пароль
    row = df_parol[df_parol.iloc[:, 0].astype(str).str.strip() == student_name.strip()]
    if len(row) == 0:
        return False
    saved_password = str(row.iloc[0, 1]).strip()
    return input_password == saved_password

def calculate_rating(df_reiting, df_baldar):
    """Балдарды есептеу"""
    # Бірінші баған - оқушы аты, қалғандары - ұпайлар
    result_df = df_reiting.copy()
    for col in df_reiting.columns[1:]:  # Оқушы аты бағанын өткізіп
        if col in df_baldar.columns:
            result_df[col] = df_reiting[col] * df_baldar[col].iloc[0]
    return result_df

def get_recommendations(scores_row, threshold=5):
    """Ұсыныстар генерациялау"""
    recommendations = []
    for subject, score in scores_row.items():
        if score < threshold:
            recommendations.append(f"📖 {subject} пәні бойынша біліміңізді көтеріңіз (ұпай: {score})")
    if not recommendations:
        recommendations.append("✅ Барлық пәндер бойынша жақсы нәтиже!")
    return recommendations

def save_diploma(uploaded_file, student_name, diploma_name):
    """Дипломды сақтау"""
    if uploaded_file is not None:
        # Файл аты: Оқушы_аты_Диплом_атауы.кеңейту
        file_ext = uploaded_file.name.split('.')[-1]
        safe_student = student_name.replace(" ", "_")
        safe_diploma = diploma_name.replace(" ", "_")
        filename = f"{safe_student}_{safe_diploma}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{file_ext}"
        filepath = os.path.join(UPLOAD_DIR, filename)
        with open(filepath, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return filepath
    return None

# ------------------------------
# 3. АУТЕНТИФИКАЦИЯ
# ------------------------------
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.current_user = None
    st.session_state.is_admin = False

if not st.session_state.authenticated:
    st.title("🔐 Оқушылар рейтингі жүйесіне кіру")
    student_name = st.text_input("Оқушы аты")
    password = st.text_input("Пароль", type="password")
    if st.button("Кіру"):
        if check_password(student_name, password):
            st.session_state.authenticated = True
            st.session_state.current_user = student_name
            st.session_state.is_admin = (student_name.lower() == "admin")  # Әкімші арнайы
            st.success("Сәтті кірдіңіз!")
            st.rerun()
        else:
            st.error("Қате пароль немесе оқушы аты")
    st.stop()

# ------------------------------
# 4. НЕГІЗГІ ҚОСЫМША
# ------------------------------
st.sidebar.title(f"Қош келдіңіз, {st.session_state.current_user}!")
st.sidebar.button("Шығу", on_click=lambda: st.session_state.update(authenticated=False, current_user=None))

# Деректерді жүктеу
df_reiting = load_data_from_github(REITING_URL)
df_baldar = load_data_from_github(BALDAR_URL)

if df_reiting is None or df_baldar is None:
    st.error("Деректерді жүктеу мүмкін емес. GitHub сілтемелерін тексеріңіз.")
    st.stop()

tab1, tab2, tab3, tab4 = st.tabs(["📊 Рейтинг кестесі", "🎓 Диплом салу", "⭐ Рейтинг шығару", "📋 Ұсыныстар"])

# ------------------------------
# TAB 1: Кесте және өшіру
# ------------------------------
with tab1:
    st.subheader("Оқушылар рейтингі")
    
    # Кестені көрсету
    edited_df = st.data_editor(df_reiting, use_container_width=True, key="reiting_table")
    
    # Әр жолдың соңында өшіру батырмасы (Streamlit-те қолдан жасау керек)
    if st.session_state.is_admin:
        st.write("---")
        st.write("🗑️ **Әкімшіге арналған өшіру**")
        row_to_delete = st.number_input("Жол нөмірін енгізіңіз (0-ден бастап)", min_value=0, max_value=len(edited_df)-1, step=1)
        if st.button("Таңдалған жолды өшіру"):
            edited_df = edited_df.drop(index=row_to_delete).reset_index(drop=True)
            save_data_to_github(edited_df)
            st.rerun()
    else:
        st.info("Тек әкімші ғана жолдарды өшіре алады.")

# ------------------------------
# TAB 2: Диплом салу
# ------------------------------
with tab2:
    st.subheader("🏆 Грамота/Диплом салу")
    uploaded_file = st.file_uploader("Дипломды таңдаңыз (JPG, PNG, PDF)", type=["jpg", "png", "pdf"])
    diploma_name = st.text_input("Диплом атауы")
    if st.button("Дипломды сақтау"):
        if uploaded_file and diploma_name:
            filepath = save_diploma(uploaded_file, st.session_state.current_user, diploma_name)
            if filepath:
                st.success(f"Диплом сақталды: {filepath}")
        else:
            st.warning("Файл және диплом атауы қажет")

# ------------------------------
# TAB 3: Рейтинг шығару
# ------------------------------
with tab3:
    st.subheader("⭐ Балдарды есептеу")
    if st.button("Рейтинг шығару"):
        rated_df = calculate_rating(df_reiting, df_baldar)
        save_data_to_github(rated_df)
        st.dataframe(rated_df)
        st.success("Рейтинг есептелді және сақталды!")

# ------------------------------
# TAB 4: Ұсыныстар
# ------------------------------
with tab4:
    st.subheader("💡 Оқушыға арналған ұсыныстар")
    # Ағымдағы оқушының жолын табу
    user_row = df_reiting[df_reiting.iloc[:, 0].astype(str).str.strip() == st.session_state.current_user.strip()]
    if len(user_row) > 0:
        scores = user_row.iloc[0, 1:]  # Оқушы атынсыз ұпайлар
        recommendations = get_recommendations(scores)
        for rec in recommendations:
            st.write(rec)
    else:
        st.warning("Сіздің деректеріңіз табылмады")