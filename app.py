import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os

# ------------------------------
# 1. КОНФИГУРАЦИЯ
# ------------------------------
st.set_page_config(
    page_title="Оқушылар рейтингі",
    page_icon="🏆",
    layout="wide"
)

# GitHub RAW сілтемелер (ДҰРЫС ЖОЛДАР)
REITING_URL = "https://raw.githubusercontent.com/aidarpavl/reitingUcheniki/main/reiting%20ucheniki.xlsx"
BALDAR_URL = "https://raw.githubusercontent.com/aidarpavl/reitingUcheniki/main/%D0%91%D0%B0%D0%BB%D0%B4%D0%B0%D1%80.xlsx"
PAROL_URL = "https://raw.githubusercontent.com/aidarpavl/reitingUcheniki/main/Parol%20reiting%20ucheniki.xlsx"

UPLOAD_DIR = "diploms"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# ------------------------------
# 2. ФУНКЦИЯЛАР
# ------------------------------
@st.cache_data(ttl=300)
def load_excel_from_github(url):
    """GitHub RAW сілтемесінен Excel жүктеу"""
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        return pd.read_excel(BytesIO(response.content))
    except Exception as e:
        st.error(f"❌ Файл жүктеу қатесі: {e}")
        return None

def compute_student_score(row):
    """
    Оқушының жалпы балын есептеу.
    Формула: 1 орын = 3 балл, 2 орын = 2 балл, 3 орын = 1 балл, номинация = 1 балл
    """
    score = 0
    # Қалалық, Облыстық, Республикалық, Халықаралық блоктар
    # Әр блок 4 бағаннан тұрады: 1ор, 2ор, 3ор, ном
    for start_col in range(3, len(row), 4):
        try:
            score += (int(row.iloc[start_col] or 0) * 3)      # 1 орын
            score += (int(row.iloc[start_col + 1] or 0) * 2)  # 2 орын
            score += (int(row.iloc[start_col + 2] or 0) * 1)  # 3 орын
            score += (int(row.iloc[start_col + 3] or 0) * 1)  # Номинация
        except (ValueError, IndexError):
            continue
    return score

def prepare_dataframe(df):
    """Деректерді өңдеу: балл есептеу және бағандарды тазалау"""
    # Екі жолдан кейін деректер басталады (бірінші екі жол - тақырыптар)
    # Бірақ нақты құрылымды тексерейік
    df_clean = df.copy()
    
    # Бірінші бағандар: номер, кл, ФИО
    df_clean = df_clean.rename(columns={
        df_clean.columns[0]: 'Нөмір',
        df_clean.columns[1]: 'Сынып',
        df_clean.columns[2]: 'Оқушы'
    })
    
    # Балдарды есептеу
    df_clean['Жалпы_балл'] = df_clean.apply(compute_student_score, axis=1)
    
    # Пустые значения → 0
    df_clean['Жалпы_балл'] = df_clean['Жалпы_балл'].fillna(0).astype(int)
    
    return df_clean

def get_top_students(df, n=3):
    """Үздік оқушыларды алу"""
    df_sorted = df.sort_values('Жалпы_балл', ascending=False)
    return df_sorted.head(n)

def get_struggling_students(df, n=3):
    """Көмек қажет оқушыларды алу (соңғы орындар)"""
    df_with_score = df[df['Жалпы_балл'] > 0]  # Тек балл жинағандар
    if len(df_with_score) == 0:
        return df.head(n)
    df_sorted = df_with_score.sort_values('Жалпы_балл', ascending=True)
    return df_sorted.head(n)

def get_recommendations(score, name):
    """Оқушыға арналған ұсыныстар"""
    if score >= 20:
        return f"🏆 **{name}**, сіз өте үздік нәтиже көрсеттіңіз! Осы деңгейді сақтаңыз. Келесі олимпиадаларға қатысуды жалғастырыңыз!"
    elif score >= 12:
        return f"✨ **{name}**, жақсы нәтиже! Бірақ әлі де өсуге орын бар. Қосымша дайындық курстарына қатысыңыз."
    elif score >= 6:
        return f"📈 **{name}**, қанағаттанарлық нәтиже. Пән мұғалімдерімен кеңесіп, әлсіз тақырыптарды пысықтаңыз."
    elif score >= 1:
        return f"⚠️ **{name}**, нәтиже орташадан төмен. Олимпиадаларға дайындықты күшейтіп, топтық жұмыстарға көбірек қатысыңыз."
    else:
        return f"🌱 **{name}**, әзірге диплом жоқ. Белсенділікті арттырып, мұғалімдермен бірлесіп даму жоспарын құрыңыз."

def save_diploma(uploaded_file, student_name, diploma_name):
    """Дипломды сақтау"""
    if uploaded_file is not None:
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
# 3. БАСТАПҚЫ ЖҮКТЕУ
# ------------------------------
# Деректерді жүктеу
df_raw = load_excel_from_github(REITING_URL)

if df_raw is None:
    st.error("❌ GitHub-тан файлды жүктеу мүмкін болмады. Сілтемені тексеріңіз.")
    st.stop()

# Деректерді өңдеу
df = prepare_dataframe(df_raw)

# ------------------------------
# 4. СТРАНИЦА
# ------------------------------
st.title("🏆 Оқушылар рейтингі")
st.caption(f"📅 Соңғы жаңарту: {datetime.now().strftime('%d.%m.%Y %H:%M')}")

# ========== ЕКІ БАҒАН: ҮЗДІК және КӨМЕК ҚАЖЕТ ==========
col1, col2 = st.columns(2)

with col1:
    st.markdown("## 🏆 Үздік оқушылар")
    top3 = get_top_students(df, 3)
    medals = ["🥇 1-орын", "🥈 2-орын", "🥉 3-орын"]
    for i, (_, student) in enumerate(top3.iterrows()):
        st.markdown(f"**{medals[i]}**")
        st.markdown(f"### {student['Оқушы']}")
        st.markdown(f"`↑ {student['Жалпы_балл']} балл`")
        st.markdown("---")

with col2:
    st.markdown("## ⚠️ Көмек қажет оқушылар")
    struggling = get_struggling_students(df, 3)
    labels = ["1-ең төмен", "2-ең төмен", "3-ең төмен"]
    for i, (_, student) in enumerate(struggling.iterrows()):
        st.markdown(f"**{labels[i]}**")
        st.markdown(f"### {student['Оқушы']}")
        st.markdown(f"`↑ {student['Жалпы_балл']} балл`")
        st.markdown("---")

# ========== ЖАЛПЫ ИТОГТАР ==========
st.markdown("## 📊 Оқушылардың жалпы итогтары")

# Тек балл жинаған оқушыларды көрсету
df_with_score = df[df['Жалпы_балл'] > 0].sort_values('Жалпы_балл', ascending=False)

if len(df_with_score) > 0:
    # Диаграмма (бағандық)
    fig = px.bar(
        df_with_score.head(30),  # Топ 30
        x='Оқушы',
        y='Жалпы_балл',
        color='Жалпы_балл',
        color_continuous_scale=['#d73027', '#fee08b', '#1a9850'],
        labels={'Жалпы_балл': 'Жалпы балл', 'Оқушы': ''},
        title='Оқушылар рейтингі (Топ 30)'
    )
    fig.update_layout(
        height=500,
        xaxis_tickangle=-45,
        showlegend=False,
        coloraxis_showscale=False
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.warning("⚠️ Ешбір оқушы әлі балл жинаған жоқ.")

# ========== ТОЛЫҚ КЕСТЕ ==========
st.markdown("## 📋 Толық кесте")

# Іздеу
search = st.text_input("🔍 Оқушыны іздеу (аты-жөні немесе сынып):", "")

if search:
    filtered = df[
        df['Оқушы'].astype(str).str.contains(search, case=False, na=False) |
        df['Сынып'].astype(str).str.contains(search, case=False, na=False)
    ]
else:
    filtered = df

# Көрсетілетін бағандар
display_cols = ['Нөмір', 'Сынып', 'Оқушы', 'Жалпы_балл']
st.dataframe(
    filtered[display_cols],
    use_container_width=True,
    height=500
)

# ========== ҰСЫНЫСТАР ==========
st.markdown("## 💡 Жеке ұсыныстар")

# Оқушыны таңдау
student_names = df['Оқушы'].tolist()
selected_student = st.selectbox("Оқушыны таңдаңыз:", student_names)

if selected_student:
    student_row = df[df['Оқушы'] == selected_student].iloc[0]
    score = student_row['Жалпы_балл']
    rec = get_recommendations(score, selected_student)
    st.info(rec)
    
    # Толық статистика
    st.markdown(f"**Сыныбы:** {student_row['Сынып']}")
    st.markdown(f"**Жалпы балл:** {score}")
    
    # Блоктар бойынша талдау
    st.markdown("### 📊 Блоктар бойынша талдау")
    
    blocks = [
        ("Қалалық", 3),
        ("Облыстық", 7),
        ("Республикалық", 11),
        ("Халықаралық", 15)
    ]
    
    for block_name, start_col in blocks:
        try:
            d1 = int(student_row.iloc[start_col] or 0)
            d2 = int(student_row.iloc[start_col + 1] or 0)
            d3 = int(student_row.iloc[start_col + 2] or 0)
            dn = int(student_row.iloc[start_col + 3] or 0)
            total = d1 * 3 + d2 * 2 + d3 * 1 + dn * 1
            if total > 0:
                st.markdown(
                    f"- **{block_name}:** 1ор={d1}, 2ор={d2}, 3ор={d3}, ном={dn} → **{total} балл**"
                )
        except (ValueError, IndexError):
            continue

# ========== FOOTER ==========
st.markdown("---")
st.caption("✅ Рейтинг формуласы: 1 орын = 3 балл, 2 орын = 2 балл, 3 орын = 1 балл, номинация = 1 балл")
st.caption("📌 Деректер көзі: GitHub репозиторийі")