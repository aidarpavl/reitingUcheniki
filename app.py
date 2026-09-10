import streamlit as st
import pandas as pd
import requests
from io import BytesIO
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

# GitHub RAW сілтемелер
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
    """GitHub RAW сілтемесінен Excel жүктеу — БЕЗ заголовков"""
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        df = pd.read_excel(BytesIO(response.content), header=None)
        return df
    except Exception as e:
        st.error(f"❌ Файл жүктеу қатесі: {e}")
        return None


def compute_student_score(row):
    """
    Оқушының жалпы балын есептеу.
    4 блок × 4 колонки: Қалалық, Облыстық, Республикалық, Халықаралық.
    Данные начинаются с колонки 3 (после Нөмір, Сынып, Оқушы).
    Формула: 1 орын = 3, 2 орын = 2, 3 орын = 1, номинация = 1
    """
    score = 0
    for block_start in [3, 7, 11, 15]:
        try:
            vals = []
            for offset in range(4):
                idx = block_start + offset
                v = row.iloc[idx] if idx < len(row) else 0
                if pd.isna(v):
                    v = 0
                try:
                    v = int(float(v))
                except (ValueError, TypeError):
                    v = 0
                vals.append(v)
            score += vals[0] * 3 + vals[1] * 2 + vals[2] * 1 + vals[3] * 1
        except Exception:
            continue
    return score


def prepare_dataframe(df_raw):
    """
    Ручная обработка двухуровневой шапки.
    Ищем строку с "номер", затем пропускаем ещё одну строку ("1 орын").
    """
    header_row_idx = None
    for i in range(min(5, len(df_raw))):
        row_vals = df_raw.iloc[i].astype(str).str.lower().tolist()
        if any('номер' in v for v in row_vals):
            header_row_idx = i
            break

    if header_row_idx is None:
        header_row_idx = 0
        data_start = 3
    else:
        data_start = header_row_idx + 2

    df_data = df_raw.iloc[data_start:].copy().reset_index(drop=True)

    df_data = df_data.rename(columns={
        0: 'Нөмір',
        1: 'Сынып',
        2: 'Оқушы'
    })

    df_data = df_data[df_data['Оқушы'].notna()].copy()
    df_data['Оқушы'] = df_data['Оқушы'].astype(str).str.strip()
    df_data = df_data[df_data['Оқушы'] != ''].copy()
    df_data = df_data[~df_data['Оқушы'].str.lower().isin(['nan', 'none', 'nat'])]
    df_data = df_data[~df_data['Оқушы'].str.lower().str.contains('итог|жалпы|барлығы|total', na=False)]

    df_data['Жалпы_балл'] = df_data.apply(compute_student_score, axis=1)

    df_data['Нөмір'] = pd.to_numeric(df_data['Нөмір'], errors='coerce').fillna(0).astype(int)
    df_data['Сынып'] = df_data['Сынып'].astype(str).str.strip()

    return df_data.reset_index(drop=True)


def get_top_students(df, n=3):
    """Үздік оқушылар"""
    df_with_score = df[df['Жалпы_балл'] > 0]
    if len(df_with_score) == 0:
        return df.head(n)
    return df_with_score.sort_values('Жалпы_балл', ascending=False).head(n)


def get_struggling_students(df, n=3):
    """Көмек қажет оқушылар (соңғы орындар, но с баллом > 0)"""
    df_with_score = df[df['Жалпы_балл'] > 0]
    if len(df_with_score) == 0:
        return df.head(n)
    return df_with_score.sort_values('Жалпы_балл', ascending=True).head(n)


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
# 3. ДЕРЕКТЕРДІ ЖҮКТЕУ
# ------------------------------
df_raw = load_excel_from_github(REITING_URL)

if df_raw is None:
    st.error("❌ GitHub-тан файлды жүктеу мүмкін болмады. Сілтемені тексеріңіз.")
    st.stop()

df = prepare_dataframe(df_raw)


# ------------------------------
# 4. БАСТЫ БЕТ
# ------------------------------
st.title("🏆 Оқушылар рейтингі")
st.caption(f"📅 Соңғы жаңарту: {datetime.now().strftime('%d.%m.%Y %H:%M')} | Барлығы: {len(df)} оқушы")


# ========== ОТЛАДКА (можно убрать после проверки) ==========
with st.expander("🐛 Деректерді тексеру (отладка)", expanded=False):
    st.write(f"**Сырые данные (df_raw):** {df_raw.shape[0]} строк × {df_raw.shape[1]} колонок")
    st.write(f"**Обработанные данные (df):** {len(df)} учеников")
    st.write("**Первые 10 учеников:**")
    st.dataframe(df[['Нөмір', 'Сынып', 'Оқушы', 'Жалпы_балл']].head(10), use_container_width=True)
    st.write("**Распределение баллов:**")
    st.write(df['Жалпы_балл'].describe())
# ==========================================================


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
        st.markdown(f"*Сынып: {student['Сынып']}*")
        st.markdown("---")

with col2:
    st.markdown("## ⚠️ Көмек қажет оқушылар")
    struggling = get_struggling_students(df, 3)
    labels = ["1-ең төмен", "2-ең төмен", "3-ең төмен"]
    for i, (_, student) in enumerate(struggling.iterrows()):
        st.markdown(f"**{labels[i]}**")
        st.markdown(f"### {student['Оқушы']}")
        st.markdown(f"`↑ {student['Жалпы_балл']} балл`")
        st.markdown(f"*Сынып: {student['Сынып']}*")
        st.markdown("---")


# ========== ЖАЛПЫ ИТОГТАР ==========
st.markdown("## 📊 Оқушылардың жалпы итогтары")

df_with_score = df[df['Жалпы_балл'] > 0].sort_values('Жалпы_балл', ascending=False)

if len(df_with_score) > 0:
    top30 = df_with_score.head(30).set_index('Оқушы')['Жалпы_балл']
    st.bar_chart(top30, height=500, use_container_width=True)
else:
    st.warning("⚠️ Ешбір оқушы әлі балл жинаған жоқ.")


# ========== ТОЛЫҚ КЕСТЕ ==========
st.markdown("## 📋 Толық кесте")

search = st.text_input("🔍 Оқушыны іздеу (аты-жөні немесе сынып):", "")

if search:
    filtered = df[
        df['Оқушы'].astype(str).str.contains(search, case=False, na=False) |
        df['Сынып'].astype(str).str.contains(search, case=False, na=False)
    ]
else:
    filtered = df

# Переименуем колонки дипломов для удобства чтения
df_display = filtered.copy()
column_names = {
    3:  'Қал. 1ор', 4:  'Қал. 2ор', 5:  'Қал. 3ор', 6:  'Қал. ном',
    7:  'Обл. 1ор', 8:  'Обл. 2ор', 9:  'Обл. 3ор', 10: 'Обл. ном',
    11: 'Респ. 1ор', 12: 'Респ. 2ор', 13: 'Респ. 3ор', 14: 'Респ. ном',
    15: 'Хал. 1ор', 16: 'Хал. 2ор', 17: 'Хал. 3ор', 18: 'Хал. ном',
}
for idx, name in column_names.items():
    if idx in df_display.columns:
        df_display = df_display.rename(columns={idx: name})

# Оставляем все нужные колонки
all_cols = ['Нөмір', 'Сынып', 'Оқушы']
all_cols += [v for k, v in column_names.items() if k in df_display.columns]
all_cols += ['Жалпы_балл']

df_display = df_display[all_cols].sort_values('Жалпы_балл', ascending=False)

# Числа → целые для красивого отображения
for col in all_cols:
    if col not in ['Сынып', 'Оқушы']:
        df_display[col] = pd.to_numeric(df_display[col], errors='coerce').fillna(0).astype(int)

st.dataframe(
    df_display,
    use_container_width=True,
    height=600
)

st.caption(f"📊 Көрсетілген жолдар: {len(df_display)} | Барлық бағандар: {len(df_display.columns)}")

# ========== ҰСЫНЫСТАР ==========
st.markdown("## 💡 Жеке ұсыныстар")

student_names = df['Оқушы'].tolist()
selected_student = st.selectbox("Оқушыны таңдаңыз:", student_names)

if selected_student:
    student_row = df[df['Оқушы'] == selected_student].iloc[0]
    score = student_row['Жалпы_балл']
    rec = get_recommendations(score, selected_student)
    st.info(rec)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**Сыныбы:** {student_row['Сынып']}")
        st.markdown(f"**Жалпы балл:** {score}")

    st.markdown("### 📊 Блоктар бойынша талдау")

    blocks = [
        ("Қалалық", 3),
        ("Облыстық", 7),
        ("Республикалық", 11),
        ("Халықаралық", 15)
    ]

    for block_name, start_col in blocks:
        try:
            d1 = int(float(student_row.iloc[start_col])) if pd.notna(student_row.iloc[start_col]) else 0
            d2 = int(float(student_row.iloc[start_col + 1])) if pd.notna(student_row.iloc[start_col + 1]) else 0
            d3 = int(float(student_row.iloc[start_col + 2])) if pd.notna(student_row.iloc[start_col + 2]) else 0
            dn = int(float(student_row.iloc[start_col + 3])) if pd.notna(student_row.iloc[start_col + 3]) else 0
            total = d1 * 3 + d2 * 2 + d3 * 1 + dn * 1
            if total > 0:
                st.markdown(
                    f"- **{block_name}:** 1ор={d1}, 2ор={d2}, 3ор={d3}, ном={dn} → **{total} балл**"
                )
        except (ValueError, IndexError, TypeError):
            continue


# ========== FOOTER ==========
st.markdown("---")
st.caption("✅ Рейтинг формуласы: 1 орын = 3 балл, 2 орын = 2 балл, 3 орын = 1 балл, номинация = 1 балл")
st.caption("📌 Деректер көзі: GitHub репозиторийі")