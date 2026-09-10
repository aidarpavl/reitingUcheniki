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
    """GitHub RAW сілтемесінен Excel жүктеу — без заголовков"""
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        return pd.read_excel(BytesIO(response.content), header=None)
    except Exception as e:
        st.error(f"❌ Файл жүктеу қатесі: {e}")
        return None


def compute_student_score(row):
    """4 блок × 4 колонки: 1ор=3, 2ор=2, 3ор=1, ном=1"""
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
    Обработка двухуровневой шапки.
    Возвращает df со всеми 19 столбцами:
      Нөмір, Сынып, Оқушы,
      Қал_1, Қал_2, Қал_3, Қал_ном,
      Обл_1, Обл_2, Обл_3, Обл_ном,
      Респ_1, Респ_2, Респ_3, Респ_ном,
      Хал_1, Хал_2, Хал_3, Хал_ном,
      Жалпы_балл
    """
    # Ищем строку с "номер"
    header_row_idx = None
    for i in range(min(5, len(df_raw))):
        row_vals = df_raw.iloc[i].astype(str).str.lower().tolist()
        if any('номер' in v for v in row_vals):
            header_row_idx = i
            break

    if header_row_idx is None:
        data_start = 3
    else:
        data_start = header_row_idx + 2

    df_data = df_raw.iloc[data_start:].copy().reset_index(drop=True)

    # Переименовываем ВСЕ 19 колонок явно
    rename_map = {
        0: 'Нөмір',
        1: 'Сынып',
        2: 'Оқушы',
        3: 'Қал_1',   4: 'Қал_2',   5: 'Қал_3',   6: 'Қал_ном',
        7: 'Обл_1',   8: 'Обл_2',   9: 'Обл_3',  10: 'Обл_ном',
        11: 'Респ_1', 12: 'Респ_2', 13: 'Респ_3', 14: 'Респ_ном',
        15: 'Хал_1',  16: 'Хал_2',  17: 'Хал_3',  18: 'Хал_ном',
    }
    df_data = df_data.rename(columns=rename_map)

    # Оставляем только эти 19 колонок (если есть больше — обрезаем)
    keep_cols = list(rename_map.values())
    df_data = df_data[[c for c in keep_cols if c in df_data.columns]].copy()

    # Чистим строки
    df_data = df_data[df_data['Оқушы'].notna()].copy()
    df_data['Оқушы'] = df_data['Оқушы'].astype(str).str.strip()
    df_data = df_data[df_data['Оқушы'] != ''].copy()
    df_data = df_data[~df_data['Оқушы'].str.lower().isin(['nan', 'none', 'nat'])]
    df_data = df_data[~df_data['Оқушы'].str.lower().str.contains('итог|жалпы|барлығы|total', na=False)]

    # Все числовые колонки → int (NaN → 0)
    numeric_cols = [c for c in keep_cols if c not in ['Сынып', 'Оқушы']]
    for col in numeric_cols:
        df_data[col] = pd.to_numeric(df_data[col], errors='coerce').fillna(0).astype(int)

    # Оставляем только 19 колонок (без Жалпы_балл — его добавим последним)
    df_data = df_data[keep_cols].copy()

    # Считаем Жалпы_балл
    df_data['Жалпы_балл'] = df_data.apply(compute_student_score, axis=1)

    return df_data.reset_index(drop=True)


def get_top_students(df, n=3):
    df_with_score = df[df['Жалпы_балл'] > 0]
    if len(df_with_score) == 0:
        return df.head(n)
    return df_with_score.sort_values('Жалпы_балл', ascending=False).head(n)


def get_struggling_students(df, n=3):
    df_with_score = df[df['Жалпы_балл'] > 0]
    if len(df_with_score) == 0:
        return df.head(n)
    return df_with_score.sort_values('Жалпы_балл', ascending=True).head(n)


def get_recommendations(score, name):
    if score >= 20:
        return f"🏆 **{name}**, сіз өте үздік нәтиже көрсеттіңіз! Осы деңгейді сақтаңыз."
    elif score >= 12:
        return f"✨ **{name}**, жақсы нәтиже! Қосымша дайындық курстарына қатысыңыз."
    elif score >= 6:
        return f"📈 **{name}**, қанағаттанарлық нәтиже. Әлсіз тақырыптарды пысықтаңыз."
    elif score >= 1:
        return f"⚠️ **{name}**, нәтиже орташадан төмен. Дайындықты күшейтіңіз."
    else:
        return f"🌱 **{name}**, әзірге диплом жоқ. Белсенділікті арттырыңыз."


# ------------------------------
# 3. ДЕРЕКТЕРДІ ЖҮКТЕУ
# ------------------------------
df_raw = load_excel_from_github(REITING_URL)

if df_raw is None:
    st.error("❌ GitHub-тан файлды жүктеу мүмкін болмады.")
    st.stop()

df = prepare_dataframe(df_raw)


# ------------------------------
# 4. БАСТЫ БЕТ
# ------------------------------
st.title("🏆 Оқушылар рейтингі")
st.caption(f"📅 {datetime.now().strftime('%d.%m.%Y %H:%M')} | Барлығы: {len(df)} оқушы | Бағандар: {len(df.columns)}")


# ========== ОТЛАДКА ==========
with st.expander("🐛 Деректерді тексеру", expanded=False):
    st.write(f"**df_raw:** {df_raw.shape[0]} × {df_raw.shape[1]}")
    st.write(f"**df:** {len(df)} строк × {len(df.columns)} колонок")
    st.write(f"**Столбцы df:** {list(df.columns)}")
    st.dataframe(df.head(10), use_container_width=True)


# ========== ТОП-3 және КӨМЕК ҚАЖЕТ ==========
col1, col2 = st.columns(2)

with col1:
    st.markdown("## 🏆 Үздік оқушылар")
    top3 = get_top_students(df, 3)
    medals = ["🥇 1-орын", "🥈 2-орын", "🥉 3-орын"]
    for i, (_, student) in enumerate(top3.iterrows()):
        st.markdown(f"**{medals[i]}**")
        st.markdown(f"### {student['Оқушы']}")
        st.markdown(f"`↑ {student['Жалпы_балл']} балл`  |  *{student['Сынып']}*")
        st.markdown("---")

with col2:
    st.markdown("## ⚠️ Көмек қажет оқушылар")
    struggling = get_struggling_students(df, 3)
    labels = ["1-ең төмен", "2-ең төмен", "3-ең төмен"]
    for i, (_, student) in enumerate(struggling.iterrows()):
        st.markdown(f"**{labels[i]}**")
        st.markdown(f"### {student['Оқушы']}")
        st.markdown(f"`↑ {student['Жалпы_балл']} балл`  |  *{student['Сынып']}*")
        st.markdown("---")


# ========== ДИАГРАММА ==========
st.markdown("## 📊 Оқушылардың жалпы итогтары")

df_with_score = df[df['Жалпы_балл'] > 0].sort_values('Жалпы_балл', ascending=False)

if len(df_with_score) > 0:
    top30 = df_with_score.head(30).set_index('Оқушы')['Жалпы_балл']
    st.bar_chart(top30, height=500, use_container_width=True)
else:
    st.warning("⚠️ Ешбір оқушы әлі балл жинаған жоқ.")


# ========== ТОЛЫҚ КЕСТЕ — ВСЕ 19 СТОЛБЦОВ ==========
st.markdown("## 📋 Толық кесте (барлық бағандар)")

search = st.text_input("🔍 Оқушыны іздеу (аты-жөні немесе сынып):", "")

if search:
    filtered = df[
        df['Оқушы'].astype(str).str.contains(search, case=False, na=False) |
        df['Сынып'].astype(str).str.contains(search, case=False, na=False)
    ]
else:
    filtered = df

# Сортируем по баллу, показываем ВСЕ колонки (19 + Жалпы_балл = 20)
filtered_sorted = filtered.sort_values('Жалпы_балл', ascending=False)

st.dataframe(
    filtered_sorted,
    use_container_width=True,
    height=600,
    column_config={
        'Нөмір': st.column_config.NumberColumn('№', width='small'),
        'Сынып': st.column_config.TextColumn('Кл', width='small'),
        'Оқушы': st.column_config.TextColumn('ФИО', width='medium'),
        'Жалпы_балл': st.column_config.NumberColumn('Жалпы балл', width='small'),
    }
)

st.caption(
    f"📊 Көрсетілген жолдар: **{len(filtered_sorted)}** | "
    f"Барлық бағандар: **{len(filtered_sorted.columns)}** "
    f"(Нөмір, Сынып, Оқушы + 16 диплом + Жалпы балл)"
)


# ========== ҰСЫНЫСТАР ==========
st.markdown("## 💡 Жеке ұсыныстар")

student_names = df['Оқушы'].tolist()
selected_student = st.selectbox("Оқушыны таңдаңыз:", student_names)

if selected_student:
    student_row = df[df['Оқушы'] == selected_student].iloc[0]
    score = student_row['Жалпы_балл']
    st.info(get_recommendations(score, selected_student))

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**Сыныбы:** {student_row['Сынып']}")
    with c2:
        st.markdown(f"**Жалпы балл:** {score}")

    st.markdown("### 📊 Блоктар бойынша талдау")

    blocks = [
        ("Қалалық",      ['Қал_1', 'Қал_2', 'Қал_3', 'Қал_ном']),
        ("Облыстық",     ['Обл_1', 'Обл_2', 'Обл_3', 'Обл_ном']),
        ("Республикалық",['Респ_1','Респ_2','Респ_3','Респ_ном']),
        ("Халықаралық",  ['Хал_1', 'Хал_2', 'Хал_3', 'Хал_ном']),
    ]

    for block_name, cols in blocks:
        try:
            d1 = int(student_row[cols[0]])
            d2 = int(student_row[cols[1]])
            d3 = int(student_row[cols[2]])
            dn = int(student_row[cols[3]])
            total = d1 * 3 + d2 * 2 + d3 * 1 + dn * 1
            if total > 0 or d1+d2+d3+dn > 0:
                st.markdown(
                    f"- **{block_name}:** 1ор={d1}, 2ор={d2}, 3ор={d3}, ном={dn} → **{total} балл**"
                )
        except (KeyError, ValueError, TypeError):
            continue


# ========== FOOTER ==========
st.markdown("---")
st.caption("✅ 1 орын = 3 балл, 2 орын = 2 балл, 3 орын = 1 балл, номинация = 1 балл")
st.caption("📌 Деректер көзі: GitHub репозиторийі")