import math
import re

import pandas as pd
import plotly.express as px
import streamlit as st
from groq import Groq


st.set_page_config(
    page_title="FINPULSE",
    page_icon="💗",
    layout="wide",
)

PINK = "#EC4899"
DARK_PINK = "#BE185D"
PALE_PINK = "#F9A8D4"
LIGHT_PINK = "#FCE7F3"
BACKGROUND = "#FFF7FA"
DARK = "#27272A"

st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {BACKGROUND}; color: {DARK}; }}
    [data-testid="stSidebar"] {{ background-color: {LIGHT_PINK}; }}
    h1, h2, h3 {{ color: {DARK_PINK}; }}
    div.stButton > button {{
        background-color: {PINK};
        color: white;
        border: 0;
        border-radius: 8px;
    }}
    div.stButton > button:hover {{
        background-color: {DARK_PINK};
        color: white;
        border: 0;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("💗 FINPULSE")
st.caption("Analyze Today. Predict Tomorrow. Decide Better.")
st.write(
    "Unggah laporan keuangan, pilih kolom akun yang sesuai, "
    "lalu tinjau rasio dan simulasi secara sederhana."
)


def parse_number(value):
    """Mengubah angka, termasuk angka berformat umum Indonesia, menjadi float."""
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        number = float(value)
        return number if math.isfinite(number) else None

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "-", "—"}:
        return None

    is_negative = text.startswith("(") and text.endswith(")")
    text = text.replace("(", "").replace(")", "")
    text = text.replace("Rp", "").replace("rp", "").replace(" ", "")
    text = re.sub(r"[^0-9,.\-]", "", text)

    if not text or text in {"-", ".", ","}:
        return None

    # Jika titik dan koma ada, pemisah yang paling akhir dianggap desimal.
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        # Contoh: 1.234,56 atau 123,45
        if len(text.split(",")[-1]) == 3 and len(text.split(",")) > 1:
            text = text.replace(",", "")
        else:
            text = text.replace(",", ".")
    elif "." in text:
        # Contoh: 1.234 dianggap pemisah ribuan.
        if len(text.split(".")[-1]) == 3 and len(text.split(".")) > 1:
            text = text.replace(".", "")

    try:
        number = float(text)
        if is_negative:
            number = -abs(number)
        return number if math.isfinite(number) else None
    except ValueError:
        return None


def safe_divide(numerator, denominator):
    """Mencegah pembagian dengan nol atau data yang kosong."""
    if numerator is None or denominator is None or denominator == 0:
        return None

    result = numerator / denominator
    return result if math.isfinite(result) else None


def display_ratio(value, percent=False):
    if value is None:
        return "Tidak tersedia"

    if percent:
        return f"{value * 100:,.2f}%"

    return f"{value:,.2f} kali"


def find_suggested_column(columns, aliases):
    """Mencari nama kolom yang mirip dengan nama akun yang umum."""
    normalized_columns = {
        str(column).strip().lower(): column for column in columns
    }

    for alias in aliases:
        if alias.lower() in normalized_columns:
            return normalized_columns[alias.lower()]

    return None


ALIASES = {
    "Periode": ["periode", "tahun", "bulan", "tanggal", "period", "year"],
    "Pendapatan": ["pendapatan", "revenue", "sales", "penjualan"],
    "Harga Pokok Penjualan": [
        "harga pokok penjualan",
        "hpp",
        "cogs",
        "cost of goods sold",
    ],
    "Laba Operasi": [
        "laba operasi",
        "laba usaha",
        "operating income",
        "operating profit",
    ],
    "Laba Bersih": ["laba bersih", "net income", "net profit"],
    "Aset Lancar": ["aset lancar", "aktiva lancar", "current assets"],
    "Liabilitas Lancar": [
        "liabilitas lancar",
        "kewajiban lancar",
        "utang lancar",
        "current liabilities",
    ],
    "Kas": ["kas", "cash", "cash and cash equivalents"],
    "Total Aset": ["total aset", "total aktiva", "total assets"],
    "Total Liabilitas": [
        "total liabilitas",
        "total kewajiban",
        "total liabilities",
    ],
    "Ekuitas": ["ekuitas", "modal", "equity", "total equity"],
}

FINANCIAL_FIELDS = [
    "Pendapatan",
    "Harga Pokok Penjualan",
    "Laba Operasi",
    "Laba Bersih",
    "Aset Lancar",
    "Liabilitas Lancar",
    "Kas",
    "Total Aset",
    "Total Liabilitas",
    "Ekuitas",
]


def read_uploaded_file(uploaded_file):
    """Membaca file CSV atau Excel."""
    file_name = uploaded_file.name.lower()

    if file_name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file)
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding="latin-1")

    if file_name.endswith(".xlsx"):
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("Pilih sheet Excel", excel_file.sheet_names)
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, sheet_name=sheet_name)

    raise ValueError("Format file harus CSV (.csv) atau Excel (.xlsx).")


uploaded_file = st.file_uploader(
    "Unggah laporan keuangan",
    type=["csv", "xlsx"],
    help="Gunakan satu baris untuk setiap periode laporan.",
)

if uploaded_file is None:
    st.info("Unggah file CSV atau Excel untuk mulai menganalisis.")
    st.stop()

try:
    df = read_uploaded_file(uploaded_file)
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
except Exception as error:
    st.error(f"File tidak dapat dibaca: {error}")
    st.stop()

if df.empty:
    st.error("File tidak berisi data.")
    st.stop()

df.columns = [str(column).strip() for column in df.columns]
columns = list(df.columns)

st.subheader("1. Periksa dan petakan kolom")
st.caption(
    "Pastikan satu baris adalah satu periode. "
    "Pilih kolom yang paling sesuai untuk tiap akun. "
    "Kolom yang tidak tersedia boleh dibiarkan tidak dipilih."
)

none_option = "— Tidak tersedia —"

period_default = find_suggested_column(columns, ALIASES["Periode"])
period_options = [none_option] + columns
period_index = (
    period_options.index(period_default)
    if period_default in period_options
    else 0
)

period_column = st.selectbox(
    "Kolom periode (opsional)",
    period_options,
    index=period_index,
)

mapping = {}
mapping_columns = st.columns(2)

for index, field in enumerate(FINANCIAL_FIELDS):
    suggested = find_suggested_column(columns, ALIASES[field])
    options = [none_option] + columns
    default_index = options.index(suggested) if suggested in options else 0

    with mapping_columns[index % 2]:
        selected = st.selectbox(
            f"Kolom untuk {field}",
            options,
            index=default_index,
            key=f"mapping_{field}",
        )

    mapping[field] = None if selected == none_option else selected

st.subheader("2. Pilih periode yang dianalisis")

if period_column != none_option:
    period_labels = [
        f"Baris {index + 1} — {df.iloc[index][period_column]}"
        for index in range(len(df))
    ]
else:
    period_labels = [f"Baris {index + 1}" for index in range(len(df))]

selected_row_index = st.selectbox(
    "Periode/baris laporan",
    options=list(range(len(df))),
    format_func=lambda index: period_labels[index],
)

selected_row = df.iloc[selected_row_index]


def value_for(field, row=None):
    if row is None:
        row = selected_row

    column = mapping.get(field)
    if column is None:
        return None

    return parse_number(row[column])


values = {
    field: value_for(field)
    for field in FINANCIAL_FIELDS
}

revenue = values["Pendapatan"]
cogs = values["Harga Pokok Penjualan"]
operating_income = values["Laba Operasi"]
net_income = values["Laba Bersih"]
current_assets = values["Aset Lancar"]
current_liabilities = values["Liabilitas Lancar"]
cash = values["Kas"]
total_assets = values["Total Aset"]
total_liabilities = values["Total Liabilitas"]
equity = values["Ekuitas"]

ratios = {
    "Rasio Lancar": safe_divide(current_assets, current_liabilities),
    "Rasio Kas": safe_divide(cash, current_liabilities),
    "Liabilitas terhadap Aset": safe_divide(
        total_liabilities,
        total_assets,
    ),
    "Liabilitas terhadap Ekuitas": safe_divide(
        total_liabilities,
        equity,
    ),
    "Margin Laba Kotor": safe_divide(
        None if revenue is None or cogs is None else revenue - cogs,
        revenue,
    ),
    "Margin Operasi": safe_divide(operating_income, revenue),
    "Margin Laba Bersih": safe_divide(net_income, revenue),
    "ROA": safe_divide(net_income, total_assets),
    "ROE": safe_divide(net_income, equity),
}

st.subheader("3. Ringkasan laporan")

summary_columns = st.columns(4)
summary_fields = [
    "Pendapatan",
    "Laba Bersih",
    "Aset Lancar",
    "Total Liabilitas",
]

for index, field in enumerate(summary_fields):
    value = values[field]
    shown_value = (
        "Tidak tersedia"
        if value is None
        else f"{value:,.2f}"
    )
    summary_columns[index].metric(field, shown_value)

st.subheader("4. Rasio keuangan")

ratio_descriptions = {
    "Rasio Lancar": "Aset lancar ÷ liabilitas lancar",
    "Rasio Kas": "Kas ÷ liabilitas lancar",
    "Liabilitas terhadap Aset": "Total liabilitas ÷ total aset",
    "Liabilitas terhadap Ekuitas": "Total liabilitas ÷ ekuitas",
    "Margin Laba Kotor": "(Pendapatan − HPP) ÷ pendapatan",
    "Margin Operasi": "Laba operasi ÷ pendapatan",
    "Margin Laba Bersih": "Laba bersih ÷ pendapatan",
    "ROA": "Laba bersih ÷ total aset",
    "ROE": "Laba bersih ÷ ekuitas",
}

percent_ratios = {
    "Liabilitas terhadap Aset",
    "Margin Laba Kotor",
    "Margin Operasi",
    "Margin Laba Bersih",
    "ROA",
    "ROE",
}

ratio_columns = st.columns(3)

for index, (name, value) in enumerate(ratios.items()):
    ratio_columns[index % 3].metric(
        name,
        display_ratio(
            value,
            percent=name in percent_ratios,
        ),
        help=ratio_descriptions[name],
    )

st.caption(
    "Rasio yang tidak dapat dihitung ditampilkan sebagai “Tidak tersedia”. "
    "Penyebab umumnya: kolom belum dipetakan, data kosong, atau penyebutnya nol."
)

st.subheader("5. Grafik")

chart_values = {
    "Pendapatan": revenue,
    "Harga Pokok Penjualan": cogs,
    "Laba Operasi": operating_income,
    "Laba Bersih": net_income,
}

chart_df = pd.DataFrame(
    [
        {"Pos keuangan": name, "Nilai": value}
        for name, value in chart_values.items()
        if value is not None
    ]
)

if not chart_df.empty:
    financial_figure = px.bar(
        chart_df,
        x="Pos keuangan",
        y="Nilai",
        color="Pos keuangan",
        title="Ringkasan pos keuangan",
        color_discrete_sequence=[
            PINK,
            PALE_PINK,
            DARK_PINK,
            "#F472B6",
        ],
    )
    financial_figure.update_layout(showlegend=False)
    st.plotly_chart(financial_figure, use_container_width=True)
else:
    st.info(
        "Grafik pos keuangan belum tersedia. "
        "Petakan kolom akun terlebih dahulu."
    )

ratio_chart_df = pd.DataFrame(
    [
        {"Rasio": name, "Nilai": value}
        for name, value in ratios.items()
        if value is not None
    ]
)

if not ratio_chart_df.empty:
    ratio_figure = px.bar(
        ratio_chart_df,
        x="Rasio",
        y="Nilai",
        title="Perbandingan rasio keuangan",
        color_discrete_sequence=[PINK],
    )
    ratio_figure.update_layout(xaxis_tickangle=-35)
    st.plotly_chart(ratio_figure, use_container_width=True)
else:
    st.info("Belum ada rasio yang dapat digambarkan.")

st.subheader("6. Simulasi skenario")

st.write(
    "Simulasi ini menggunakan angka periode yang dipilih sebagai dasar. "
    "Hasilnya adalah perkiraan sederhana, bukan proyeksi audit atau nasihat keuangan."
)

if revenue is None or cogs is None:
    st.warning(
        "Petakan kolom Pendapatan dan Harga Pokok Penjualan "
        "untuk menggunakan simulasi."
    )
else:
    scenario_columns = st.columns(3)

    with scenario_columns[0]:
        revenue_growth = st.number_input(
            "Perubahan pendapatan (%)",
            min_value=-100.0,
            max_value=500.0,
            value=10.0,
            step=1.0,
        )

    with scenario_columns[1]:
        cogs_change = st.number_input(
            "Perubahan rasio HPP (poin persentase)",
            min_value=-100.0,
            max_value=100.0,
            value=0.0,
            step=1.0,
            help="Contoh: 2 berarti rasio HPP naik 2 poin persentase.",
        )

    with scenario_columns[2]:
        operating_expense_growth = st.number_input(
            "Perubahan beban operasi (%)",
            min_value=-100.0,
            max_value=500.0,
            value=5.0,
            step=1.0,
        )

    projected_revenue = revenue * (1 + revenue_growth / 100)
    base_cogs_ratio = safe_divide(cogs, revenue)

    if base_cogs_ratio is None:
        st.warning(
            "Simulasi HPP tidak dapat dihitung karena pendapatan dasar nol."
        )
    else:
        projected_cogs_ratio = max(
            0.0,
            base_cogs_ratio + cogs_change / 100,
        )
        projected_cogs = projected_revenue * projected_cogs_ratio

        projected_operating_income = None
        projected_net_income = None

        if operating_income is not None:
            base_operating_expenses = revenue - cogs - operating_income
            projected_operating_expenses = max(
                0.0,
                base_operating_expenses
                * (1 + operating_expense_growth / 100),
            )
            projected_operating_income = (
                projected_revenue
                - projected_cogs
                - projected_operating_expenses
            )

            if net_income is not None and operating_income != 0:
                net_to_operating_ratio = net_income / operating_income
                projected_net_income = (
                    projected_operating_income
                    * net_to_operating_ratio
                )
            elif net_income is not None:
                base_net_margin = safe_divide(net_income, revenue)
                if base_net_margin is not None:
                    projected_net_income = (
                        projected_revenue * base_net_margin
                    )

        result_columns = st.columns(4)

        result_columns[0].metric(
            "Perkiraan pendapatan",
            f"{projected_revenue:,.2f}",
        )
        result_columns[1].metric(
            "Perkiraan HPP",
            f"{projected_cogs:,.2f}",
        )
        result_columns[2].metric(
            "Perkiraan laba operasi",
            "Tidak tersedia"
            if projected_operating_income is None
            else f"{projected_operating_income:,.2f}",
        )
        result_columns[3].metric(
            "Perkiraan laba bersih",
            "Tidak tersedia"
            if projected_net_income is None
            else f"{projected_net_income:,.2f}",
        )

        st.caption(
            "Asumsi: beban operasi diperkirakan dari pendapatan − HPP − "
            "laba operasi pada periode dasar. Perubahan pajak, bunga, aset, "
            "liabilitas, dan faktor lainnya tidak dimodelkan."
        )

st.subheader("7. Tanya-jawab dengan Groq")

st.write(
    "Ajukan pertanyaan tentang rasio dan angka ringkasan di atas. "
    "Aplikasi mengirimkan ringkasan angka dan pertanyaan Anda ke Groq, "
    "bukan mengirimkan seluruh isi file."
)

if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

question = st.text_input(
    "Pertanyaan",
    placeholder="Contoh: Apa arti rasio lancar saya?",
)

if st.button("Tanyakan ke FINPULSE"):
    if not question.strip():
        st.warning("Tuliskan pertanyaan terlebih dahulu.")
    else:
        try:
            api_key = st.secrets["GROQ_API_KEY"]
            client = Groq(api_key=api_key)

            ratio_summary = {
                name: display_ratio(
                    value,
                    percent=name in percent_ratios,
                )
                for name, value in ratios.items()
            }

            financial_summary = {
                name: value
                for name, value in values.items()
                if value is not None
            }

            messages = [
                {
                    "role": "system",
                    "content": (
                        "Anda adalah asisten edukasi akuntansi berbahasa "
                        "Indonesia untuk mahasiswa pemula. Jelaskan dengan "
                        "bahasa sederhana. Gunakan hanya ringkasan data yang "
                        "diberikan. Jangan mengarang data yang tidak tersedia. "
                        "Ingatkan pengguna bahwa analisis ini bukan nasihat "
                        "profesional."
                    ),
                },
                {
                    "role": "user",
                    "content": (
                        f"Ringkasan angka keuangan: {financial_summary}\n"
                        f"Rasio yang dihitung: {ratio_summary}\n"
                        f"Pertanyaan pengguna: {question}"
                    ),
                },
            ]

            response = client.chat.completions.create(
                model="openai/gpt-oss-120b",
                messages=messages,
                temperature=0.3,
                max_tokens=700,
            )

            answer = response.choices[0].message.content

            st.session_state.chat_history.append(
                {
                    "question": question,
                    "answer": answer,
                }
            )

        except KeyError:
            st.error(
                "Secret GROQ_API_KEY belum ditemukan. "
                "Tambahkan API key pada Secrets di pengaturan Streamlit Cloud."
            )
        except Exception as error:
            st.error(
                f"Permintaan ke Groq gagal: {error}. "
                "Periksa API key, koneksi, dan ketersediaan model Groq."
            )

for chat_item in reversed(st.session_state.chat_history):
    with st.chat_message("user"):
        st.write(chat_item["question"])

    with st.chat_message("assistant"):
        st.write(chat_item["answer"])

st.divider()
st.caption(
    "FINPULSE adalah alat bantu belajar. Periksa kembali sumber laporan dan "
    "rumus yang digunakan sebelum mengambil keputusan."
)
