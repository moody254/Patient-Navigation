from pathlib import Path
from datetime import datetime
import re
import shutil

import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURATION
# ============================================================

APP_DIR = Path(__file__).resolve().parent
EXCEL_FILE = APP_DIR / "patient_database.xlsx"
SHEET_NAME = "Dataset"

COLUMNS = [
    "Patient ID",
    "Patient Name",
    "Age",
    "Cancer Type",
    "Biomarkers",
    "Risk Factors",
    "Marital Status",
    "Gender",
    "Type of Support",
    "Amount in KSH",
    "Date",
    "Cancer Stage",
    "Year of Diagnosis",
    "Comorbidities",
    "Residence",
    "Type of Treatment",
    "Hospital of Treatment",
    "Patient Status",
    "Date of Passing",
]

NAVY = "#102A43"
DEEP_NAVY = "#071C2C"
MUSTARD = "#D4A017"
LIGHT_MUSTARD = "#F1C75B"


# ============================================================
# INITIAL DATA
# ============================================================

CLINICAL_SEED = [
    ["Juliet Mwango", 33, "Female", "Ca Oesophagus", "", "Used contraceptive for more Than 5yrs(femiplan)"],
    ["Gladys Gacheri", 40, "Female", "Ca Cx (Cervix)", "", "Uncertain,Lifestyle"],
    ["John Mtunga Musembi", 41, "Male", "Ca Mouth (Parotid Gland)", "", "Alcohol, Smoking, industrial exposure"],
    ["Christopher", 41, "Male", "Ca Colorectum", "", "Genetic,Lifestyle"],
    ["Lucy Kageni", 63, "Female", "Ca Cx (Cervix)", "", "Used contraceptive for more Than 5yrs(femiplan)"],
    ["Margaret Karimi", 48, "Female", "Ca Esophagus", "", "Long use of alcohol"],
    ["Caroline Atieno", 40, "Female", "Ca Breast", "", "HIV reactive, Lifestyle"],
    ["Simon Njoroge", 28, "Male", "Ca Penile", "", "Uncertain,Lifestyle"],
    ["Pheny Atieno", 37, "Female", "Ca Cx (Cervix)", "", "Used contraceptive for more Than 5yrs(femiplan)"],
    [
        "Hellen Murugi Macharia",
        46,
        "Female",
        "Invasive Ductal Carcinoma(NOS)",
        "PR(+Ve),ER(+Ve),HER2(+ve)2+,Ki67 30%",
        "Used contraceptive for more Than 5yrs(Femiplan)",
    ],
    ["kitheka Kwinga", 86, "Male", "Ca Esophagus", "", "Ciggaret smoking 4packets a day"],
    ["Josphat Mbaabu", 47, "Male", "Ca Tangue", "", "Long use of alcohol,Lifestyle"],
    [
        "Ruth Muthoni Kariuki",
        40,
        "Female",
        "Ca Breast",
        "ER(-Ve), PR(-Ve),HER2(+Ve)3+,Ki6735-40%",
        "Used contraceptive for more Than 7yrs",
    ],
    ["Ruth Wafula", None, "Female", "", "", ""],
    ["Justus Kawawa", 50, "Male", "Ca Lung", "", "Ciggarets,alcohol, lifestyle"],
]


SUPPORT_SEED = [
    ["Charles Gathiga Gathuku", 67, "Ca esophageal", "married", "Male", "PEG tube", 52377, "11-Mar-24", "Grade II", 2022, "Mulnutrinion", "Nyeri", "ChemoRad", "KNH"],
    ["Wilson Karani", 63, "Ca esophageal", "married", "Male", "Radio Planning", 10000, "18-Mar-24", "Stage III", 2022, "Diabetes", "Muranga", "Radiotherapy", "KNH"],
    ["Robai Nekesa", 41, "Ca Breast", "Seperated", "Female", "Medication", 3800, "22-Feb-24", "Grade III", 2022, "Osrioporosis", "Narok", "Chemo", "KNH"],
    ["Freshia Wairimu", 67, "Karposis Sarcoma", "Widowed", "Female", "Medication", 5500, "21-Feb-24", "Stage IV", 2022, "Paralised", "Kiambu", "ChemoRad", "KUTTRH"],
    ["Charles Gathiga Gathuku", 67, "Ca esophageal", "married", "Male", "NHIF", 6000, "13-Feb-24", "Grade II", 2022, "Mulnutrinion", "Nyeri", "ChemoRad", "KNH"],
    ["Heaven Wanyama", 4, "Hodkin'S lymphoma", "Null", "Female", "Treatment", 254104, "2-Feb-24", "Stage III", 2022, "None", "Bungoma", "Chemo", "KNH"],
    ["Francis Mwangi", 73, "Ca esophageal", "Seperated", "Male", "Labs", 2360, "31-Jan-24", "Stage IV", 2021, "Asthmatic", "Kiambu", "Chemo", "Coptic"],
    ["Esther Wandia", 44, "Ca Cx", "Widowed", "Female", "Labs", 2360, "31-Jan-24", "Stage III", 2022, "HIV Reactive", "Kiambu", "ChemoRad", "KNH"],
    ["Robai Nekesa", 41, "Ca Breast", "Seperated", "Female", "Chemotherapy", 78159, "2-Mar-24", "Grade III", 2022, "Osrioporosis", "Narok", "Chemo", "KNH"],
    ["Esther Kirema", 75, "Ca Cx", "Widowed", "Female", "Radiotherapy", 18000, "18-Feb-24", "Stage III", 2021, "Dementia", "Meru", "ChemoRad", "KNH"],
    ["Charity Wanjiku", 43, "Ca Breast", "Seperated", "Female", "Chemotherapy", 148000, "7-Feb-24", "Stage II", 2022, "Null", "Kiambu", "Chemotherapy", "KNH"],
    ["Esther Kirema", 75, "Ca Cx", "Widowed", "Female", "CT, chest, pelvic,abd24000", 24000, "31-Jan-24", "Stage III", 2021, "Dementia", "Meru", "ChemoRad", "Fortis Diagonostics"],
    ["Elizabeth wangui Kamau", 36, "Ca Vulva", "separated", "Female", "PET scan", 50000, "25-Mar-24", "Stage III", 2022, "None", "Kiambu", "ChemoRad", "KUTRH"],
    ["Hedgwig Musabi", 47, "Ca Breast", "Seperated", "Female", "U/S biopsy", 15000, "5/27/2024", "Stage I", 2024, "None", "Nairobi", "ChemoRad", "KNH"],
    ["Janet Awinja", 55, "Ca Cx", "Seperated", "Female", "Brachytherapy", 40000, "6-Mar-24", "Stage III", 2023, "HIV Reactive", "Nairobi", "Brachytherapy", "KNH"],
    ["Zacharia Nduati", 42, "Ca esophageal", "married", "Male", "MRI", 18000, "23-Apr-24", "Stage III", 2023, "None", "Kirinyaga", "ChemoRad", "Texas"],
    ["Pauline Wanjiru Mwangi", 66, "Plasmacytoma myeloma", "married", "Female", "Treatment", 171022, "4/12/2024", "Stage IV", 2022, "Odema", "Muranga", "ChemoRad", "KNH"],
    ["Pauline Wanjiru Mwangi", 66, "Plasmacytoma myeloma", "married", "Female", "Treatment", 62475, "4/12/2024", "Stage IV", 2022, "Odema", "Muranga", "Admision", "PCEA Kikuyu Hospital"],
    ["Agnes Wanjiru Muchiri", 52, "Ca Breast", "married", "Female", "Consultation", 3000, "2/14/2024", "Stage III", 2023, "Diabetes", "Kirinyaga", "ChemoRad", "Nairobi west Hsp"],
    ["Rael Kalekye Mutua", 59, "Neuroendocrine Carcinoma", "married", "Female", "Labs", 9910, "7/4/2024", "Stage III", 2023, "Hypertension", "Makueni", "ChemoRad", "KUTTRH"],
    ["Edward Salongo", 52, "Colorectal", "Seperated", "Male", "Radio Planning", 10000, "30-Sep-24", "Stage II", 2023, "colostomy Bag", "Kampla Uganda", "ChemoRad", "KNH"],
    ["Josephat Mwaniki", 78, "Ca esophageal", "Widowed", "Male", "Treatment", 8200, "9-Oct-24", "Stage II", 2023, "Hypertension", "Embu", "Radiotherapy", "KNH"],
    ["Thomas Kyalo Nzioka", 42, "Ca Spinal cord", "married", "Male", "Labs", 6500, "7-Oct-24", "Stage II", 2024, "HIV Reactive", "Machakos", "Radiotherapy", "KNH"],
    ["Martha Mwikali Mulae", 42, "Ca Vulva", "Seperated", "Female", "Treatment", 60966, "6-Sep-24", "Stage III", 2022, "HIV Reactive", "Kajiado", "Chemo", "Texas Cancer Center"],
]


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def norm(value):
    if pd.isna(value) or value is None:
        return ""

    return re.sub(r"\s+", " ", str(value).strip().lower())


def safe_date(value):
    if pd.isna(value) or value in ("", None):
        return pd.NaT

    return pd.to_datetime(value, errors="coerce", dayfirst=True)


def fingerprint(row):
    date = safe_date(row.get("Date"))

    if pd.isna(date):
        date_key = ""
    else:
        date_key = date.strftime("%Y-%m-%d")

    amount = pd.to_numeric(
        row.get("Amount in KSH"),
        errors="coerce"
    )

    if pd.isna(amount):
        amount_key = ""
    else:
        amount_key = f"{float(amount):.2f}"

    return "|".join([
        norm(row.get("Patient ID")),
        norm(row.get("Type of Support")),
        amount_key,
        date_key,
    ])


def clean_df(df):
    df = df.copy()

    for column in COLUMNS:
        if column not in df.columns:
            df[column] = ""

    df = df[COLUMNS]

    df["Patient Name"] = (
        df["Patient Name"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["Patient ID"] = (
        df["Patient ID"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["Amount in KSH"] = pd.to_numeric(
        df["Amount in KSH"],
        errors="coerce"
    ).fillna(0.0)

    df["Age"] = pd.to_numeric(
        df["Age"],
        errors="coerce"
    )

    df["Year of Diagnosis"] = pd.to_numeric(
        df["Year of Diagnosis"],
        errors="coerce"
    )

    df["Date"] = pd.to_datetime(
        df["Date"],
        errors="coerce",
        dayfirst=True
    )

    df["Date of Passing"] = pd.to_datetime(
        df["Date of Passing"],
        errors="coerce",
        dayfirst=True
    )

    text_columns = [
        column for column in COLUMNS
        if column not in [
            "Amount in KSH",
            "Age",
            "Year of Diagnosis",
            "Date",
            "Date of Passing",
        ]
    ]

    for column in text_columns:
        df[column] = (
            df[column]
            .fillna("")
            .astype(str)
        )

    df["Patient Status"] = (
        df["Patient Status"]
        .replace("", "Alive")
        .fillna("Alive")
    )

    return df


def next_patient_id(df):
    numbers = []

    for value in df["Patient ID"].dropna():
        match = re.search(
            r"NBO-(\d+)-26$",
            str(value)
        )

        if match:
            numbers.append(int(match.group(1)))

    next_number = max(numbers, default=0) + 1

    return f"TLMHS-JH-NBO-{next_number:03d}-26"


def assign_ids(df):
    df = clean_df(df)

    name_to_id = {}

    for _, row in df.iterrows():

        name = norm(row["Patient Name"])
        patient_id = str(row["Patient ID"]).strip()

        if (
            name
            and patient_id.startswith("TLMHS-JH-NBO-")
        ):
            name_to_id[name] = patient_id

    used_ids = set(name_to_id.values())

    highest_number = 0

    for patient_id in used_ids:

        match = re.search(
            r"NBO-(\d+)-26$",
            patient_id
        )

        if match:
            highest_number = max(
                highest_number,
                int(match.group(1))
            )

    changed = False

    for index in df.index:

        name = norm(
            df.at[index, "Patient Name"]
        )

        patient_id = str(
            df.at[index, "Patient ID"]
        ).strip()

        if not name:
            continue

        if name in name_to_id:

            correct_id = name_to_id[name]

            if patient_id != correct_id:

                df.at[index, "Patient ID"] = correct_id
                changed = True

        else:

            highest_number += 1

            new_id = (
                f"TLMHS-JH-NBO-"
                f"{highest_number:03d}-26"
            )

            name_to_id[name] = new_id
            df.at[index, "Patient ID"] = new_id

            changed = True

    return df, changed


# ============================================================
# CREATE INITIAL DATA
# ============================================================

def create_seed_dataframe():

    rows = []

    for (
        name,
        age,
        gender,
        cancer,
        biomarkers,
        risks,
    ) in CLINICAL_SEED:

        rows.append({
            "Patient ID": "",
            "Patient Name": name,
            "Age": age,
            "Cancer Type": cancer,
            "Biomarkers": biomarkers,
            "Risk Factors": risks,
            "Marital Status": "",
            "Gender": gender,
            "Type of Support": "",
            "Amount in KSH": 0,
            "Date": pd.NaT,
            "Cancer Stage": "",
            "Year of Diagnosis": None,
            "Comorbidities": "",
            "Residence": "",
            "Type of Treatment": "",
            "Hospital of Treatment": "",
            "Patient Status": "Alive",
            "Date of Passing": pd.NaT,
        })

    for (
        name,
        age,
        cancer,
        marital,
        gender,
        support,
        amount,
        date,
        stage,
        year,
        comorbidities,
        residence,
        treatment,
        hospital,
    ) in SUPPORT_SEED:

        rows.append({
            "Patient ID": "",
            "Patient Name": name,
            "Age": age,
            "Cancer Type": cancer,
            "Biomarkers": "",
            "Risk Factors": "",
            "Marital Status": marital,
            "Gender": gender,
            "Type of Support": support,
            "Amount in KSH": amount,
            "Date": safe_date(date),
            "Cancer Stage": stage,
            "Year of Diagnosis": year,
            "Comorbidities": comorbidities,
            "Residence": residence,
            "Type of Treatment": treatment,
            "Hospital of Treatment": hospital,
            "Patient Status": "Alive",
            "Date of Passing": pd.NaT,
        })

    df = pd.DataFrame(
        rows,
        columns=COLUMNS
    )

    df, _ = assign_ids(df)

    return df


# ============================================================
# MERGE SEED DATA WITHOUT DUPLICATING EXISTING DATA
# ============================================================

def merge_seed_data(df):

    df = clean_df(df)

    df, _ = assign_ids(df)

    existing_names = {
        norm(value)
        for value in df["Patient Name"]
        if norm(value)
    }

    # Add clinical profiles that do not already exist
    for (
        name,
        age,
        gender,
        cancer,
        biomarkers,
        risks,
    ) in CLINICAL_SEED:

        if norm(name) not in existing_names:

            new_id = next_patient_id(df)

            new_patient = {
                "Patient ID": new_id,
                "Patient Name": name,
                "Age": age,
                "Cancer Type": cancer,
                "Biomarkers": biomarkers,
                "Risk Factors": risks,
                "Marital Status": "",
                "Gender": gender,
                "Type of Support": "",
                "Amount in KSH": 0,
                "Date": pd.NaT,
                "Cancer Stage": "",
                "Year of Diagnosis": None,
                "Comorbidities": "",
                "Residence": "",
                "Type of Treatment": "",
                "Hospital of Treatment": "",
                "Patient Status": "Alive",
                "Date of Passing": pd.NaT,
            }

            df = pd.concat(
                [
                    df,
                    pd.DataFrame([new_patient]),
                ],
                ignore_index=True,
            )

            existing_names.add(norm(name))

    existing_fingerprints = {
        fingerprint(row)
        for _, row in df.iterrows()
        if norm(row.get("Type of Support"))
    }

    # Add support records
    for (
        name,
        age,
        cancer,
        marital,
        gender,
        support,
        amount,
        date,
        stage,
        year,
        comorbidities,
        residence,
        treatment,
        hospital,
    ) in SUPPORT_SEED:

        matches = df[
            df["Patient Name"].map(norm)
            == norm(name)
        ]

        if matches.empty:

            new_id = next_patient_id(df)

            patient_profile = {
                "Patient ID": new_id,
                "Patient Name": name,
                "Age": age,
                "Cancer Type": cancer,
                "Biomarkers": "",
                "Risk Factors": "",
                "Marital Status": marital,
                "Gender": gender,
                "Type of Support": "",
                "Amount in KSH": 0,
                "Date": pd.NaT,
                "Cancer Stage": stage,
                "Year of Diagnosis": year,
                "Comorbidities": comorbidities,
                "Residence": residence,
                "Type of Treatment": treatment,
                "Hospital of Treatment": hospital,
                "Patient Status": "Alive",
                "Date of Passing": pd.NaT,
            }

            df = pd.concat(
                [
                    df,
                    pd.DataFrame([patient_profile]),
                ],
                ignore_index=True,
            )

            matches = df[
                df["Patient Name"].map(norm)
                == norm(name)
            ]

        patient_id = matches.iloc[0]["Patient ID"]

        support_record = {
            "Patient ID": patient_id,
            "Patient Name": name,
            "Age": age,
            "Cancer Type": cancer,
            "Biomarkers": "",
            "Risk Factors": "",
            "Marital Status": marital,
            "Gender": gender,
            "Type of Support": support,
            "Amount in KSH": amount,
            "Date": safe_date(date),
            "Cancer Stage": stage,
            "Year of Diagnosis": year,
            "Comorbidities": comorbidities,
            "Residence": residence,
            "Type of Treatment": treatment,
            "Hospital of Treatment": hospital,
            "Patient Status": "Alive",
            "Date of Passing": pd.NaT,
        }

        record_fingerprint = fingerprint(
            support_record
        )

        if record_fingerprint not in existing_fingerprints:

            df = pd.concat(
                [
                    df,
                    pd.DataFrame([support_record]),
                ],
                ignore_index=True,
            )

            existing_fingerprints.add(
                record_fingerprint
            )

    # Fill missing profile information from other records
    profile_columns = [
        "Age",
        "Cancer Type",
        "Biomarkers",
        "Risk Factors",
        "Marital Status",
        "Gender",
        "Cancer Stage",
        "Year of Diagnosis",
        "Comorbidities",
        "Residence",
        "Type of Treatment",
        "Hospital of Treatment",
    ]

    for patient_id, group in df.groupby(
        "Patient ID",
        sort=False
    ):

        if not patient_id:
            continue

        for column in profile_columns:

            valid_values = [
                value
                for value in group[column]
                if str(value).strip()
                not in ("", "nan", "NaT")
            ]

            if valid_values:

                value = valid_values[0]

                df.loc[
                    df["Patient ID"] == patient_id,
                    column,
                ] = value

    return clean_df(df)


# ============================================================
# SAVE / LOAD DATABASE
# ============================================================

def save_database(df, create_backup=True):

    df = clean_df(df)

    df, _ = assign_ids(df)

    if (
        create_backup
        and EXCEL_FILE.exists()
    ):

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        backup_file = (
            APP_DIR
            / f"patient_database_backup_{timestamp}.xlsx"
        )

        shutil.copy2(
            EXCEL_FILE,
            backup_file
        )

    with pd.ExcelWriter(
        EXCEL_FILE,
        engine="openpyxl",
        mode="w",
    ) as writer:

        df.to_excel(
            writer,
            sheet_name=SHEET_NAME,
            index=False,
        )


def load_database():

    if not EXCEL_FILE.exists():

        df = create_seed_dataframe()

        save_database(
            df,
            create_backup=False
        )

        return df

    try:

        df = pd.read_excel(
            EXCEL_FILE,
            sheet_name=SHEET_NAME,
        )

    except Exception:

        df = create_seed_dataframe()

        save_database(
            df,
            create_backup=False
        )

        return df

    original_length = len(df)
    original_columns = list(df.columns)

    df = merge_seed_data(df)

    df, ids_changed = assign_ids(df)

    needs_save = (
        original_length != len(df)
        or original_columns != list(df.columns)
        or ids_changed
    )

    if needs_save:

        save_database(
            df,
            create_backup=False
        )

    return df


# ============================================================
# UI
# ============================================================

def apply_css():

    st.markdown(
        f"""
        <style>

        .stApp {{
            background: {DEEP_NAVY};
            color: #F5F0E6;
        }}

        [data-testid="stHeader"] {{
            background: {DEEP_NAVY};
        }}

        [data-testid="stSidebar"] {{
            background: {NAVY};
            border-right: 1px solid {MUSTARD};
        }}

        [data-testid="stSidebar"] * {{
            color: #F5F0E6 !important;
        }}

        h1, h2, h3, h4 {{
            color: {LIGHT_MUSTARD} !important;
        }}

        p, label {{
            color: #F5F0E6 !important;
        }}

        .metric-card {{
            background: {NAVY};
            border: 1px solid {MUSTARD};
            border-radius: 10px;
            padding: 18px;
            min-height: 120px;
            margin-bottom: 12px;
        }}

        .metric-label {{
            color: {LIGHT_MUSTARD};
            font-size: 0.80rem;
            font-weight: 700;
            text-transform: uppercase;
        }}

        .metric-value {{
            color: #F5F0E6;
            font-size: 1.9rem;
            font-weight: 800;
            margin-top: 8px;
        }}

        .metric-note {{
            color: #AFC0D2;
            font-size: 0.75rem;
            margin-top: 4px;
        }}

        div[data-baseweb="input"],
        div[data-baseweb="select"],
        textarea,
        input {{
            background: #0C2639 !important;
            color: #F5F0E6 !important;
            border-color: #49657A !important;
        }}

        div[data-baseweb="select"] * {{
            color: #F5F0E6 !important;
            background: #0C2639 !important;
        }}

        .stButton > button {{
            background: {MUSTARD};
            color: {DEEP_NAVY};
            border: 1px solid {MUSTARD};
            font-weight: 800;
            border-radius: 7px;
        }}

        .stButton > button:hover {{
            background: {LIGHT_MUSTARD};
            color: {DEEP_NAVY};
        }}

        [data-testid="stDataFrame"] {{
            border: 1px solid #49657A;
        }}

        .section {{
            border-left: 4px solid {MUSTARD};
            padding-left: 12px;
            margin: 20px 0 12px;
        }}

        </style>
        """,
        unsafe_allow_html=True)


def metric(label, value, note=""):

    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def patient_list(df):

    return (
        df[
            ["Patient ID", "Patient Name"]
        ]
        .drop_duplicates()
        .sort_values("Patient Name")
    )


# ============================================================
# DASHBOARD
# ============================================================

def dashboard(df):

    st.title(
        " TLMHS Patient Navigation Management System"
    )

    st.caption(
        "TLMHS patient navigation, accommodation and support database"
    )

    patient_count = (
        df["Patient ID"]
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )

    alive_count = (
        df.loc[
            df["Patient Status"] == "Alive",
            "Patient ID",
        ]
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )

    deceased_count = (
        df.loc[
            df["Patient Status"] == "Deceased",
            "Patient ID",
        ]
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )

    support_count = int(
        df["Type of Support"]
        .map(norm)
        .ne("")
        .sum()
    )

    total_support = float(
        df["Amount in KSH"].sum()
    )

    a, b, c, d, e = st.columns(5)

    with a:
        metric(
            "Registered Patients",
            patient_count,
            "Unique patient IDs",
        )

    with b:
        metric(
            "Alive Patients",
            alive_count,
            "Current recorded status",
        )

    with c:
        metric(
            "Deceased Patients",
            deceased_count,
            "Recorded outcomes",
        )

    with d:
        metric(
            "Support Records",
            support_count,
            "Support entries",
        )

    with e:
        metric(
            "Total Support",
            f"KES {total_support:,.0f}",
            "Recorded support value",
        )

    st.markdown(
        '<div class="section"><h3>Patient Overview</h3></div>',
        unsafe_allow_html=True,
    )

    left, right = st.columns(2)

    patient_profiles = (
        df.drop_duplicates(
            "Patient ID",
            keep="first",
        )
    )

    with left:

        cancer_counts = (
            patient_profiles["Cancer Type"]
            .replace("", "Not recorded")
            .value_counts()
            .head(10)
        )

        st.bar_chart(cancer_counts)

    with right:

        gender_counts = (
            patient_profiles["Gender"]
            .replace("", "Not recorded")
            .value_counts()
        )

        st.bar_chart(gender_counts)

    st.markdown(
        '<div class="section"><h3>Recent Support Activity</h3></div>',
        unsafe_allow_html=True,
    )

    recent = (
        df[df["Date"].notna()]
        .sort_values(
            "Date",
            ascending=False,
        )
        .head(10)
    )

    if recent.empty:

        st.info(
            "No dated support records found."
        )

    else:

        display = recent[
            [
                "Date",
                "Patient ID",
                "Patient Name",
                "Type of Support",
                "Amount in KSH",
                "Hospital of Treatment",
            ]
        ].copy()

        display["Date"] = (
            display["Date"]
            .dt.strftime("%d %b %Y")
        )

        st.dataframe(
            display,
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# PATIENT REGISTRY
# ============================================================

def patient_registry(df):

    st.title("Patient Registry")
    df = add_new_patient(df)
    search = st.text_input(
        "Search patient",
        placeholder="Name, ID, cancer type or residence",
    )

    work = df.copy()

    if search:

        query = norm(search)

        mask = (
            work["Patient Name"]
            .map(norm)
            .str.contains(
                query,
                na=False,
            )
            |
            work["Patient ID"]
            .map(norm)
            .str.contains(
                query,
                na=False,
            )
            |
            work["Cancer Type"]
            .map(norm)
            .str.contains(
                query,
                na=False,
            )
            |
            work["Residence"]
            .map(norm)
            .str.contains(
                query,
                na=False,
            )
        )

        work = work[mask]

    patient_display = work.drop_duplicates(
        "Patient ID",
        keep="first",
    )

    st.dataframe(
        patient_display[
            [
                "Patient ID",
                "Patient Name",
                "Age",
                "Gender",
                "Cancer Type",
                "Patient Status",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )

    st.markdown(
        '<div class="section"><h3>Patient Profile</h3></div>',
        unsafe_allow_html=True,
    )

    choices = patient_list(work)

    if choices.empty:

        st.info(
            "No patient matches the search."
        )

        return

    selected_id = st.selectbox(
        "Select patient",
        choices["Patient ID"].tolist(),
        format_func=lambda pid:
            f"{pid} — "
            f"{choices.loc[choices['Patient ID'] == pid, 'Patient Name'].iloc[0]}",
    )

    patient_records = df[
        df["Patient ID"] == selected_id
    ].copy()

    profile = patient_records.iloc[0]

    a, b, c, d = st.columns(4)

    with a:
        metric(
            "Status",
            profile["Patient Status"],
            "Recorded outcome",
        )

    with b:
        age_value = (
            "Not recorded"
            if pd.isna(profile["Age"])
            else int(profile["Age"])
        )

        metric(
            "Age",
            age_value,
            "Years",
        )

    with c:
        metric(
            "Total Support",
            f"KES {patient_records['Amount in KSH'].sum():,.0f}",
            "Cumulative support",
        )

    with d:
        metric(
            "Support Records",
            int(
                patient_records[
                    "Type of Support"
                ]
                .map(norm)
                .ne("")
                .sum()
            ),
            "Entries",
        )

    st.subheader(
        profile["Patient Name"]
    )

    profile_fields = [
        "Patient ID",
        "Cancer Type",
        "Gender",
        "Marital Status",
        "Cancer Stage",
        "Year of Diagnosis",
        "Biomarkers",
        "Risk Factors",
        "Comorbidities",
        "Residence",
        "Type of Treatment",
        "Hospital of Treatment",
        "Patient Status",
        "Date of Passing",
    ]

    profile_values = []

    for field in profile_fields:

        value = profile[field]

        if pd.isna(value):
            value = "Not recorded"

        elif field == "Date of Passing":

            if value:
                value = pd.to_datetime(
                    value,
                    errors="coerce",
                )

                if pd.notna(value):
                    value = value.strftime(
                        "%d %b %Y"
                    )
                else:
                    value = "Not recorded"

        profile_values.append(value)

    profile_table = pd.DataFrame(
        {
            "Field": profile_fields,
            "Value": profile_values,
        }
    )

    st.dataframe(
        profile_table,
        use_container_width=True,
        hide_index=True,
    )

    history = patient_records[
        patient_records[
            "Type of Support"
        ]
        .map(norm)
        .ne("")
    ][
        [
            "Date",
            "Type of Support",
            "Amount in KSH",
            "Cancer Stage",
            "Type of Treatment",
            "Hospital of Treatment",
        ]
    ].copy()

    if not history.empty:

        st.subheader(
            "Support History"
        )

        history["Date"] = (
            history["Date"]
            .dt.strftime("%d %b %Y")
        )

        st.dataframe(
            history,
            use_container_width=True,
            hide_index=True,
        )
        # ============================================================
# ADD NEW PATIENT
# ============================================================

def add_new_patient(df):

    with st.expander("➕ Add New Patient", expanded=False):

        with st.form("new_patient_form", clear_on_submit=True):

            left, right = st.columns(2)

            with left:
                name = st.text_input("Patient Name *")
                age = st.number_input(
                    "Age", min_value=0, max_value=120, step=1, value=0
                )
                gender = st.selectbox(
                    "Gender", ["", "Female", "Male", "Other"]
                )
                marital_status = st.text_input("Marital Status")
                cancer_type = st.text_input("Cancer Type")
                cancer_stage = st.text_input("Cancer Stage")

            with right:
                year_of_diagnosis = st.number_input(
                    "Year of Diagnosis",
                    min_value=1950,
                    max_value=datetime.now().year,
                    value=datetime.now().year,
                    step=1,
                )
                biomarkers = st.text_input("Biomarkers")
                risk_factors = st.text_input("Risk Factors")
                comorbidities = st.text_input("Comorbidities")
                residence = st.text_input("Residence")
                treatment = st.text_input("Type of Treatment")

            hospital = st.text_input("Hospital of Treatment")

            submitted = st.form_submit_button(
                "Save New Patient", icon=":material/person_add:"
            )

        if not submitted:
            return df

        name = name.strip()

        if not name:
            st.error("Patient Name is required.")
            return df

        existing_names = {
            norm(value) for value in df["Patient Name"]
        }

        if norm(name) in existing_names:
            st.error(
                "A patient with this name already exists. "
                "Use Add Support to add a record for them instead."
            )
            return df

        new_id = next_patient_id(df)

        new_patient = {
            "Patient ID": new_id,
            "Patient Name": name,
            "Age": age if age > 0 else None,
            "Cancer Type": cancer_type,
            "Biomarkers": biomarkers,
            "Risk Factors": risk_factors,
            "Marital Status": marital_status,
            "Gender": gender,
            "Type of Support": "",
            "Amount in KSH": 0,
            "Date": pd.NaT,
            "Cancer Stage": cancer_stage,
            "Year of Diagnosis": year_of_diagnosis,
            "Comorbidities": comorbidities,
            "Residence": residence,
            "Type of Treatment": treatment,
            "Hospital of Treatment": hospital,
            "Patient Status": "Alive",
            "Date of Passing": pd.NaT,
        }

        df = pd.concat(
            [df, pd.DataFrame([new_patient])],
            ignore_index=True,
        )

        save_database(df)

        st.session_state.data = df

        st.success(f"Patient '{name}' added with ID {new_id}.")

        st.rerun()

    return df
    # --------------------------------------------------------
    # DELETE PATIENT
    # --------------------------------------------------------

    st.markdown(
        '<div class="section"><h3>Delete Entire Patient</h3></div>',
        unsafe_allow_html=True,
    )

    st.warning(
        "This permanently removes the patient and ALL associated support records. "
        "A backup of the database will be created first."
    )

    confirmation = st.checkbox(
        "I understand that this deletes the entire patient record."
    )

    typed_id = st.text_input(
        "Type the Patient ID to confirm deletion"
    )

    delete_allowed = (
        confirmation
        and typed_id.strip() == selected_id
    )

    if st.button(
        "Delete Patient",
        icon=":material/delete:",
        disabled=not delete_allowed,
    ):

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        if EXCEL_FILE.exists():

            backup_file = (
                APP_DIR
                / f"patient_database_backup_{timestamp}.xlsx"
            )

            shutil.copy2(
                EXCEL_FILE,
                backup_file,
            )

        new_df = df[
            df["Patient ID"] != selected_id
        ].copy()

        save_database(
            new_df,
            create_backup=False,
        )

        st.session_state.data = new_df

        st.success(
            "Patient and all associated records were deleted."
        )

        st.rerun()


# ============================================================
# ADD SUPPORT
# ============================================================

def add_support(df):

    st.title("Add Support Record")

    choices = patient_list(df)

    if choices.empty:

        st.info(
            "No patients are currently available."
        )

        return

    with st.form(
        "support_form",
        clear_on_submit=False,
    ):

        selected_id = st.selectbox(
            "Patient",
            choices["Patient ID"].tolist(),
            format_func=lambda pid:
                f"{pid} — "
                f"{choices.loc[choices['Patient ID'] == pid, 'Patient Name'].iloc[0]}",
        )

        patient = df[
            df["Patient ID"] == selected_id
        ].iloc[0]

        st.caption(
            f"Patient: {patient['Patient Name']} | "
            f"Status: {patient['Patient Status']}"
        )

        left, right = st.columns(2)

        with left:

            support_type = st.text_input(
                "Type of Support",
                placeholder="Medication, Labs, Accommodation, etc.",
            )

            amount = st.number_input(
                "Amount in KSH",
                min_value=0.0,
                step=100.0,
            )

        with right:

            support_date = st.date_input(
                "Date",
                value=datetime.now().date(),
            )

            hospital = st.text_input(
                "Hospital of Treatment"
            )

        treatment = st.text_input(
            "Type of Treatment"
        )

        notes = st.text_area(
            "Support / Case Notes"
        )

        submitted = st.form_submit_button(
            "Save Support Record",
            icon=":material/save:",
        )

    if not submitted:
        return

    support_type = support_type.strip()

    if not support_type:

        st.error(
            "Please enter the Type of Support."
        )

        return

    new_record = {
        "Patient ID": selected_id,
        "Patient Name": patient["Patient Name"],
        "Age": patient["Age"],
        "Cancer Type": patient["Cancer Type"],
        "Biomarkers": patient["Biomarkers"],
        "Risk Factors": patient["Risk Factors"],
        "Marital Status": patient["Marital Status"],
        "Gender": patient["Gender"],
        "Type of Support": support_type,
        "Amount in KSH": amount,
        "Date": pd.Timestamp(support_date),
        "Cancer Stage": patient["Cancer Stage"],
        "Year of Diagnosis": patient["Year of Diagnosis"],
        "Comorbidities": (
            notes.strip()
            or patient["Comorbidities"]
        ),
        "Residence": patient["Residence"],
        "Type of Treatment": (
            treatment.strip()
            or patient["Type of Treatment"]
        ),
        "Hospital of Treatment": (
            hospital.strip()
            or patient["Hospital of Treatment"]
        ),
        "Patient Status": patient["Patient Status"],
        "Date of Passing": patient["Date of Passing"],
    }

    record_key = fingerprint(
        new_record
    )

    # Session-level duplicate protection
    if (
        record_key
        == st.session_state.get(
            "last_saved_key"
        )
    ):

        st.warning(
            "This exact record was already saved. "
            "No duplicate was created."
        )

        return

    # Database-level duplicate protection
    existing_keys = {
        fingerprint(row)
        for _, row in df.iterrows()
    }

    if record_key in existing_keys:

        st.error(
            "This exact support record already exists. "
            "No duplicate was created."
        )

        return

    df = pd.concat(
        [
            df,
            pd.DataFrame([new_record]),
        ],
        ignore_index=True,
    )

    save_database(df)

    st.session_state.data = df

    st.session_state.last_saved_key = (
        record_key
    )

    st.success(
        "Support record saved successfully."
    )

    st.rerun()


# ============================================================
# DATA QUALITY / DUPLICATES
# ============================================================

def data_quality(df):

    st.title(
        "Data Quality & Record Control"
    )

    working = df.copy()

    working["_Fingerprint"] = (
        working.apply(
            fingerprint,
            axis=1,
        )
    )

    working["Duplicate"] = (
        working["_Fingerprint"]
        .duplicated(keep=False)
        &
        working["Type of Support"]
        .map(norm)
        .ne("")
    )

    duplicates = working[
        working["Duplicate"]
    ].copy()

    a, b, c = st.columns(3)

    with a:
        metric(
            "Duplicate Records",
            len(duplicates),
            "Exact matching fingerprints",
        )

    with b:
        metric(
            "Blank Patient IDs",
            int(
                df["Patient ID"]
                .map(norm)
                .eq("")
                .sum()
            ),
            "Records requiring IDs",
        )

    with c:
        metric(
            "Profile-Only Rows",
            int(
                df["Type of Support"]
                .map(norm)
                .eq("")
                .sum()
            ),
            "No support entry",
        )

    if duplicates.empty:

        st.success(
            "No exact duplicate support records detected."
        )

        return

    st.subheader(
        "Duplicate Support Records"
    )

    display = duplicates[
        [
            "Patient ID",
            "Patient Name",
            "Type of Support",
            "Amount in KSH",
            "Date",
            "Hospital of Treatment",
        ]
    ].copy()

    display["Date"] = (
        display["Date"]
        .dt.strftime("%d %b %Y")
    )

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
    )

    selected_index = st.selectbox(
        "Select the exact record to delete",
        duplicates.index.tolist(),
        format_func=lambda index:
            f"Row {index + 1} — "
            f"{df.loc[index, 'Patient Name']} — "
            f"{df.loc[index, 'Type of Support']} — "
            f"KES {df.loc[index, 'Amount in KSH']:,.2f}",
    )

    confirmation = st.checkbox(
        "I understand this permanently deletes the selected record."
    )

    if st.button(
        "Delete Selected Record",
        icon=":material/delete:",
        disabled=not confirmation,
    ):

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        if EXCEL_FILE.exists():

            backup_file = (
                APP_DIR
                / f"patient_database_backup_{timestamp}.xlsx"
            )

            shutil.copy2(
                EXCEL_FILE,
                backup_file,
            )

        new_df = (
            df.drop(
                index=selected_index
            )
            .reset_index(drop=True)
        )

        save_database(
            new_df,
            create_backup=False,
        )

        st.session_state.data = new_df

        st.success(
            "Selected duplicate record was deleted."
        )

        st.rerun()


# ============================================================
# PATIENT OUTCOMES
# ============================================================

def outcomes(df):

    st.title(
        "Patient Outcomes"
    )

    choices = patient_list(df)

    alive_ids = (
        df.loc[
            df["Patient Status"] == "Alive",
            "Patient ID",
        ]
        .drop_duplicates()
        .tolist()
    )

    if not alive_ids:

        st.info(
            "No patients are currently recorded as alive."
        )

    else:

        selected_id = st.selectbox(
            "Patient to mark as deceased",
            alive_ids,
            format_func=lambda pid:
                f"{pid} — "
                f"{choices.loc[choices['Patient ID'] == pid, 'Patient Name'].iloc[0]}",
        )

        passing_date = st.date_input(
            "Date of Passing",
            value=datetime.now().date(),
        )

        confirmation = st.checkbox(
            "I confirm this patient outcome should be recorded as deceased."
        )

        if st.button(
            "Record Death",
            icon=":material/event_busy:",
            disabled=not confirmation,
        ):

            df.loc[
                df["Patient ID"] == selected_id,
                "Patient Status",
            ] = "Deceased"

            df.loc[
                df["Patient ID"] == selected_id,
                "Date of Passing",
            ] = pd.Timestamp(
                passing_date
            )

            save_database(df)

            st.session_state.data = df

            st.success(
                "Patient outcome recorded successfully."
            )

            st.rerun()

    deceased = (
        df[
            df["Patient Status"] == "Deceased"
        ][
            [
                "Patient ID",
                "Patient Name",
                "Date of Passing",
            ]
        ]
        .drop_duplicates("Patient ID")
    )

    if not deceased.empty:

        st.subheader(
            "Recorded Deceased Patients"
        )

        deceased = deceased.copy()

        deceased["Date of Passing"] = (
            deceased["Date of Passing"]
            .dt.strftime("%d %b %Y")
        )

        st.dataframe(
            deceased,
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# MAIN APPLICATION
# ============================================================

def main():

    st.set_page_config(
        page_title="TLMHS Joanna's House",
        page_icon=":material/health_and_safety:",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    apply_css()

    if "data" not in st.session_state:

        st.session_state.data = (
            load_database()
        )

    df = clean_df(
        st.session_state.data
    )

    with st.sidebar:

        st.markdown(
            "# TLMHS"
        )

        st.markdown(
            "### Joanna's House"
        )

        st.caption(
            "Patient Support Management"
        )

        st.divider()

        page = st.radio(
            "Navigation",
            [
                "Dashboard",
                "Patient Registry",
                "Add Support",
                "Data Quality",
                "Outcomes",
            ],
            label_visibility="collapsed",
        )

        st.divider()

        st.caption(
            f"Database: {EXCEL_FILE.name}"
        )

        st.caption(
            "Patients: "
            f"{df['Patient ID'].replace('', pd.NA).dropna().nunique()}"
        )

    if page == "Dashboard":

        dashboard(df)

    elif page == "Patient Registry":

        patient_registry(df)

    elif page == "Add Support":

        add_support(df)

    elif page == "Data Quality":

        data_quality(df)

    elif page == "Outcomes":

        outcomes(df)


if __name__ == "__main__":
    main()