import io
import re
import unicodedata
from dataclasses import dataclass
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st


CATEGORIES: Dict[str, List[str]] = {
    "Income": ["Wages", "Gifts"],
    "Food": ["Groceries"],
    "House": ["Rent"],
    "Vehicle": ["Fuel", "Car Insurance"],
    "Life": ["Pet", "Vacations", "Education", "Sport/Fitness", "Health Care"],
    "Internet": [],
    "Taxes": [],
    "Savings": [],
    "Extras": ["Entertainment", "Gifts", "Electronics", "Home", "Clothes"],
}

# One rule per line: Expense Type = keyword1, keyword2, keyword phrase
# The left side must be one of the expense types listed above.
DEFAULT_RULES = """# Income
Wages = salary, wage, payroll, vencimento, ordenado
Gifts = gift, present, oferta

# Food
Groceries = grocery, groceries, supermarket, continente, pingo doce, lidl, mercadona, auchan, mini preco, minipreco

# House
Rent = rent, renda, landlord

# Vehicle
Fuel = fuel, gas, gasoline, diesel, petrol, galp, repsol, bp, prio
Car Insurance = car insurance, auto insurance, seguro auto, seguro carro

# Life
Pet = pet, vet, veterinary, animal
Vacations = hotel, booking, airbnb, flight, airline, vacation, travel
Education = school, university, course, udemy, coursera, livro, books
Sport/Fitness = gym, fitness, sport, sports, ginásio, ginasio
Health Care = pharmacy, farmacia, farmácia, doctor, hospital, clinic, health

# Other main categories
Internet = internet, telecom, vodafone, meo, nos
Taxes = tax, taxes, imposto, finanças, financas
Savings = savings, saving, poupanca, poupança, investment

# Extras
Entertainment = netflix, spotify, cinema, movie, concert, games
Gifts = gift, present, oferta
Electronics = electronics, apple, microsoft, amazon, tech
Home = ikea, furniture, home, household
Clothes = clothes, clothing, zara, hm, h&m, primark
"""


@dataclass
class Rule:
    expense_type: str
    keywords: List[str]


def normalize_text(value: object) -> str:
    """Normalize transaction descriptions for matching."""
    if pd.isna(value):
        return ""
    text = str(value).lower().strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text)
    return text


def valid_expense_types() -> List[str]:
    values: List[str] = []
    for category, subcategories in CATEGORIES.items():
        if subcategories:
            values.extend(subcategories)
        else:
            values.append(category)
    return values


def expense_type_to_category() -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    for category, subcategories in CATEGORIES.items():
        if subcategories:
            for expense_type in subcategories:
                mapping[expense_type] = category
        else:
            mapping[category] = category
    return mapping


def parse_rules(rules_text: str) -> Tuple[List[Rule], List[str]]:
    """Parse rules like: Expense Type = keyword1, keyword2."""
    valid_types = set(valid_expense_types())
    rules: List[Rule] = []
    warnings: List[str] = []

    for line_number, raw_line in enumerate(rules_text.splitlines(), start=1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            warnings.append(f"Line {line_number} ignored: missing '='.")
            continue

        expense_type, keywords_part = line.split("=", 1)
        expense_type = expense_type.strip()
        keywords = [normalize_text(k) for k in keywords_part.split(",") if k.strip()]

        if expense_type not in valid_types:
            warnings.append(f"Line {line_number} ignored: '{expense_type}' is not in the category list.")
            continue
        if not keywords:
            warnings.append(f"Line {line_number} ignored: no keywords provided.")
            continue

        rules.append(Rule(expense_type=expense_type, keywords=keywords))

    return rules, warnings


def find_header_row(raw_df: pd.DataFrame) -> int:
    """Find the row containing the transaction headers."""
    expected_headers = {"data lanc.", "data valor", "descricao", "valor", "saldo"}
    for idx, row in raw_df.iterrows():
        row_values = {normalize_text(v) for v in row.tolist()}
        if len(expected_headers.intersection(row_values)) >= 3:
            return idx
    raise ValueError("Could not find the transaction header row.")


def load_transactions(uploaded_file: io.BytesIO) -> pd.DataFrame:
    """Load Excel and return the clean transaction table."""
    raw_df = pd.read_excel(uploaded_file, header=None)
    header_row = find_header_row(raw_df)

    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, header=header_row)
    df = df.dropna(how="all")

    expected = ["Data Lanc.", "Data Valor", "Descrição", "Valor", "Saldo"]
    available = [col for col in expected if col in df.columns]
    if available:
        df = df[available]

    for col in ["Data Lanc.", "Data Valor"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    for col in ["Valor", "Saldo"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def classify_description(description: object, rules: List[Rule]) -> str:
    text = normalize_text(description)
    for rule in rules:
        if any(keyword in text for keyword in rule.keywords):
            return rule.expense_type
    return "Unclassified"


def clean_transactions(df: pd.DataFrame, rules: List[Rule]) -> pd.DataFrame:
    cleaned = df.copy()
    if "Descrição" not in cleaned.columns:
        raise ValueError("Column 'Descrição' was not found in the uploaded file.")

    type_to_category = expense_type_to_category()
    cleaned["Expense Type"] = cleaned["Descrição"].apply(lambda x: classify_description(x, rules))
    cleaned["Expense Category"] = cleaned["Expense Type"].map(type_to_category).fillna("Unclassified")

    # Income should always be positive. Expenses can remain positive or negative.
    if "Valor" in cleaned.columns:
        cleaned.loc[cleaned["Expense Category"] == "Income", "Valor"] = cleaned.loc[
            cleaned["Expense Category"] == "Income", "Valor"
        ].abs()

    return cleaned


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl", datetime_format="yyyy-mm-dd") as writer:
        df.to_excel(writer, index=False, sheet_name="Cleaned Transactions")
    return output.getvalue()


def main() -> None:
    st.set_page_config(page_title="Expense Cleaner", layout="wide")
    st.title("Expense Cleaner")
    st.write("Upload a bank Excel file, classify each row, and download the cleaned file.")

    with st.expander("Allowed categories and expense types", expanded=True):
        for category, expense_types in CATEGORIES.items():
            if expense_types:
                st.write(f"**{category}:** {', '.join(expense_types)}")
            else:
                st.write(f"**{category}**")

    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

    rules_text = st.text_area(
        "Classification rules",
        value=DEFAULT_RULES,
        height=420,
        help="Format: Expense Type = keyword1, keyword2, keyword phrase",
    )

    if uploaded_file is None:
        st.info("Upload an Excel file to begin.")
        return

    try:
        transactions = load_transactions(uploaded_file)
        rules, warnings = parse_rules(rules_text)
        for warning in warnings:
            st.warning(warning)

        cleaned = clean_transactions(transactions, rules)

        st.subheader("Preview")
        st.dataframe(cleaned, use_container_width=True)

        if "Valor" in cleaned.columns:
            st.subheader("Summary by Expense Category and Type")
            summary = (
                cleaned.groupby(["Expense Category", "Expense Type"], dropna=False)["Valor"]
                .sum()
                .reset_index()
            )
            st.dataframe(summary, use_container_width=True)

        excel_bytes = to_excel_bytes(cleaned)
        st.download_button(
            label="Download cleaned Excel",
            data=excel_bytes,
            file_name="cleaned_expenses.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.error(f"Could not process the file: {exc}")


if __name__ == "__main__":
    main()
