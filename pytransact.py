import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import random
from io import BytesIO


def save_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    return processed_data


def generate_random_data():
    return {
        "ITEM ID": random.randint(1000, 9999),
        "ITEM NAME": f"Item_{random.randint(1, 100)}",
        "ITEM PRICE": round(random.uniform(10, 500), 2),
        "GROUP": f"Group_{random.randint(1, 100)}",
        "BRAND": f"Brand_{random.randint(1, 500)}",
        "MANUFACTUR": f"Manufact_{random.randint(1, 5)}",
        "EXPIRE": (pd.Timestamp.now() + pd.DateOffset(days=random.randint(30, 365))).strftime('%Y-%m-%d')
    }


st.title("Excel Data Input & Visualizer")

# Initialize session state for persistent data
if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame(
        columns=["ITEM ID", "ITEM NAME", "ITEM PRICE", "GROUP", "BRAND", "MANUFACTUR", "EXPIRE"])

# Input Form
with st.form("data_entry"):
    new_data = {}
    for col in st.session_state.df.columns:
        new_data[col] = st.text_input(f"Enter {col}")
    submitted = st.form_submit_button("Add Data")

    if submitted:
        new_row = pd.DataFrame([new_data])
        st.session_state.df = pd.concat([st.session_state.df, new_row], ignore_index=True)
        st.success("Data Added Successfully!")
        st.dataframe(st.session_state.df)

# Random Data Button
if st.button("Insert Random Data"):
    random_entry = generate_random_data()
    random_row = pd.DataFrame([random_entry])
    st.session_state.df = pd.concat([st.session_state.df, random_row], ignore_index=True)
    st.success("Random Data Inserted Successfully!")
    st.dataframe(st.session_state.df)

# Download Updated File
if not st.session_state.df.empty:
    st.download_button("Download Excel", data=save_excel(st.session_state.df), file_name="data.xlsx",
                       mime="application/vnd.ms-excel")

# Visualization
# Visualization
if not st.session_state.df.empty:
    st.write("### Data Visualization")
    st.session_state.df["EXPIRE"] = pd.to_datetime(st.session_state.df["EXPIRE"], errors='coerce')
    st.session_state.df['Year'] = st.session_state.df["EXPIRE"].dt.year
    st.session_state.df['Month'] = st.session_state.df["EXPIRE"].dt.month
    st.session_state.df['Day'] = st.session_state.df["EXPIRE"].dt.day

    chart_type = st.selectbox("Select Chart Type", ["Line", "Bar", "Pie"])

    if chart_type == "Line":
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(st.session_state.df["EXPIRE"].astype(str), st.session_state.df["ITEM PRICE"], marker='o')
        ax.set_xlabel("Expire Date")
        ax.set_ylabel("Item Price")
        ax.set_xticklabels(st.session_state.df["EXPIRE"].astype(str), rotation=45)
        st.pyplot(fig)
    elif chart_type == "Bar":
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.bar(st.session_state.df["EXPIRE"].astype(str), st.session_state.df["ITEM PRICE"], color='skyblue')
        ax.set_xlabel("Expire Date")
        ax.set_ylabel("Item Price")
        ax.set_xticklabels(st.session_state.df["EXPIRE"].astype(str), rotation=45)
        st.pyplot(fig)
    elif chart_type == "Pie":
        fig, ax = plt.subplots(figsize=(7, 7))
        ax.pie(st.session_state.df["ITEM PRICE"], labels=st.session_state.df["EXPIRE"].astype(str), autopct='%1.1f%%',
               startangle=140)
        st.pyplot(fig)
