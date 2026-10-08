import streamlit as st
import pandas as pd

st.set_page_config(page_title="GMAO Biomédicale", layout="wide")
st.title("🏥 GMAO Biomédicale")

st.write("Bienvenue sur la GMAO du service biomédical.")

uploaded_file = st.file_uploader("Déposer le fichier Excel des équipements", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("Données chargées avec succès !")
    st.dataframe(df)
