# streamlit_app.py

import streamlit as st
from openpyxl import load_workbook
from openpyxl.writer.excel import save_virtual_workbook
from Einspringprogramm import berechne_einspringer_statistik

st.title("Einspringprogramm Web-App (openpyxl)")

uploaded_file = st.file_uploader("Excel-Datei (Dienstpläne2025.xlsx) hochladen", type=["xlsx"])

if uploaded_file is not None:
    wb = load_workbook(uploaded_file)
    wb = berechne_einspringer_statistik(wb)

    st.success("Einspringer-Statistik erfolgreich berechnet!")

    # Optional: Beispielwert anzeigen (z. B. erstes Feld der Statistik)
    ws_stat = wb["Einspringer-Statistik"]
    beispielwert = ws_stat["A2"].value if ws_stat.max_row >= 2 else "Keine Daten"
    st.write("Erster Name in Statistik:", beispielwert)

    excel_bytes = save_virtual_workbook(wb)
    st.download_button(
        label="Bearbeitete Datei herunterladen",
        data=excel_bytes,
        file_name="Dienstpläne2025_mit_Statistik.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
