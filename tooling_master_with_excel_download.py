
import streamlit as st
import openpyxl
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Tooling Master – Daten mit Artikelnummern", layout="centered")
st.title("🧾 Werkzeugdaten – Maschinenauswahl")

excel_path = "C:/Users/pasca/Desktop/01_20230630 Tooling Master_230630.xlsx"

def is_number(value):
    try:
        float(value)
        return True
    except:
        return False

def extract_columns(ws):
    df = pd.DataFrame(ws.values)
    columns = df.iloc[0]
    columns = [f"Unnamed_{i}" if pd.isna(col) else col for i, col in enumerate(columns)]
    return columns

# Excel-Datei laden
try:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    sheet_list = wb.sheetnames

    selected_sheet = st.selectbox("📄 Tabellenblatt auswählen", sheet_list)

    if selected_sheet:
        ws = wb[selected_sheet]
        df = pd.DataFrame(ws.values)

        columns = extract_columns(ws)
        df.columns = columns

        st.write(f"**Vorschau der Daten für: {selected_sheet}**")

        beschreibungen = df.iloc[:, 0].dropna().tolist()
        st.write("🔧 **Beschreibungen:**")
        st.write(beschreibungen)

        st.write("🔧 **Daten im Originalformat:**")
        st.dataframe(df.iloc[:, 1:], use_container_width=True)

        # Excel-Datei für den Download erstellen
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=selected_sheet)
        output.seek(0)

        st.download_button(
            label="📥 Excel-Datei herunterladen",
            data=output,
            file_name=f"{selected_sheet}_Daten.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

except FileNotFoundError:
    st.error("❌ Excel-Datei wurde nicht gefunden.")
except Exception as e:
    st.error(f"❌ Fehler: {e}")
