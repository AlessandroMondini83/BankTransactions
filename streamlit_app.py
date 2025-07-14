import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Elabora File Excel Multipli", layout="centered")
st.title("📂 Preparazione file per RentGer")

uploaded_files = st.file_uploader(
    "Carica uno o più file Excel", type=["xlsx"], accept_multiple_files=True
)

def is_valid_date(val):
    try:
        datetime.strptime(str(val), "%d.%m.%Y")
        return True
    except:
        return False

all_dfs = []
errore = False  # Flag per tracciare se almeno un file ha causato errore

if uploaded_files:
    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file)

            # Filtra righe con date valide nella seconda colonna
            df = df[df.iloc[:, 1].apply(is_valid_date)].copy()
           
            # Poi filtra
            df = df[df.iloc[:, 9].astype(str) != '14 - Cedole, dividendi e premi estratti'].copy()
            
            # Converte date
            df.iloc[:, 1] = df.iloc[:, 1].apply(
                lambda x: datetime.strptime(str(x), "%d.%m.%Y").strftime("%d/%m/%Y")
            )
            df.iloc[:, 2] = df.iloc[:, 2].apply(
                lambda x: x.replace('.', '/') if isinstance(x, str) and '.' in x else x
            )

            # Calcola differenza
            quantità = df.iloc[:, 5].fillna(0) - df.iloc[:, 4].fillna(0)

            # Descrizione estesa
            descrizione = (
                df.iloc[:, 18].fillna("").astype(str) + " - " +
                df.iloc[:, 10].fillna("").astype(str) + " - " +
                df.iloc[:, 11].fillna("").astype(str)
            ).str.strip(" -").str.upper()

            # Costruzione DataFrame finale
            final_df = pd.DataFrame({
                "Data 1": df.iloc[:, 1].astype(str),
                "Data 2": df.iloc[:, 2].astype(str),
                "Quantità": quantità,
                "Descrizione Estesa": descrizione,
            })
            final_df["Bilancio"] = 0
             # 🔽 Rimuove le righe con Quantità pari a -0.50 e -1.30
            final_df = final_df[final_df["Quantità"] != -0.50]
            final_df = final_df[final_df["Quantità"] != -1.25]
            final_df = final_df[final_df["Quantità"] != -1.30]
            all_dfs.append(final_df)

        except Exception:
            errore = True  # Segnala che almeno un file ha causato errore

    # Mostra messaggio di errore se necessario
    if errore:
        st.error("❌ Uno o più file non sono corretti.")

    # Unione e visualizzazione dei risultati
    if all_dfs:
        df_totale = pd.concat(all_dfs, ignore_index=True)
        st.success(f"✅ File elaborati correttamente: {len(all_dfs)}")
        st.subheader("📊 Tabella aggregata")
        st.dataframe(df_totale.head())

        # Esportazione
        st.subheader("💾 Scarica il file aggregato")
        output_file = "dati_aggregati.xlsx"
        df_totale.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button("📥 Scarica Excel", f, file_name=output_file)

st.markdown("---")
st.caption("🔧 Versione: v1.1.2 – Ultimo aggiornamento: Luglio 2025")
