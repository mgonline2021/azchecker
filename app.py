import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.title("Report Automatico da File Excel")
st.write("Carica un file Excel contenente i dati per generare il report.")

# Carica il file Excel tramite l'interfaccia web
uploaded_file = st.file_uploader("Carica il file Excel", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Leggi il file Excel (modifica il nome del foglio se necessario)
        df = pd.read_excel(uploaded_file, sheet_name="Sheet1")
        
        st.subheader("Anteprima dei dati")
        st.dataframe(df.head())

        # Verifica che le colonne richieste siano presenti
        required_columns = ['Kategoria', 'PCS', 'Cena regularna brutto']
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            st.error(f"Le seguenti colonne sono mancanti nel file: {', '.join(missing_cols)}")
        else:
            # Converti le colonne in formato numerico
            df['PCS'] = pd.to_numeric(df['PCS'], errors='coerce')
            df['Cena regularna brutto'] = pd.to_numeric(df['Cena regularna brutto'], errors='coerce')
            df = df.dropna(subset=required_columns)

            # Calcola il valore totale per ogni prodotto
            df['Valore'] = df['Cena regularna brutto'] * df['PCS']

            # Raggruppa per categoria e calcola totali e prezzo medio
            grouped = df.groupby('Kategoria').agg({
                'PCS': 'sum',
                'Valore': 'sum'
            }).reset_index()
            grouped['PrezzoMedio'] = grouped['Valore'] / grouped['PCS']

            # Calcola i totali globali
            total_pcs = grouped['PCS'].sum()
            total_value = grouped['Valore'].sum()
            avg_price = total_value / total_pcs if total_pcs != 0 else 0

            # Visualizza il riepilogo globale
            st.subheader("Riepilogo Globale")
            st.write(f"**Totale Pezzi:** {total_pcs}")
            st.write(f"**Valore Retail Totale:** {total_value:.2f} EUR")
            st.write(f"**Prezzo Medio:** {avg_price:.2f} EUR")

            # Visualizza il riepilogo per categoria
            st.subheader("Riepilogo per Categoria")
            st.dataframe(grouped)

            # Crea il grafico a torta per la ripartizione del valore per categoria
            st.subheader("Grafico: Ripartizione Valore per Categoria")
            fig, ax = plt.subplots(figsize=(6, 6))
            ax.pie(grouped['Valore'], labels=grouped['Kategoria'], autopct='%1.1f%%', startangle=140)
            ax.set_title("Ripartizione Valore per Categoria")
            st.pyplot(fig)

    except Exception as e:
        st.error(f"Errore nel processare il file: {e}")
else:
    st.info("Attendi il caricamento del file Excel per generare il report.")
