import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import requests
from bs4 import BeautifulSoup
import time
import random
import re

# Imposta il layout wide per utilizzare tutta la larghezza dello schermo
st.set_page_config(page_title="Report Automatico", layout="wide")

st.title("Report Automatico da File Excel")
st.write("Carica un file Excel contenente i dati per generare il report.")

@st.cache_data(show_spinner=False)
def get_product_weight_from_url(asin):
    """
    Effettua il web scraping della pagina Amazon per estrarre il peso del prodotto.
    Prova diverse strategie:
      1. Cerca in tabelle con ID noti scansionando tutte le righe.
      2. Fallback: cerca nella sezione detailBullets_feature_div.
    Viene aggiunto un delay casuale per ridurre il rischio di blocchi.
    """
    url = f"https://www.amazon.it/dp/{asin}"
    headers = {
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/90.0.4430.93 Safari/537.36"),
        "Accept-Language": "it-IT,it;q=0.9",
        "Referer": "https://wwww.amazon.it/"
    }
    try:
        response = requests.get(url, headers=headers, timeout=10)
        time.sleep(random.uniform(1, 2))  # Delay per evitare blocchi
        if response.status_code != 200:
            return None
        
        soup = BeautifulSoup(response.content, "html.parser")
        
        # Lista di possibili ID per le tabelle contenenti le specifiche
        table_ids = ["productDetails_techSpec_section_1", "productDetails_detailBullets_sections1"]
        for tid in table_ids:
            table = soup.find("table", id=tid)
            if table:
                rows = table.find_all("tr")
                for row in rows:
                    cells = row.find_all("td")
                    for cell in cells:
                        text = cell.get_text(separator=" ", strip=True)
                        if "kg" in text:
                            match = re.search(r"([\d,.]+)\s*kg", text, re.IGNORECASE)
                            if match:
                                peso_str = match.group(1).replace(",", ".")
                                try:
                                    return float(peso_str)
                                except ValueError:
                                    continue
        # Fallback: cerca nella sezione detailBullets_feature_div
        detail_div = soup.find("div", id="detailBullets_feature_div")
        if detail_div:
            bullets = detail_div.find_all("span", class_="a-list-item")
            for bullet in bullets:
                text = bullet.get_text(separator=" ", strip=True)
                if "kg" in text:
                    match = re.search(r"([\d,.]+)\s*kg", text, re.IGNORECASE)
                    if match:
                        peso_str = match.group(1).replace(",", ".")
                        try:
                            return float(peso_str)
                        except ValueError:
                            continue
        return None
    except Exception as e:
        print(f"Errore per ASIN {asin}: {e}")
        return None

# Carica il file Excel tramite l'interfaccia web
uploaded_file = st.file_uploader("Carica il file Excel", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Legge il file Excel (modifica il nome del foglio se necessario)
        df = pd.read_excel(uploaded_file, sheet_name="Sheet1")
        
        st.subheader("Anteprima dei dati")
        st.dataframe(df.head())

        # Verifica che le colonne richieste siano presenti
        required_columns = ['Kategoria', 'PCS', 'Cena regularna brutto']
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            st.error(f"Le seguenti colonne sono mancanti nel file: {', '.join(missing_cols)}")
        else:
            # Converte le colonne in formato numerico e rimuove eventuali righe con dati mancanti
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

            st.subheader("Riepilogo Globale")
            st.write(f"**Totale Pezzi:** {total_pcs}")
            st.write(f"**Valore Retail Totale:** {total_value:.2f} EUR")
            st.write(f"**Prezzo Medio:** {avg_price:.2f} EUR")

            st.subheader("Riepilogo per Categoria")
            st.dataframe(grouped)

            # Se la colonna 'Kod 2' (ASIN) è presente, tenta di ottenere il peso
            if 'Kod 2' in df.columns:
                st.subheader("Informazioni sul Peso dei Prodotti")
                # Creiamo una lista per i risultati dei pesi
                weight_results = []
                n = len(df)
                progress_bar = st.progress(0)
                progress_text = st.empty()
                
                # Uso uno spinner per indicare l'elaborazione in corso
                with st.spinner("Recupero dei pesi in corso..."):
                    for i, asin in enumerate(df['Kod 2']):
                        peso = get_product_weight_from_url(asin)
                        weight_results.append(peso)
                        progress_bar.progress((i + 1) / n)
                        progress_text.text(f"Elaborati {i + 1} di {n} prodotti")
                
                progress_text.empty()  # Rimuove il messaggio di avanzamento
                df['Peso'] = weight_results
                st.dataframe(df[['Kod 2', 'Peso']].head(10))
                
                # Calcola statistiche sul peso (escludendo valori nulli)
                peso_validi = pd.to_numeric(df['Peso'], errors='coerce').dropna()
                if not peso_validi.empty:
                    peso_totale = peso_validi.sum()
                    peso_medio = peso_validi.mean()
                    st.write(f"**Peso Totale:** {peso_totale:.2f} kg")
                    st.write(f"**Peso Medio:** {peso_medio:.2f} kg")
                else:
                    st.warning("Non sono stati trovati dati di peso validi per i prodotti.")
            else:
                st.warning("La colonna 'Kod 2' (ASIN) non è presente nel file.")

            # Grafici affiancati
            st.subheader("Grafici Affiancati")
            col1, col2 = st.columns(2)
            
            # ----- GRAFICO 1: Valore per Categoria -----
            with col1:
                st.markdown("**Ripartizione Valore per Categoria**")
                grouped_sorted_value = grouped.sort_values(by='Valore', ascending=False)
                percentages_value = (grouped_sorted_value['Valore'] / total_value) * 100
                legend_labels_value = [
                    f"{cat.replace('gl_', '')} - {perc:.1f}%" 
                    for cat, perc in zip(grouped_sorted_value['Kategoria'], percentages_value)
                ]
                fig1, ax1 = plt.subplots(figsize=(6, 6))
                wedges1, texts1, autotexts1 = ax1.pie(
                    grouped_sorted_value['Valore'],
                    labels=None,
                    autopct='%1.1f%%',
                    startangle=140,
                    textprops={'fontsize': 8},
                    wedgeprops={'linewidth': 1, 'edgecolor': 'white'}
                )
                ax1.set_title("Valore per Categoria")
                ax1.legend(
                    wedges1,
                    legend_labels_value,
                    title="Categoria",
                    loc="center left",
                    bbox_to_anchor=(1, 0, 0.5, 1)
                )
                st.pyplot(fig1)

            # ----- GRAFICO 2: Quantità per Categoria -----
            with col2:
                st.markdown("**Ripartizione Quantità per Categoria**")
                grouped_sorted_qty = grouped.sort_values(by='PCS', ascending=False)
                percentages_qty = (grouped_sorted_qty['PCS'] / total_pcs) * 100
                legend_labels_qty = [
                    f"{cat.replace('gl_', '')} - {perc:.1f}%" 
                    for cat, perc in zip(grouped_sorted_qty['Kategoria'], percentages_qty)
                ]
                fig2, ax2 = plt.subplots(figsize=(6, 6))
                wedges2, texts2, autotexts2 = ax2.pie(
                    grouped_sorted_qty['PCS'],
                    labels=None,
                    autopct='%1.1f%%',
                    startangle=140,
                    textprops={'fontsize': 8},
                    wedgeprops={'linewidth': 1, 'edgecolor': 'white'}
                )
                ax2.set_title("Quantità per Categoria")
                ax2.legend(
                    wedges2,
                    legend_labels_qty,
                    title="Categoria",
                    loc="center left",
                    bbox_to_anchor=(1, 0, 0.5, 1)
                )
                st.pyplot(fig2)

    except Exception as e:
        st.error(f"Errore nel processare il file: {e}")
else:
    st.info("Attendi il caricamento del file Excel per generare il report.")
