import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import requests
from bs4 import BeautifulSoup
import time
import random
import re
from concurrent.futures import ThreadPoolExecutor, as_completed

st.set_page_config(page_title="Report Automatico", layout="wide")
st.title("Report Automatico da File Excel")
st.write("Carica un file Excel contenente i dati per generare il report.")

def get_product_weight_from_url(asin, session, max_retries=3):
url = f"https://www.amazon.it/dp/{asin}?th=1"
    headers = {
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/90.0.4430.93 Safari/537.36"),
        "Accept-Language": "it-IT,it;q=0.9",
        "Referer": "https://www.amazon.it/"
    }
    delay = 2  # inizia con 2 secondi
    for attempt in range(max_retries):
        try:
            response = session.get(url, headers=headers, timeout=10)
            st.write(f"ASIN {asin}: Tentativo {attempt+1}, HTTP Status Code: {response.status_code}")
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, "html.parser")
                table_ids = ["productDetails_techSpec_section_1", "productDetails_detailBullets_sections1"]
                for tid in table_ids:
                    table = soup.find("table", id=tid)
                    if table:
                        rows = table.find_all("tr")
                        for row in rows:
                            cells = row.find_all("td")
                            for cell in cells:
                                text = cell.get_text(separator=" ", strip=True)
                                if "kg" in text.lower():
                                    match = re.search(r"([\d,.]+)\s*kg", text, re.IGNORECASE)
                                    if match:
                                        peso_str = match.group(1).replace(",", ".")
                                        try:
                                            return float(peso_str)
                                        except ValueError:
                                            continue
                detail_div = soup.find("div", id="detailBullets_feature_div")
                if detail_div:
                    bullets = detail_div.find_all("span", class_="a-list-item")
                    for bullet in bullets:
                        text = bullet.get_text(separator=" ", strip=True)
                        if "kg" in text.lower():
                            match = re.search(r"([\d,.]+)\s*kg", text, re.IGNORECASE)
                            if match:
                                peso_str = match.group(1).replace(",", ".")
                                try:
                                    return float(peso_str)
                                except ValueError:
                                    continue
                return None
            else:
                st.write(f"Errore HTTP {response.status_code} per ASIN {asin}. Riprovo...")
                time.sleep(delay)
                delay *= 2  # Backoff esponenziale
        except Exception as e:
            st.write(f"Errore per ASIN {asin} nel tentativo {attempt+1}: {e}")
            time.sleep(delay)
            delay *= 2
    return None

@st.cache_data(show_spinner=False)
def get_weight_for_asin_list(asins):
    weight_results = {}
    with requests.Session() as session:
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = {executor.submit(get_product_weight_from_url, asin, session): idx
                       for idx, asin in enumerate(asins)}
            for future in as_completed(futures):
                idx = futures[future]
                try:
                    weight = future.result()
                except Exception:
                    weight = None
                weight_results[idx] = weight
    return [weight_results[i] for i in sorted(weight_results.keys())]

uploaded_file = st.file_uploader("Carica il file Excel", type=["xlsx"])
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Sheet1")
        st.subheader("Anteprima dei dati")
        st.dataframe(df.head())
        required_columns = ['Kategoria', 'PCS', 'Cena regularna brutto']
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            st.error(f"Le seguenti colonne sono mancanti: {', '.join(missing_cols)}")
        else:
            df['PCS'] = pd.to_numeric(df['PCS'], errors='coerce')
            df['Cena regularna brutto'] = pd.to_numeric(df['Cena regularna brutto'], errors='coerce')
            df = df.dropna(subset=required_columns)
            df['Valore'] = df['Cena regularna brutto'] * df['PCS']
            grouped = df.groupby('Kategoria').agg({'PCS': 'sum', 'Valore': 'sum'}).reset_index()
            grouped['PrezzoMedio'] = grouped['Valore'] / grouped['PCS']
            total_pcs = grouped['PCS'].sum()
            total_value = grouped['Valore'].sum()
            avg_price = total_value / total_pcs if total_pcs != 0 else 0
            st.subheader("Riepilogo Globale")
            st.write(f"**Totale Pezzi:** {total_pcs}")
            st.write(f"**Valore Retail Totale:** {total_value:.2f} EUR")
            st.write(f"**Prezzo Medio:** {avg_price:.2f} EUR")
            st.subheader("Riepilogo per Categoria")
            st.dataframe(grouped)
            if 'Kod 2' in df.columns:
                st.subheader("Informazioni sul Peso dei Prodotti")
                n = len(df)
                progress_bar = st.progress(0)
                progress_text = st.empty()
                asins = df['Kod 2'].tolist()
                weights = []
                with st.spinner("Recupero dei pesi in corso..."):
                    weights = get_weight_for_asin_list(asins)
                    for i in range(n):
                        progress_bar.progress((i + 1) / n)
                        progress_text.text(f"Elaborati {i + 1} di {n} prodotti")
                progress_text.empty()
                df['Peso'] = weights
                st.dataframe(df[['Kod 2', 'Peso']].head(10))
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
            st.subheader("Grafici Affiancati")
            col1, col2 = st.columns(2)
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
