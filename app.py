import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
from fpdf import FPDF
from PIL import Image

# Imposta il layout wide per utilizzare tutta la larghezza dello schermo
st.set_page_config(page_title="Report Automatico", layout="wide")
st.title("Report Automatico da File Excel")
st.write("Carica uno o più file Excel contenenti i dati per generare il report.")

# Modifica del file uploader per permettere il caricamento di più file contemporaneamente
uploaded_files = st.file_uploader("Carica i file Excel", type=["xlsx"], accept_multiple_files=True)

def process_file(file):
    """
    Legge il file Excel e processa i dati necessari:
      - Verifica la presenza delle colonne richieste.
      - Converte le colonne 'PCS' e 'Cena regularna brutto' in formato numerico.
      - Calcola il valore totale per ogni prodotto.
      - Raggruppa i dati per 'Kategoria' e calcola il numero totale di pezzi, il valore totale e il prezzo medio.
    Restituisce il dataframe elaborato, il riepilogo per categoria e le statistiche globali.
    """
    df = pd.read_excel(file, sheet_name="Sheet1")
    required_columns = ['Kategoria', 'PCS', 'Cena regularna brutto']
    if not all(col in df.columns for col in required_columns):
        raise Exception("Il file non contiene tutte le colonne richieste: 'Kategoria', 'PCS', 'Cena regularna brutto'")
    
    df['PCS'] = pd.to_numeric(df['PCS'], errors='coerce')
    df['Cena regularna brutto'] = pd.to_numeric(df['Cena regularna brutto'], errors='coerce')
    df = df.dropna(subset=required_columns)
    
    # Calcola il valore totale per prodotto
    df['Valore'] = df['Cena regularna brutto'] * df['PCS']
    
    # Raggruppa per categoria e calcola le statistiche
    grouped = df.groupby('Kategoria').agg({
        'PCS': 'sum',
        'Valore': 'sum'
    }).reset_index()
    grouped['PrezzoMedio'] = grouped['Valore'] / grouped['PCS']
    
    total_pcs = grouped['PCS'].sum()
    total_value = grouped['Valore'].sum()
    avg_price = total_value / total_pcs if total_pcs != 0 else 0
    
    global_summary = {
         'total_pcs': total_pcs,
         'total_value': total_value,
         'avg_price': avg_price
    }
    return {
         'df': df,
         'grouped': grouped,
         'global_summary': global_summary
    }

def generate_pie_chart_grouped_value(grouped, total_value):
    """
    Genera il grafico a torta che mostra la ripartizione del valore per categoria.
    """
    grouped_sorted = grouped.sort_values(by='Valore', ascending=False)
    percentages = (grouped_sorted['Valore'] / total_value) * 100
    legend_labels = [
        f"{cat.replace('gl_', '')} - {perc:.1f}%" 
        for cat, perc in zip(grouped_sorted['Kategoria'], percentages)
    ]
    fig, ax = plt.subplots(figsize=(6, 6))
    wedges, texts, autotexts = ax.pie(
         grouped_sorted['Valore'],
         labels=None,
         autopct='%1.1f%%',
         startangle=140,
         textprops={'fontsize': 8},
         wedgeprops={'linewidth': 1, 'edgecolor': 'white'}
    )
    ax.set_title("Valore per Categoria")
    ax.legend(wedges, legend_labels, title="Categoria", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
    return fig

def generate_pie_chart_grouped_qty(grouped, total_pcs):
    """
    Genera il grafico a torta che mostra la ripartizione della quantità per categoria.
    """
    grouped_sorted = grouped.sort_values(by='PCS', ascending=False)
    percentages = (grouped_sorted['PCS'] / total_pcs) * 100
    legend_labels = [
         f"{cat.replace('gl_', '')} - {perc:.1f}%" 
         for cat, perc in zip(grouped_sorted['Kategoria'], percentages)
    ]
    fig, ax = plt.subplots(figsize=(6, 6))
    wedges, texts, autotexts = ax.pie(
         grouped_sorted['PCS'],
         labels=None,
         autopct='%1.1f%%',
         startangle=140,
         textprops={'fontsize': 8},
         wedgeprops={'linewidth': 1, 'edgecolor': 'white'}
    )
    ax.set_title("Quantità per Categoria")
    ax.legend(wedges, legend_labels, title="Categoria", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
    return fig

def save_fig_to_buffer(fig):
    """
    Salva la figura matplotlib in un buffer in formato PNG.
    """
    buf = io.BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight')
    buf.seek(0)
    return buf

def generate_pdf_report(report_data, filename):
    """
    Genera un report PDF utilizzando FPDF.
    Il report include:
      - Titolo e dati globali (totale pezzi, valore totale e prezzo medio).
      - Una tabella con il riepilogo per categoria.
      - I due grafici (Valore per Categoria e Quantità per Categoria).
    
    Viene restituito il report come buffer in memoria.
    """
    global_summary = report_data['global_summary']
    grouped = report_data['grouped']
    buf1 = report_data['buf1']
    buf2 = report_data['buf2']
    
    pdf = FPDF()
    pdf.add_page()
    
    # Titolo del report
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "Report Automatico", ln=1, align='C')
    pdf.ln(10)
    
    # Dati globali
    pdf.set_font("Arial", '', 12)
    pdf.cell(0, 10, f"Totale Pezzi: {global_summary['total_pcs']}", ln=1)
    pdf.cell(0, 10, f"Valore Retail Totale: {global_summary['total_value']:.2f} EUR", ln=1)
    pdf.cell(0, 10, f"Prezzo Medio: {global_summary['avg_price']:.2f} EUR", ln=1)
    pdf.ln(10)
    
    # Riepilogo per categoria - creazione di una tabella manuale
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(50, 8, "Kategoria", border=1)
    pdf.cell(30, 8, "PCS", border=1)
    pdf.cell(40, 8, "Valore", border=1)
    pdf.cell(40, 8, "Prezzo Medio", border=1)
    pdf.ln()
    
    pdf.set_font("Arial", '', 10)
    for index, row in grouped.iterrows():
        pdf.cell(50, 8, str(row['Kategoria']), border=1)
        pdf.cell(30, 8, f"{row['PCS']}", border=1)
        pdf.cell(40, 8, f"{row['Valore']:.2f}", border=1)
        pdf.cell(40, 8, f"{row['PrezzoMedio']:.2f}", border=1)
        pdf.ln()
    pdf.ln(10)
    
    # Aggiunta del primo grafico (Valore per Categoria)
    pdf.cell(0, 10, "Grafico: Valore per Categoria", ln=1)
    image1 = Image.open(buf1)
    temp_image1 = "temp_image1.png"
    image1.save(temp_image1)
    pdf.image(temp_image1, x=None, y=None, w=pdf.w - 20)
    pdf.ln(10)
    
    # Aggiunta del secondo grafico (Quantità per Categoria)
    pdf.cell(0, 10, "Grafico: Quantità per Categoria", ln=1)
    image2 = Image.open(buf2)
    temp_image2 = "temp_image2.png"
    image2.save(temp_image2)
    pdf.image(temp_image2, x=None, y=None, w=pdf.w - 20)
    
    # Genera il PDF come buffer in memoria
    pdf_bytes = pdf.output(dest='S').encode('latin1')
    pdf_buffer = io.BytesIO(pdf_bytes)
    return pdf_buffer

if uploaded_files:
    # Per ogni file caricato, si esegue il processamento e la generazione del PDF
    for uploaded_file in uploaded_files:
        st.subheader(f"Report per il file: {uploaded_file.name}")
        try:
            # Processa il file Excel
            data = process_file(uploaded_file)
            df = data['df']
            grouped = data['grouped']
            global_summary = data['global_summary']
            
            st.write("Anteprima dei dati:")
            st.dataframe(df.head())
            
            st.write("Riepilogo Globale:")
            st.write(f"Totale Pezzi: {global_summary['total_pcs']}")
            st.write(f"Valore Retail Totale: {global_summary['total_value']:.2f} EUR")
            st.write(f"Prezzo Medio: {global_summary['avg_price']:.2f} EUR")
            
            st.write("Riepilogo per Categoria:")
            st.dataframe(grouped)
            
            # Genera i grafici
            fig1 = generate_pie_chart_grouped_value(grouped, global_summary['total_value'])
            fig2 = generate_pie_chart_grouped_qty(grouped, global_summary['total_pcs'])
            st.pyplot(fig1)
            st.pyplot(fig2)
            
            # Salva le figure in buffer
            buf1 = save_fig_to_buffer(fig1)
            buf2 = save_fig_to_buffer(fig2)
            
            # Prepara i dati per il report
            report_data = {
                'global_summary': global_summary,
                'grouped': grouped,
                'buf1': buf1,
                'buf2': buf2
            }
            
            # Genera il report PDF per questo file
            pdf_buffer = generate_pdf_report(report_data, uploaded_file.name)
            st.download_button(
                label="Scarica PDF Report",
                data=pdf_buffer,
                file_name=f"report_{uploaded_file.name}.pdf",
                mime="application/pdf"
            )
            
        except Exception as e:
            st.error(f"Errore nel processare il file {uploaded_file.name}: {e}")
else:
    st.info("Attendi il caricamento dei file Excel per generare il report.")
