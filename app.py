import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
from fpdf import FPDF
from PIL import Image
import zipfile

# Imposta il layout wide per utilizzare tutta la larghezza dello schermo
st.set_page_config(page_title="Report Automatico", layout="wide")
st.title("Report Automatico da File Excel")
st.write("Carica uno o più file Excel contenenti i dati per generare il report.")

# File uploader per il caricamento multiplo dei file Excel
uploaded_files = st.file_uploader("Carica i file Excel", type=["xlsx"], accept_multiple_files=True)

# Placeholder per il pulsante di download ZIP, posizionato in alto
zip_btn_placeholder = st.empty()

# Se sono presenti file caricati, creiamo un indicatore di avanzamento complessivo
if uploaded_files:
    total_files = len(uploaded_files)
    st.write(f"{total_files} file caricati. Inizio del processamento...")
    overall_progress = st.progress(0)
    overall_progress_text = st.empty()
    
    # Variabili per il report aggregato complessivo
    agg_total_pcs = 0
    agg_total_value = 0.0
    # Lista per salvare tutti i PDF generati (nome, buffer)
    all_pdf_reports = []
    
    # Funzione per processare ciascun file
    def process_file(file):
        """
        Legge il file Excel e processa i dati:
          - Verifica la presenza delle colonne richieste.
          - Converte 'PCS' e 'Cena regularna brutto' in numerico.
          - Calcola il valore totale per riga.
          - Raggruppa per 'Kategoria' e calcola statistiche: totale pezzi, valore totale e prezzo medio.
        Restituisce il dataframe elaborato, il riepilogo per categoria e le statistiche globali.
        """
        df = pd.read_excel(file, sheet_name="Sheet1")
        required_columns = ['Kategoria', 'PCS', 'Cena regularna brutto']
        if not all(col in df.columns for col in required_columns):
            raise Exception("Il file non contiene tutte le colonne richieste: 'Kategoria', 'PCS', 'Cena regularna brutto'")
        
        df['PCS'] = pd.to_numeric(df['PCS'], errors='coerce')
        df['Cena regularna brutto'] = pd.to_numeric(df['Cena regularna brutto'], errors='coerce')
        df = df.dropna(subset=required_columns)
        
        df['Valore'] = df['Cena regularna brutto'] * df['PCS']
        
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
    
    def generate_pdf_report(report_data):
        """
        Genera un report PDF utilizzando FPDF.
        Il report include:
          - Titolo e dati globali.
          - Una tabella con il riepilogo per categoria.
          - I due grafici (Valore e Quantità per Categoria).
        Viene restituito il report come buffer.
        """
        global_summary = report_data['global_summary']
        grouped = report_data['grouped']
        buf1 = report_data['buf1']
        buf2 = report_data['buf2']
        
        pdf = FPDF()
        pdf.add_page()
        
        # Titolo
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "Report Automatico", ln=1, align='C')
        pdf.ln(10)
        
        # Dati globali
        pdf.set_font("Arial", '', 12)
        pdf.cell(0, 10, f"Totale Pezzi: {global_summary['total_pcs']}", ln=1)
        pdf.cell(0, 10, f"Valore Retail Totale: {global_summary['total_value']:.2f} EUR", ln=1)
        pdf.cell(0, 10, f"Prezzo Medio: {global_summary['avg_price']:.2f} EUR", ln=1)
        pdf.ln(10)
        
        # Tabella riepilogo per categoria
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
        
        # Grafico: Valore per Categoria
        pdf.cell(0, 10, "Grafico: Valore per Categoria", ln=1)
        image1 = Image.open(buf1)
        temp_image1 = "temp_image1.png"
        image1.save(temp_image1)
        pdf.image(temp_image1, x=None, y=None, w=pdf.w - 20)
        pdf.ln(10)
        
        # Grafico: Quantità per Categoria
        pdf.cell(0, 10, "Grafico: Quantità per Categoria", ln=1)
        image2 = Image.open(buf2)
        temp_image2 = "temp_image2.png"
        image2.save(temp_image2)
        pdf.image(temp_image2, x=None, y=None, w=pdf.w - 20)
        
        pdf_bytes = pdf.output(dest='S').encode('latin1')
        pdf_buffer = io.BytesIO(pdf_bytes)
        return pdf_buffer

    # Elaborazione di ogni file
    for i, uploaded_file in enumerate(uploaded_files):
        overall_progress_text.text(f"Elaborato {i+1} di {total_files} file...")
        try:
            st.subheader(f"Report per il file: {uploaded_file.name}")
            data = process_file(uploaded_file)
            df = data['df']
            grouped = data['grouped']
            global_summary = data['global_summary']
            
            # Visualizzazione dati singoli
            st.write("Anteprima dei dati:")
            st.dataframe(df.head())
            st.write("Riepilogo Globale:")
            st.write(f"Totale Pezzi: {global_summary['total_pcs']}")
            st.write(f"Valore Retail Totale: {global_summary['total_value']:.2f} EUR")
            st.write(f"Prezzo Medio: {global_summary['avg_price']:.2f} EUR")
            st.write("Riepilogo per Categoria:")
            st.dataframe(grouped)
            
            # Generazione grafici
            fig1 = generate_pie_chart_grouped_value(grouped, global_summary['total_value'])
            fig2 = generate_pie_chart_grouped_qty(grouped, global_summary['total_pcs'])
            st.pyplot(fig1)
            st.pyplot(fig2)
            
            # Salvataggio dei grafici in buffer
            buf1 = save_fig_to_buffer(fig1)
            buf2 = save_fig_to_buffer(fig2)
            
            # Prepara il report PDF per il file corrente
            report_data = {
                'global_summary': global_summary,
                'grouped': grouped,
                'buf1': buf1,
                'buf2': buf2
            }
            pdf_buffer = generate_pdf_report(report_data)
            pdf_filename = f"report_{uploaded_file.name}.pdf"
            all_pdf_reports.append((pdf_filename, pdf_buffer))
            
            # Pulsante per scaricare il singolo PDF
            st.download_button(
                label="Scarica PDF Report",
                data=pdf_buffer,
                file_name=pdf_filename,
                mime="application/pdf"
            )
            
            # Aggrega i dati per il report complessivo
            agg_total_pcs += global_summary['total_pcs']
            agg_total_value += global_summary['total_value']
            
        except Exception as e:
            st.error(f"Errore nel processare il file {uploaded_file.name}: {e}")
        
        # Aggiorna la barra di avanzamento complessiva
        overall_progress.progress((i + 1) / total_files)
    
    # Calcolo del report aggregato complessivo
    overall_avg_price = agg_total_value / agg_total_pcs if agg_total_pcs != 0 else 0
    st.subheader("Report Complessivo Aggregato")
    st.write(f"**Totale Pezzi:** {agg_total_pcs}")
    st.write(f"**Valore Retail Totale:** {agg_total_value:.2f} EUR")
    st.write(f"**Prezzo Medio:** {overall_avg_price:.2f} EUR")
    
    # Genera il file ZIP contenente tutti i PDF
    if all_pdf_reports:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for filename, pdf_buf in all_pdf_reports:
                pdf_buf.seek(0)
                zip_file.writestr(filename, pdf_buf.read())
        zip_buffer.seek(0)
        # Inserisce il pulsante per scaricare lo ZIP in alto, utilizzando il placeholder creato all'inizio
        zip_btn_placeholder.download_button(
            label="Scarica tutti i PDF (ZIP)",
            data=zip_buffer,
            file_name="reports.zip",
            mime="application/zip"
        )
    
    overall_progress_text.empty()  # Rimuove il messaggio di avanzamento
else:
    st.info("Attendi il caricamento dei file Excel per generare il report.")
