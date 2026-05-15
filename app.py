import streamlit as st
import pandas as pd
import io
import zipfile
import re

# Configurazione della pagina web
st.set_page_config(page_title="Generatore Excel DDT", layout="centered")

def sanitize_filename(name):
    """Rimuove i caratteri non validi per i nomi dei file."""
    if pd.isna(name): return "Sconosciuto"
    return re.sub(r'[\\/*?:"<>|]', "", str(name)).strip()

st.title("🚛 Generatore File DDT per Clienti e Autisti")
st.markdown("Carica il file riepilogativo mensile per generare automaticamente i file separati.")

# Widget per il caricamento del file
uploaded_file = st.file_uploader("Carica il file Excel o CSV mensile", type=["xlsx", "csv"])

if uploaded_file is not None:
    if st.button("Elabora Dati", type="primary"):
        try:
            # 1. Lettura dinamica in memoria
            if uploaded_file.name.lower().endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            df.columns = df.columns.str.strip().str.upper()
            
            # 2. Estrazione Mese
            df['DATA DDT_DT'] = pd.to_datetime(df['DATA DDT'], errors='coerce')
            first_valid = df['DATA DDT_DT'].dropna()
            mese_str = first_valid.iloc[0].strftime("%Y-%m") if not first_valid.empty else "MESE_IGNOTO"
            df = df.drop(columns=['DATA DDT_DT'])
            
            # 3. Creazione dell'archivio ZIP in memoria (approccio Cloud-safe)
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                
                # --- SUDDIVISIONE CLIENTI ---
                cols_clienti_target = ['DATA DI CARICO', 'CLIENTE', 'DITTA DI CARICO', 'LUOGO DI CARICO', 'DESTINAZIONE', 'VARIE', 'PESO', 'TARIFFA', 'TOTALE', 'N. DDT', 'DATA DDT']
                cols_clienti = [col for col in cols_clienti_target if col in df.columns]
                
                for cliente, group in df.groupby('CLIENTE'):
                    safe_cliente = sanitize_filename(cliente)
                    filename = f"Clienti/{safe_cliente}_{mese_str}.xlsx"
                    
                    # Scrittura virtuale dell'Excel
                    excel_buffer = io.BytesIO()
                    group[cols_clienti].to_excel(excel_buffer, index=False, engine='openpyxl')
                    zip_file.writestr(filename, excel_buffer.getvalue())
                
                # --- SUDDIVISIONE AUTISTI ---
                if 'AUTISTA AL CARICO' in df.columns:
                    for autista, group in df.groupby('AUTISTA AL CARICO'):
                        safe_autista = sanitize_filename(autista)
                        filename = f"Autisti/Autista_{safe_autista}_{mese_str}.xlsx"
                        
                        excel_buffer = io.BytesIO()
                        group.to_excel(excel_buffer, index=False, engine='openpyxl')
                        zip_file.writestr(filename, excel_buffer.getvalue())

            # 4. Fornisce il file ZIP per il download
            st.success("✅ Elaborazione completata con successo!")
            st.download_button(
                label="📥 Scarica Archivio ZIP (Clienti e Autisti)",
                data=zip_buffer.getvalue(),
                file_name=f"DDT_Elaborati_{mese_str}.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"Si è verificato un errore durante la lettura del file. Dettagli tecnici: {e}")
