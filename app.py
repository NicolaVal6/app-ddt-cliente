import streamlit as st
import pandas as pd
import io
import zipfile
import re
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# Configurazione della pagina web
st.set_page_config(page_title="Generatore Excel DDT", layout="centered")

# --- 1. SISTEMA DI SICUREZZA (PASSWORD) ---
PASSWORD_CLIENTE = "Trasporti2024!"

def check_password():
    if "password_corretta" not in st.session_state:
        st.session_state["password_corretta"] = False
    if not st.session_state["password_corretta"]:
        st.warning("🔒 Area Riservata. Inserisci la password per accedere.")
        pwd_inserita = st.text_input("Password", type="password")
        if st.button("Accedi"):
            if pwd_inserita == PASSWORD_CLIENTE:
                st.session_state["password_corretta"] = True
                st.rerun()
            else:
                st.error("Password errata. Riprova.")
        return False
    return True

if not check_password():
    st.stop()

# --- 2. FUNZIONI DI SUPPORTO ---

def sanitize_filename(name):
    if pd.isna(name): return "Sconosciuto"
    return re.sub(r'[\\/*?:"<>|]', "", str(name)).strip()

def formatta_excel(writer, sheet_name):
    """Formatta l'Excel: larghezza automatica, intestazione e formato DATA."""
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    
    # Elenco delle colonne che devono essere formattate come data
    date_cols_to_format = ['DATA DI CARICO', 'DATA DDT']

    for col_idx, col in enumerate(worksheet.columns, 1):
        max_length = 0
        column_letter = get_column_letter(col_idx)
        header_value = col[0].value # Il nome della colonna (es. "DATA DDT")

        # Stile Intestazione
        col[0].font = header_font
        col[0].fill = header_fill

        for cell in col:
            # Se la colonna è una di quelle data, applica il formato numerico Excel
            if header_value in date_cols_to_format and cell.row > 1:
                cell.number_format = 'DD/MM/YYYY'
            
            try:
                if cell.value:
                    # Se è una data, usiamo una stringa standard per calcolare la larghezza
                    val_str = str(cell.value)
                    if len(val_str) > max_length:
                        max_length = len(val_str)
            except:
                pass
        
        # Allarga la colonna (margine di sicurezza +2)
        adjusted_width = (max_length + 2) if max_length > 10 else 15
        worksheet.column_dimensions[column_letter].width = adjusted_width

# --- 3. INTERFACCIA APP PRINCIPALE ---

try:
    st.image("Valsecchi_trasporti_logo.png", width=250)
except Exception:
    pass

st.title("🚛 Generatore File DDT")

uploaded_file = st.file_uploader("Carica il file Excel o CSV mensile", type=["xlsx", "csv"])

if uploaded_file is not None:
    if st.button("Elabora Dati", type="primary"):
        try:
            # Lettura e pulizia nomi colonne
            if uploaded_file.name.lower().endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            df.columns = df.columns.str.strip().str.upper()

            # CONVERSIONE DATE: Trasformiamo le colonne in vero formato Data
            # Questo evita che Excel le veda come semplici scritte (stringhe)
            date_columns = ['DATA DI CARICO', 'DATA DDT']
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
            
            # Recupero mese per il nome del file ZIP
            first_valid = df['DATA DDT'].dropna()
            mese_str = first_valid.iloc[0].strftime("%Y-%m") if not first_valid.empty else "MESE_IGNOTO"
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                
                # --- LOGICA CLIENTI ---
                cols_clienti_target = ['DATA DI CARICO', 'CLIENTE', 'DITTA DI CARICO', 'LUOGO DI CARICO', 'DESTINAZIONE', 'VARIE', 'PESO', 'TARIFFA', 'TOTALE', 'N. DDT', 'DATA DDT']
                cols_clienti = [col for col in cols_clienti_target if col in df.columns]
                
                for cliente, group in df.groupby('CLIENTE'):
                    safe_cliente = sanitize_filename(cliente)
                    filename = f"Clienti/{safe_cliente}_{mese_str}.xlsx"
                    
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        # Rimuoviamo il fuso orario se presente per evitare errori in Excel
                        group_to_export = group[cols_clienti].copy()
                        group_to_export.to_excel(writer, index=False, sheet_name='Dati_DDT')
                        formatta_excel(writer, 'Dati_DDT')
                    
                    zip_file.writestr(filename, excel_buffer.getvalue())
                
                # --- LOGICA AUTISTI ---
                if 'AUTISTA AL CARICO' in df.columns:
                    for autista, group in df.groupby('AUTISTA AL CARICO'):
                        safe_autista = sanitize_filename(autista)
                        filename = f"Autisti/Autista_{safe_autista}_{mese_str}.xlsx"
                        
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            group.to_excel(writer, index=False, sheet_name='Dati_DDT')
                            formatta_excel(writer, 'Dati_DDT')
                        
                        zip_file.writestr(filename, excel_buffer.getvalue())

            st.success("✅ Elaborazione completata!")
            st.download_button(
                label="📥 Scarica Archivio ZIP",
                data=zip_buffer.getvalue(),
                file_name=f"DDT_Elaborati_{mese_str}.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"Errore: {e}")
