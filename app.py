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
# Cambia questa password con quella che desideri
PASSWORD_CLIENTE = "Trasporti2024!"

def check_password():
    """Schermata di blocco con password"""
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

# Se la password non è inserita, il programma si ferma qui.
if not check_password():
    st.stop()


# --- 2. FUNZIONI DI SUPPORTO ---

def sanitize_filename(name):
    if pd.isna(name): return "Sconosciuto"
    return re.sub(r'[\\/*?:"<>|]', "", str(name)).strip()

def formatta_excel(writer, sheet_name):
    """Formatta l'Excel: larghezza colonne automatica e intestazione colorata."""
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # Crea lo stile per l'intestazione (Blu scuro, testo bianco, grassetto)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    
    # Applica lo stile alla prima riga
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill

    # Adatta la larghezza di ogni colonna in base al testo più lungo
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter # Prende la lettera della colonna (es. 'A')
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) # Aggiunge un po' di margine
        worksheet.column_dimensions[column].width = adjusted_width


# --- 3. INTERFACCIA APP PRINCIPALE ---

# Cerca di caricare il logo. Se hai chiamato il file "logo.png" su GitHub, apparirà.
try:
    st.image("logo.png", width=250)
except Exception:
    pass

st.title("🚛 Generatore File DDT")
st.markdown("Area riservata per la generazione automatica dei file mensili.")

uploaded_file = st.file_uploader("Carica il file Excel o CSV mensile", type=["xlsx", "csv"])

if uploaded_file is not None:
    if st.button("Elabora Dati", type="primary"):
        try:
            # Lettura file
            if uploaded_file.name.lower().endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            df.columns = df.columns.str.strip().str.upper()
            
            df['DATA DDT_DT'] = pd.to_datetime(df['DATA DDT'], errors='coerce')
            first_valid = df['DATA DDT_DT'].dropna()
            mese_str = first_valid.iloc[0].strftime("%Y-%m") if not first_valid.empty else "MESE_IGNOTO"
            df = df.drop(columns=['DATA DDT_DT'])
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                
                # --- CLIENTI ---
                cols_clienti_target = ['DATA DI CARICO', 'CLIENTE', 'DITTA DI CARICO', 'LUOGO DI CARICO', 'DESTINAZIONE', 'VARIE', 'PESO', 'TARIFFA', 'TOTALE', 'N. DDT', 'DATA DDT']
                cols_clienti = [col for col in cols_clienti_target if col in df.columns]
                
                for cliente, group in df.groupby('CLIENTE'):
                    safe_cliente = sanitize_filename(cliente)
                    filename = f"Clienti/{safe_cliente}_{mese_str}.xlsx"
                    
                    excel_buffer = io.BytesIO()
                    # Uso ExcelWriter per applicare la formattazione
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        group[cols_clienti].to_excel(writer, index=False, sheet_name='Dati_DDT')
                        formatta_excel(writer, 'Dati_DDT')
                    
                    zip_file.writestr(filename, excel_buffer.getvalue())
                
                # --- AUTISTI ---
                if 'AUTISTA AL CARICO' in df.columns:
                    for autista, group in df.groupby('AUTISTA AL CARICO'):
                        safe_autista = sanitize_filename(autista)
                        filename = f"Autisti/Autista_{safe_autista}_{mese_str}.xlsx"
                        
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            group.to_excel(writer, index=False, sheet_name='Dati_DDT')
                            formatta_excel(writer, 'Dati_DDT')
                        
                        zip_file.writestr(filename, excel_buffer.getvalue())

            st.success("✅ File elaborati e formattati con successo!")
            st.download_button(
                label="📥 Scarica Archivio ZIP (Clienti e Autisti)",
                data=zip_buffer.getvalue(),
                file_name=f"DDT_Elaborati_{mese_str}.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"Si è verificato un errore. Dettagli: {e}")
