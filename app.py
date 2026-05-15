import streamlit as st
import pandas as pd
import io
import zipfile
import re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configurazione Pagina
st.set_page_config(page_title="Trasporti App v3.1", layout="centered")

# Visualizzazione Versione per debug
st.caption("Versione Sistema: 3.1 (Ottimizzazione Gestionale + Formattazione)")

# --- 1. SICUREZZA ---
PASSWORD_CLIENTE = "Trasporti2024!"

def check_password():
    if "password_corretta" not in st.session_state:
        st.session_state["password_corretta"] = False
    if not st.session_state["password_corretta"]:
        st.warning("🔒 Inserisci la password per accedere.")
        pwd_inserita = st.text_input("Password", type="password")
        if st.button("Accedi"):
            if pwd_inserita == PASSWORD_CLIENTE:
                st.session_state["password_corretta"] = True
                st.rerun()
            else:
                st.error("Password errata.")
        return False
    return True

if not check_password():
    st.stop()

# --- 2. FUNZIONI DI FORMATTAZIONE AVANZATA ---

def formatta_excel_professionale(writer, sheet_name, is_casati=False, cliente_nome="", info_anagrafica=""):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # Font standard per tutto il foglio
    font_standard = Font(name='Calibri', size=11)
    for row in worksheet.iter_rows():
        for cell in row:
            cell.font = font_standard

    start_row_header = 5 if is_casati else 1
    
    if is_casati:
        # Intestazione Spettabile
        worksheet['A2'] = f"SPETT.LE   {info_anagrafica if info_anagrafica else cliente_nome}"
        worksheet['A2'].font = Font(name='Calibri', size=12, bold=True)
        worksheet['A3'] = "RIF. NS FATT. N. ........ DEL  ..../..../202.."
        worksheet['A3'].font = Font(name='Calibri', size=11, bold=True)

    # Stile Intestazione Tabella (Blu Scuro Casati)
    header_font = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Applica stile alla riga delle intestazioni
    for cell in worksheet[start_row_header]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = border

    # Formattazione Colonne
    for col_idx, col in enumerate(worksheet.columns, 1):
        column_letter = get_column_letter(col_idx)
        header_val = str(worksheet.cell(row=start_row_header, column=col_idx).value).upper()
        
        max_length = 0
        for cell in col:
            if cell.row >= start_row_header:
                # Applica Bordi
                cell.border = border
                
                # Formattazione Date (Rimuove orario)
                if any(x in header_val for x in ['DATA', 'DDT']):
                    cell.number_format = 'DD/MM/YYYY'
                
                # Calcolo Larghezza
                try:
                    if cell.value:
                        l = len(str(cell.value))
                        if l > max_length: max_length = l
                except: pass
        
        # Imposta larghezza minima ragionevole
        adjusted_width = max(max_length + 3, 12)
        worksheet.column_dimensions[column_letter].width = adjusted_width

# --- 3. LOGICA DI ELABORAZIONE ---

try:
    st.image("logo.png", width=250)
except:
    pass

st.title("🚛 Gestione Trasporti Automatica")

# Layout verticale per evitare che i pulsanti spariscano
st.subheader("1️⃣ Carica il file del Gestionale")
uploaded_file = st.file_uploader("File trasporti mensili (.xlsx o .csv)", type=["xlsx", "csv"], key="file_principale")

st.subheader("2️⃣ Carica Anagrafiche (Solo se vuoi indirizzi completi)")
uploaded_anagrafica = st.file_uploader("File anagrafiche clienti", type=["xlsx", "csv"], key="file_ana")

if uploaded_file:
    if st.button("🚀 ELABORA E GENERA FILE", type="primary"):
        try:
            # Lettura Anagrafiche
            anagrafiche = {}
            if uploaded_anagrafica:
                df_ana = pd.read_excel(uploaded_anagrafica) if uploaded_anagrafica.name.endswith('.xlsx') else pd.read_csv(uploaded_anagrafica)
                df_ana.columns = df_ana.columns.str.strip().str.upper()
                if 'CLIENTE' in df_ana.columns and 'INTESTAZIONE COMPLETA' in df_ana.columns:
                    anagrafiche = dict(zip(df_ana['CLIENTE'].astype(str).str.upper(), df_ana['INTESTAZIONE COMPLETA']))

            # Lettura Dati Trasporti
            df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
            df.columns = df.columns.str.strip().str.upper()
            df = df.loc[:, ~df.columns.str.contains('^UNNAMED', na=False)]

            # Mappatura Colonne Gestionale -> Target Casati
            mappa = {
                'DITTA PRESA': 'DITTA PRESA', 'DITTA DI CARICO': 'DITTA PRESA',
                'LUOGO PRESA': 'LUOGO PRESA', 'LUOGO DI CARICO': 'LUOGO PRESA',
                'DESTINAZIONE FINALE': 'DESTINAZIONE FINALE', 'DESTINAZIONE': 'DESTINAZIONE FINALE',
                'TIPO MERCE': 'VARIE', 'VARIE': 'VARIE',
                'QUANTITA': "QUANTITA'", 'PESO': "QUANTITA'",
                'DDT': 'N. DDT', 'N. DDT': 'N. DDT',
                'TARIFFA': 'TARIFFA', 'TOTALE': 'TOTALE', 'DATA DDT': 'DATA DDT'
            }
            
            # Applica Rinomina e Conversione Date
            df_mapped = df.rename(columns=mappa)
            if 'DATA DDT' in df_mapped.columns:
                df_mapped['DATA DDT'] = pd.to_datetime(df_mapped['DATA DDT'], errors='coerce').dt.date
            
            cols_finali = ['DITTA PRESA', 'LUOGO PRESA', 'DESTINAZIONE FINALE', 'VARIE', "QUANTITA'", 'TARIFFA', 'TOTALE', 'N. DDT', 'DATA DDT']
            cols_presenti = [c for c in cols_finali if c in df_mapped.columns]

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                
                # --- FILE CLIENTI ---
                for cliente, group in df_mapped.groupby('CLIENTE'):
                    safe_cliente = re.sub(r'[\\/*?:"<>|]', "", str(cliente)).strip()
                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                        data_export = group[cols_presenti]
                        data_export.to_excel(writer, index=False, sheet_name='Dati', startrow=4)
                        info = anagrafiche.get(str(cliente).upper(), "")
                        formatta_excel_professionale(writer, 'Dati', is_casati=True, cliente_nome=str(cliente), info_anagrafica=info)
                    zip_file.writestr(f"Clienti/{safe_cliente}.xlsx", buf.getvalue())

                # --- FILE AUTISTI ---
                if 'AUTISTA AL CARICO' in df.columns:
                    for autista, group in df.groupby('AUTISTA AL CARICO'):
                        safe_autista = re.sub(r'[\\/*?:"<>|]', "", str(autista)).strip()
                        buf = io.BytesIO()
                        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                            group.to_excel(writer, index=False, sheet_name='Dati')
                            formatta_excel_professionale(writer, 'Dati', is_casati=False)
                        zip_file.writestr(f"Autisti/{safe_autista}.xlsx", buf.getvalue())

            st.success("✅ Elaborazione completata!")
            st.download_button("📥 SCARICA ZIP RISULTATI", zip_buffer.getvalue(), "Trasporti_Elaborati.zip", "application/zip")

        except Exception as e:
            st.error(f"Si è verificato un errore: {e}")
