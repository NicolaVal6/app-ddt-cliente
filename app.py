import streamlit as st
import pandas as pd
import io
import zipfile
import re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- CONFIGURAZIONE ---
st.set_page_config(page_title="Trasporti App v4.0", layout="centered")
st.caption("Versione Sistema: 4.0 (Formattazione CASATI Pixel-Perfect)")

PASSWORD_CLIENTE = "Trasporti2024!"

def check_password():
    if "password_corretta" not in st.session_state:
        st.session_state["password_corretta"] = False
    if not st.session_state["password_corretta"]:
        st.warning("🔒 Area Riservata. Inserisci la password.")
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


# --- MOTORE DI FORMATTAZIONE EXCEL ---
def formatta_excel_professionale(writer, sheet_name, is_casati=False, cliente_nome="", info_anagrafica=""):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # Definizione Stili
    font_standard = Font(name='Calibri', size=11)
    font_bold = Font(name='Calibri', size=11, bold=True)
    header_font = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    center_aligned = Alignment(horizontal="center", vertical="center")
    
    # Bordo sottile nero per la griglia
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    # Imposta il font standard su tutto il foglio per sicurezza
    for row in worksheet.iter_rows():
        for cell in row:
            cell.font = font_standard

    # Logica di Layout
    start_row_header = 5 if is_casati else 1
    
    if is_casati:
        # Iniezione Intestazione Aziendale (Righe 2 e 3)
        testo_spettabile = info_anagrafica if info_anagrafica else cliente_nome
        worksheet['A2'] = f"SPETT.LE   {testo_spettabile}"
        worksheet['A2'].font = Font(name='Calibri', size=12, bold=True)
        
        worksheet['A3'] = "RIF. NS FATT. N. ........ DEL  ..../..../202.."
        worksheet['A3'].font = font_bold

    # Applicazione Stile alle Intestazioni della Tabella
    for cell in worksheet[start_row_header]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_aligned
        cell.border = thin_border

    # Formattazione Dati e Calcolo Larghezza Colonne
    for col_idx, col in enumerate(worksheet.columns, 1):
        column_letter = get_column_letter(col_idx)
        header_val = str(worksheet.cell(row=start_row_header, column=col_idx).value).upper()
        
        max_length = 0
        for cell in col:
            # Lavora solo sulle celle della tabella (esclude Spett.le in alto)
            if cell.row >= start_row_header:
                cell.border = thin_border # Disegna la griglia
                
                # Applica il formato Data puro senza orario
                if any(x in header_val for x in ['DATA', 'DDT']) and cell.row > start_row_header:
                    cell.number_format = 'DD/MM/YYYY'
                    cell.alignment = center_aligned
                
                # Calcolo per allargare la colonna
                try:
                    if cell.value:
                        l = len(str(cell.value))
                        if l > max_length: max_length = l
                except: pass
        
        # Allarga la colonna (minimo 12, massimo flessibile)
        adjusted_width = max(max_length + 2, 12)
        worksheet.column_dimensions[column_letter].width = adjusted_width


# --- INTERFACCIA E LOGICA DATI ---
try:
    st.image("logo.png", width=250)
except:
    pass

st.title("🚛 Generatore Trasporti (Standard CASATI)")

st.subheader("1️⃣ Dati Mensili")
uploaded_file = st.file_uploader("Carica Excel/CSV dal Gestionale", type=["xlsx", "csv"], key="file_principale")

st.subheader("2️⃣ Anagrafiche Clienti (Opzionale)")
uploaded_anagrafica = st.file_uploader("Carica Excel Anagrafiche per indirizzi completi", type=["xlsx", "csv"], key="file_ana")

if uploaded_file:
    if st.button("🚀 ELABORA FILE", type="primary"):
        try:
            # 1. Parsing Anagrafiche
            anagrafiche = {}
            if uploaded_anagrafica:
                df_ana = pd.read_excel(uploaded_anagrafica) if uploaded_anagrafica.name.endswith('.xlsx') else pd.read_csv(uploaded_anagrafica)
                df_ana.columns = df_ana.columns.str.strip().str.upper()
                if 'CLIENTE' in df_ana.columns and 'INTESTAZIONE COMPLETA' in df_ana.columns:
                    anagrafiche = dict(zip(df_ana['CLIENTE'].astype(str).str.upper(), df_ana['INTESTAZIONE COMPLETA']))

            # 2. Parsing Dati Mensili
            df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
            df.columns = df.columns.str.strip().str.upper()
            df = df.loc[:, ~df.columns.str.contains('^UNNAMED', na=False)] # Pulizia colonne vuote

            # Mappatura Infallibile Colonne
            mappa = {
                'DITTA PRESA': 'DITTA PRESA', 'DITTA DI CARICO': 'DITTA PRESA',
                'LUOGO PRESA': 'LUOGO PRESA', 'LUOGO DI CARICO': 'LUOGO PRESA',
                'DESTINAZIONE FINALE': 'DESTINAZIONE FINALE', 'DESTINAZIONE': 'DESTINAZIONE FINALE',
                'TIPO MERCE': 'VARIE', 'VARIE': 'VARIE',
                'QUANTITA': "QUANTITA'", 'PESO': "QUANTITA'",
                'DDT': 'N. DDT', 'N. DDT': 'N. DDT',
                'TARIFFA': 'TARIFFA', 'TOTALE': 'TOTALE', 'DATA DDT': 'DATA DDT'
            }
            
            df_mapped = df.rename(columns=mappa)
            
            # Gestione sicura delle date
            if 'DATA DDT' in df_mapped.columns:
                df_mapped['DATA DDT'] = pd.to_datetime(df_mapped['DATA DDT'], errors='coerce').dt.date
            
            # Recupero Mese
            first_valid = df_mapped['DATA DDT'].dropna() if 'DATA DDT' in df_mapped.columns else pd.Series()
            mese_str = first_valid.iloc[0].strftime("%Y-%m") if not first_valid.empty else "MESE_IGNOTO"

            cols_finali = ['DITTA PRESA', 'LUOGO PRESA', 'DESTINAZIONE FINALE', 'VARIE', "QUANTITA'", 'TARIFFA', 'TOTALE', 'N. DDT', 'DATA DDT']
            cols_presenti = [c for c in cols_finali if c in df_mapped.columns]

            # 3. Creazione Motore ZIP in Memoria
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                
                # ELABORAZIONE CLIENTI (STILE CASATI)
                for cliente, group in df_mapped.groupby('CLIENTE'):
                    safe_cliente = re.sub(r'[\\/*?:"<>|]', "", str(cliente)).strip()
                    buf = io.BytesIO()
                    
                    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                        data_export = group[cols_presenti]
                        # startrow=4 significa che l'intestazione parte dalla riga 5 di Excel
                        data_export.to_excel(writer, index=False, sheet_name='Dati', startrow=4)
                        
                        info = anagrafiche.get(str(cliente).upper(), "")
                        formatta_excel_professionale(writer, 'Dati', is_casati=True, cliente_nome=str(cliente), info_anagrafica=info)
                    
                    zip_file.writestr(f"Clienti/{safe_cliente}_{mese_str}.xlsx", buf.getvalue())

                # ELABORAZIONE AUTISTI (STILE STANDARD PULITO)
                if 'AUTISTA AL CARICO' in df.columns:
                    for autista, group in df.groupby('AUTISTA AL CARICO'):
                        safe_autista = re.sub(r'[\\/*?:"<>|]', "", str(autista)).strip()
                        buf = io.BytesIO()
                        
                        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                            group.to_excel(writer, index=False, sheet_name='Dati')
                            formatta_excel_professionale(writer, 'Dati', is_casati=False)
                        
                        zip_file.writestr(f"Autisti/{safe_autista}_{mese_str}.xlsx", buf.getvalue())

            st.success("✅ Elaborazione CASATI completata con successo!")
            st.download_button("📥 SCARICA FILE PRONTI", zip_buffer.getvalue(), f"DDT_{mese_str}.zip", "application/zip")

        except Exception as e:
            st.error(f"Si è verificato un errore durante l'elaborazione: {e}")
