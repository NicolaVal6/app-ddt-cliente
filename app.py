import streamlit as st
import pandas as pd
import io
import zipfile
import re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 🎨 PANNELLO DI CONFIGURAZIONE STILE EXCEL
# ==========================================
# Modifica questi valori per farli combaciare al 100% con il file del cliente

FONT_NOME = "Arial"                  # Es: "Calibri", "Arial", "Tahoma"
FONT_DIMENSIONE_DATI = 10            # Dimensione del testo normale
FONT_DIMENSIONE_SPETTABILE = 11      # Dimensione della scritta "SPETT.LE"
COLORE_SFONDO_INTESTAZIONI = "B4C6E7" # Colore celle riga 5 (D9D9D9 = Grigio Chiaro. Usa "FFFF00" per Giallo, "B4C6E7" per Azzurro)
COLORE_TESTO_INTESTAZIONI = "000000" # Colore testo riga 5 (000000 = Nero, FFFFFF = Bianco)
ALTEZZA_RIGA_SPETTABILE = 25         # Quanto deve essere alta la riga 2
ALTEZZA_RIGHE_DATI = 18              # Spaziatura verticale (altezza) delle righe della tabella

# ==========================================

st.set_page_config(page_title="Generatore DDT Formattato", layout="centered")

PASSWORD_CLIENTE = "Trasporti2024!"

def check_password():
    if "password_corretta" not in st.session_state:
        st.session_state["password_corretta"] = False
    if not st.session_state["password_corretta"]:
        st.warning("🔒 Area Riservata.")
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

def formatta_excel_custom(writer, sheet_name, is_casati=False, cliente_nome="", info_anagrafica=""):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # 1. Definizione Stili dal Pannello
    font_dati = Font(name=FONT_NOME, size=FONT_DIMENSIONE_DATI)
    font_bold = Font(name=FONT_NOME, size=FONT_DIMENSIONE_DATI, bold=True)
    font_spettabile = Font(name=FONT_NOME, size=FONT_DIMENSIONE_SPETTABILE, bold=True)
    
    header_fill = PatternFill(start_color=COLORE_SFONDO_INTESTAZIONI, end_color=COLORE_SFONDO_INTESTAZIONI, fill_type="solid")
    header_font = Font(name=FONT_NOME, size=FONT_DIMENSIONE_DATI, bold=True, color=COLORE_TESTO_INTESTAZIONI)
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # 2. Applicazione Font Generale e Altezza Righe
    for row_idx, row in enumerate(worksheet.iter_rows(), 1):
        worksheet.row_dimensions[row_idx].height = ALTEZZA_RIGHE_DATI
        for cell in row:
            cell.font = font_dati
            cell.alignment = align_left

    start_row_header = 5 if is_casati else 1
    
    if is_casati:
        # 3. Spaziature e dimensioni per l'intestazione
        worksheet.row_dimensions[2].height = ALTEZZA_RIGA_SPETTABILE
        
        testo = info_anagrafica if info_anagrafica else cliente_nome
        # Formattazione spaziata stile "CASATI" originale
        worksheet['A2'] = f"SPETT.LE   {testo}"
        worksheet['A2'].font = font_spettabile
        
        worksheet['A3'] = "RIF. NS FATT. N. ........ DEL  ..../..../202.."
        worksheet['A3'].font = font_bold

    # 4. Stile Intestazioni Tabella (Riga 5)
    for cell in worksheet[start_row_header]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = thin_border

    # 5. Formattazione Dati e Dimensioni Colonne Rigorose
    date_cols = ['DATA DDT', 'DATA DI CARICO']
    
    for col_idx, col in enumerate(worksheet.columns, 1):
        col_letter = get_column_letter(col_idx)
        header_val = str(worksheet.cell(row=start_row_header, column=col_idx).value).upper()
        
        max_len = 0
        for cell in col:
            if cell.row >= start_row_header:
                cell.border = thin_border # Griglia chiara
                
                # Formato Date e Allineamenti
                if any(x in header_val for x in date_cols) and cell.row > start_row_header:
                    cell.number_format = 'DD/MM/YYYY'
                    cell.alignment = align_center
                
                # Calcolo per allargamento dinamico ma proporzionato
                if cell.value and cell.row >= start_row_header:
                    l = len(str(cell.value))
                    if l > max_len: max_len = l
        
        # Larghezze personalizzate in base al tipo di colonna per evitare restringimenti
        if "DITTA" in header_val or "LUOGO" in header_val or "DESTINAZIONE" in header_val:
            worksheet.column_dimensions[col_letter].width = max(max_len + 2, 20)
        elif "QUANTITA" in header_val or "PESO" in header_val:
            worksheet.column_dimensions[col_letter].width = 15
        elif "DATA" in header_val:
            worksheet.column_dimensions[col_letter].width = 14
        else:
            worksheet.column_dimensions[col_letter].width = max(max_len + 2, 12)


try: st.image("logo.png", width=250)
except: pass

st.title("🚛 Gestione Trasporti (Formattato)")

uploaded_file = st.file_uploader("1. Carica File Gestionale", type=["xlsx", "csv"], key="file_principale")
uploaded_anagrafica = st.file_uploader("2. Carica Anagrafiche Clienti", type=["xlsx", "csv"], key="file_ana")

if uploaded_file and st.button("🚀 ELABORA DATI", type="primary"):
    try:
        anagrafiche = {}
        if uploaded_anagrafica:
            df_ana = pd.read_excel(uploaded_anagrafica) if uploaded_anagrafica.name.endswith('.xlsx') else pd.read_csv(uploaded_anagrafica)
            df_ana.columns = df_ana.columns.str.strip().str.upper()
            if 'CLIENTE' in df_ana.columns and 'INTESTAZIONE COMPLETA' in df_ana.columns:
                anagrafiche = dict(zip(df_ana['CLIENTE'].astype(str).str.upper(), df_ana['INTESTAZIONE COMPLETA']))

        df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
        df.columns = df.columns.str.strip().str.upper()
        df = df.loc[:, ~df.columns.str.contains('^UNNAMED', na=False)]

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
        if 'DATA DDT' in df_mapped.columns:
            df_mapped['DATA DDT'] = pd.to_datetime(df_mapped['DATA DDT'], errors='coerce').dt.date
        
        first_valid = df_mapped['DATA DDT'].dropna() if 'DATA DDT' in df_mapped.columns else pd.Series()
        mese_str = first_valid.iloc[0].strftime("%Y-%m") if not first_valid.empty else "MESE_IGNOTO"

        cols_finali = ['DITTA PRESA', 'LUOGO PRESA', 'DESTINAZIONE FINALE', 'VARIE', "QUANTITA'", 'TARIFFA', 'TOTALE', 'N. DDT', 'DATA DDT']
        cols_presenti = [c for c in cols_finali if c in df_mapped.columns]

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            
            # ESPORTAZIONE CLIENTI
            for cliente, group in df_mapped.groupby('CLIENTE'):
                safe_cliente = re.sub(r'[\\/*?:"<>|]', "", str(cliente)).strip()
                buf = io.BytesIO()
                
                with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                    data_export = group[cols_presenti]
                    data_export.to_excel(writer, index=False, sheet_name='Dati', startrow=4)
                    
                    info = anagrafiche.get(str(cliente).upper(), "")
                    formatta_excel_custom(writer, 'Dati', is_casati=True, cliente_nome=str(cliente), info_anagrafica=info)
                
                zip_file.writestr(f"Clienti/{safe_cliente}_{mese_str}.xlsx", buf.getvalue())

            # ESPORTAZIONE AUTISTI
            if 'AUTISTA AL CARICO' in df.columns:
                for autista, group in df.groupby('AUTISTA AL CARICO'):
                    safe_autista = re.sub(r'[\\/*?:"<>|]', "", str(autista)).strip()
                    buf = io.BytesIO()
                    
                    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                        group.to_excel(writer, index=False, sheet_name='Dati')
                        formatta_excel_custom(writer, 'Dati', is_casati=False)
                    
                    zip_file.writestr(f"Autisti/{safe_autista}_{mese_str}.xlsx", buf.getvalue())

        st.success("✅ File elaborati!")
        st.download_button("📥 SCARICA ZIP", zip_buffer.getvalue(), f"DDT_{mese_str}.zip", "application/zip")

    except Exception as e:
        st.error(f"Errore: {e}")
