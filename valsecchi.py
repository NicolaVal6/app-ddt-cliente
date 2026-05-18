import streamlit as st
import pandas as pd
import io
import zipfile
import re
import os
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# --- CONFIGURAZIONE DI SISTEMA ---
st.set_page_config(page_title="Trasporti App v7.0", layout="centered")
st.caption("Versione Sistema: 7.0 (Specifiche Valsecchi Layout Definitivo)")

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


# --- MOTORE DI FORMATTAZIONE AVANZATO EXCEL ---
def formatta_excel_valsecchi(writer, sheet_name, is_casati=False, cliente_nome="", info_anagrafica=""):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # Stili Rigorosi (Tutto l'allegato rigorosamente in Grassetto)
    font_bold_standard = Font(name='Calibri', size=11, bold=True, color="000000")
    font_bold_calibri12 = Font(name='Calibri', size=12, bold=True, color="000000")
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    
    # Sfondo Azzurrino per l'intestazione della tabella
    fill_azzurrino = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
    )

    start_row_header = 7 if is_casati else 1
    
    # Configurazione Piè di pagina dinamica per la stampa
    worksheet.oddFooter.center.text = "Pagina &P"

    if is_casati:
        # Ripetizione delle righe di intestazione (da 1 a 7) su ogni singola pagina stampata
        worksheet.print_title_rows = '1:7'
        
        # 1. Ingrandimento casella A1 per far alloggiare correttamente il logo
        worksheet.row_dimensions[1].height = 45
        worksheet.column_dimensions['A'].width = 18
        
        # Iniezione automatica Logo in A1
        if os.path.exists("logo.png"):
            try:
                from openpyxl.drawing.image import Image as OpenpyxlImage
                img = OpenpyxlImage("logo.png")
                img.width = 110
                img.height = 40
                worksheet.add_image(img, 'A1')
            except:
                pass
        
        # 2. Intestazione Aziendale Valsecchi a DESTRA (Formato Calibri Dim 12, a capo per riga)
        azienda_lines = [
            "VALSECCHI TRASPORTI S.R.L.",
            "VIA GIUSEPPE PARINI N. 8",
            "CESANA BRIANZA (LC)",
            "P.IVA 01932020132"
        ]
        for idx, line in enumerate(azienda_lines, 1):
            cell = worksheet.cell(row=idx, column=9) # Colonna 9 = I (Estrema destra della tabella)
            cell.value = line
            cell.font = font_bold_calibri12
            cell.alignment = align_right
            worksheet.row_dimensions[idx].height = 18
        
        # 3. Spett.le Cliente a SINISTRA (Formato Calibri Dim 12, a capo)
        testo_spettabile = info_anagrafica if info_anagrafica else cliente_nome
        cell_spett = worksheet['A4']
        cell_spett.value = f"SPETT.LE   {testo_spettabile}"
        cell_spett.font = font_bold_calibri12
        cell_spett.alignment = align_left
        worksheet.row_dimensions[4].height = 20
        
        # 4. Riferimento Fattura (Riga 5)
        cell_fatt = worksheet['A5']
        cell_fatt.value = "RIF. NS FATT. N. ........ DEL  ..../..../202.."
        cell_fatt.font = font_bold_standard
        cell_fatt.alignment = align_left
        worksheet.row_dimensions[5].height = 18
        
        # Nota: La riga 6 rimane vuota ("linea di spazio") automaticamente grazie a startrow=6 in pandas

    # 5. Formattazione Griglia Dati (Tutto Grassetto, Tutto Centrato)
    for row in worksheet.iter_rows(min_row=start_row_header):
        for cell in row:
            cell.font = font_bold_standard
            cell.border = thin_border
            cell.alignment = align_center

    # Colore Azzurrino applicato solo alla riga di intestazione della tabella
    for cell in worksheet[start_row_header]:
        cell.fill = fill_azzurrino

    # Formattazione Numerica (Due decimali) e Date pulite
    for col_idx in range(1, worksheet.max_column + 1):
        col_letter = get_column_letter(col_idx)
        header_val = str(worksheet.cell(row=start_row_header, column=col_idx).value).upper()
        
        for row_idx in range(start_row_header + 1, worksheet.max_row + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            if cell.value is not None:
                if any(x in header_val for x in ['TARIFFA', 'TOTALE']):
                    cell.number_format = '#,##0.00'
                if any(x in header_val for x in ['DATA', 'DDT']) and not isinstance(cell.value, str):
                    cell.number_format = 'DD/MM/YYYY'
        
        # Auto-fit colonne proporzionato per evitare testi troncati
        max_len = max(len(str(cell.value or '')) for cell in worksheet[col_letter] if cell.row >= start_row_header)
        worksheet.column_dimensions[col_letter].width = max(max_len + 4, 16)

    # 6. CALCOLO AUTOMATICO DEL TOTALE DELLE SOMME SOTTO LA COLONNA "DITTA PRESA"
    if is_casati:
        totale_col_idx = None
        for col_idx in range(1, worksheet.max_column + 1):
            if "TOTALE" in str(worksheet.cell(row=start_row_header, column=col_idx).value).upper():
                totale_col_idx = col_idx
                break
        
        if totale_col_idx:
            last_data_row = worksheet.max_row
            total_row_idx = last_data_row + 1
            
            # Scritta TOTALE rigorosamente sotto la colonna DITTA PRESA (Colonna 1 / A)
            cell_label = worksheet.cell(row=total_row_idx, column=1)
            cell_label.value = "TOTALE"
            
            # Iniezione della formula SUM dinamica sulla colonna TOTALE (Formato due decimali)
            col_letter = get_column_letter(totale_col_idx)
            cell_sum = worksheet.cell(row=total_row_idx, column=totale_col_idx)
            cell_sum.value = f"=SUM({col_letter}8:{col_letter}{last_data_row})"
            cell_sum.number_format = '#,##0.00'
            
            # Applica lo stile griglia e grassetto a tutta la riga finale dei totali
            for col_idx in range(1, worksheet.max_column + 1):
                c = worksheet.cell(row=total_row_idx, column=col_idx)
                c.font = font_bold_standard
                c.border = thin_border
                c.alignment = align_center


# --- INTERFACCIA WEB STREAMLIT ---
try: st.image("logo.png", width=250)
except: pass

st.title("🚛 Sistema Gestionale Valsecchi Trasporti")

uploaded_file = st.file_uploader("1. Carica Estrazione Gestionale (.xlsx o .csv)", type=["xlsx", "csv"])
uploaded_anagrafica = st.file_uploader("2. Carica Anagrafiche Clienti (Opzionale)", type=["xlsx", "csv"])

if uploaded_file and st.button("🚀 GENERA DOCUMENTI FISCALI", type="primary"):
    try:
        # Parsing Anagrafiche
        anagrafiche = {}
        if uploaded_anagrafica:
            df_ana = pd.read_excel(uploaded_anagrafica) if uploaded_anagrafica.name.endswith('.xlsx') else pd.read_csv(uploaded_anagrafica)
            df_ana.columns = df_ana.columns.str.strip().str.upper()
            if 'CLIENTE' in df_ana.columns and 'INTESTAZIONE COMPLETA' in df_ana.columns:
                anagrafiche = dict(zip(df_ana['CLIENTE'].astype(str).str.upper(), df_ana['INTESTAZIONE COMPLETA']))

        # Lettura file trasporti principale
        df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
        df.columns = df.columns.str.strip().str.upper()
        df = df.loc[:, ~df.columns.str.contains('^UNNAMED', na=False)]

        # Traduzione e uniformazione dei campi del gestionale
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

        # ORDINE TASSATIVO COLONNE (Richiesto al punto 5)
        cols_ordinate = ['DITTA PRESA', 'LUOGO PRESA', 'DESTINAZIONE FINALE', 'VARIE', "QUANTITA'", 'TARIFFA', 'TOTALE', 'N. DDT', 'DATA DDT']
        cols_presenti = [c for c in cols_ordinate if c in df_mapped.columns]

        first_valid = df_mapped['DATA DDT'].dropna() if 'DATA DDT' in df_mapped.columns else pd.Series()
        mese_str = first_valid.iloc[0].strftime("%Y-%m") if not first_valid.empty else "MESE_IGNOTO"

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            
            # --- 1. GENERAZIONE EXCEL CLIENTI ---
            for cliente, group in df_mapped.groupby('CLIENTE'):
                safe_cliente = re.sub(r'[\\/*?:"<>|]', "", str(cliente)).strip()
                buf = io.BytesIO()
                
                with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                    data_export = group[cols_presenti]
                    # startrow=6 inserisce la tabella a partire dalla riga 7 di Excel, lasciando lo spazio esatto sopra
                    data_export.to_excel(writer, index=False, sheet_name='Prospetto_DDT', startrow=6)
                    
                    info = anagrafiche.get(str(cliente).upper(), "")
                    formatta_excel_valsecchi(writer, 'Prospetto_DDT', is_casati=True, cliente_nome=str(cliente), info_anagrafica=info)
                
                zip_file.writestr(f"Clienti/{safe_cliente}_{mese_str}.xlsx", buf.getvalue())

            # --- 2. GENERAZIONE EXCEL AUTISTI CON FILTRO ALLO SCARICO ---
            if 'AUTISTA ALLO SCARICO' in df.columns:
                for autista, group in df.groupby('AUTISTA ALLO SCARICO'):
                    safe_autista = re.sub(r'[\\/*?:"<>|]', "", str(autista)).strip()
                    buf = io.BytesIO()
                    
                    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                        group.to_excel(writer, index=False, sheet_name='Prospetto_Autista')
                        formatta_excel_valsecchi(writer, 'Prospetto_Autista', is_casati=False)
                        
                    zip_file.writestr(f"Autisti/Scarico_{safe_autista}_{mese_str}.xlsx", buf.getvalue())
            else:
                st.warning("⚠️ Colonna 'AUTISTA ALLO SCARICO' non trovata nel file. Sezione autisti saltata.")

        st.success("✅ Documenti elaborati secondo i criteri Valsecchi Trasporti!")
        st.download_button("📥 SCARICA ARCHIVIO COMPLETO", zip_buffer.getvalue(), f"Trasporti_Valsecchi_{mese_str}.zip", "application/zip")

    except Exception as e:
        st.error(f"Errore durante l'esecuzione del codice: {e}")
