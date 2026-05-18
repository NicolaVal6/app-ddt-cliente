import streamlit as st
import pandas as pd
import io
import zipfile
import re
import os
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# --- CONFIGURAZIONE DI SISTEMA ---
st.set_page_config(page_title="Trasporti App v7.3", layout="centered")
st.caption("Versione Sistema: 7.3 (Fix Scrittura Excel + Visualizzazione Pagine A4 + Righe 24.9)")

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
    
    # Stili Rigorosi (Tutto l'allegato in Grassetto)
    font_bold_standard = Font(name='Calibri', size=11, bold=True, color="000000")
    font_bold_calibri12 = Font(name='Calibri', size=12, bold=True, color="000000")
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    
    # Sfondo Azzurrino Intestazione
    fill_azzurrino = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
    )

    start_row_header = 7 if is_casati else 1
    
    # Piè di pagina dinamico per numero pagina
    worksheet.oddFooter.center.text = "Pagina &P"

    # 1. FORZATURA DIVISIONE IN PAGINE VISIBILE ALL'APERTURA
    worksheet.page_setup.orientation = 'portrait'
    worksheet.page_setup.paperSize = 9  # Codice Excel per foglio A4
    
    # Questo comando costringe Excel a mostrare le interruzioni di pagina e la visualizzazione "Layout di Pagina"
    try:
        worksheet.views.sheetView[0].view = 'pageLayout'
        worksheet.sheet_view.showPageBreaks = True
    except:
        pass

    if is_casati:
        # Ripetizione automatica delle intestazioni (righe da 1 a 7) su ogni foglio stampato
        worksheet.print_title_rows = '1:7'
        
        # Configurazione geometrica casella Logo A1
        worksheet.row_dimensions[1].height = 45
        
        if os.path.exists("logo.png"):
            try:
                from openpyxl.drawing.image import Image as OpenpyxlImage
                img = OpenpyxlImage("logo.png")
                img.width = 110
                img.height = 40
                worksheet.add_image(img, 'A1')
            except:
                pass
        
        # 2. INTESTAZIONE VALSECCHI: Spostata in A2, a capo interno, bloccata a larghezza 30.9
        worksheet.row_dimensions[2].height = 65 
        cell_azienda = worksheet['A2']
        cell_azienda.value = "VALSECCHI TRASPORTI S.R.L.\nVIA GIUSEPPE PARINI N. 8\nCESANA BRIANZA (LC)\nP.IVA 01932020132"
        cell_azienda.font = font_bold_calibri12
        cell_azienda.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        
        # 3. Spett.le Cliente a SINISTRA (in A4)
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

    # 5. MODIFICA 1: Formattazione Righe Tabella con Altezza Esatta a 24.9
    for row in worksheet.iter_rows(min_row=start_row_header):
        worksheet.row_dimensions[row[0].row].height = 24.9  # Imposta l'altezza riga richiesta
        for cell in row:
            cell.font = font_bold_standard
            cell.border = thin_border
            cell.alignment = align_center

    # Applica sfondo azzurrino solo alla riga delle intestazioni (Riga 7)
    for cell in worksheet[start_row_header]:
        cell.fill = fill_azzurrino

    # Formattazione Numeri Decimali e Date
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
        
        # MODIFICA 2: Colonna A bloccata rigidamente a 30.9 per contenere l'intestazione Valsecchi
        if col_letter == 'A' and is_casati:
            worksheet.column_dimensions[col_letter].width = 30.9
        else:
            max_len = max(len(str(cell.value or '')) for cell in worksheet[col_letter] if cell.row >= start_row_header)
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 16)

    # 6. CALCOLO AUTOMATICO DEL TOTALE DELLE SOMME SOTTO "DITTA PRESA"
    if is_casati:
        totale_col_idx = None
        for col_idx in range(1, worksheet.max_column + 1):
            if "TOTALE" in str(worksheet.cell(row=start_row_header, column=col_idx).value).upper():
                totale_col_idx = col_idx
                break
        
        if totale_col_idx:
            last_data_row = worksheet.max_row
            total_row_idx = last_data_row + 1
            
            # Altezza fissa anche per la riga del totale
            worksheet.row_dimensions[total_row_idx].height = 24.9
            
            # Scritta TOTALE sotto la colonna DITTA PRESA (Colonna 1 / A)
            cell_label = worksheet.cell(row=total_row_idx, column=1)
            cell_label.value = "TOTALE"
            
            # Formula SUM dinamica
            col_letter = get_column_letter(totale_col_idx)
            cell_sum = worksheet.cell(row=total_row_idx, column=totale_col_idx)
            cell_sum.value = f"=SUM({col_letter}8:{col_letter}{last_data_row})"
            cell_sum.number_format = '#,##0.00'
            
            # Disegna la griglia finale della riga totali
            for col_idx in range(1, worksheet.max_column + 1):
                c = worksheet.cell(row=total_row_idx, column=col_idx)
                c.font = font_bold_standard
                c.border = thin_border
                c.alignment = align_center


# --- INTERFACCIA WEB ---
try: st.image("logo.png", width=250)
except: pass

st.title("🚛 Sistema Gestionale Valsecchi Trasporti")

uploaded_file = st.file_uploader("1. Carica Estrazione Gestionale (.xlsx o .csv)", type=["xlsx", "csv"])
uploaded_anagrafica = st.file_uploader("2. Carica Anagrafiche Clienti (Opzionale)", type=["xlsx", "csv"])

if uploaded_file and st.button("🚀 GENERA DOCUMENTI FISCALI", type="primary"):
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

        # ORDINE SEQUENZIALE COLONNE RICHIESTO
        cols_ordinate = ['DITTA PRESA', 'LUOGO PRESA', 'DESTINAZIONE FINALE', 'VARIE', "QUANTITA'", 'TARIFFA', 'TOTALE', 'N. DDT', 'DATA DDT']
        cols_presenti = [c for c in cols_ordinate if c in df_mapped.columns]

        first_valid = df_mapped['DATA DDT'].dropna() if 'DATA DDT' in df_mapped.columns else pd.Series()
        mese_str = first_valid.iloc[0].strftime("%Y-%m") if not first_valid.empty else "MESE_IGNOTO"

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            
            # --- 1. FILE CLIENTI ---
            for cliente, group in df_mapped.groupby('CLIENTE'):
                safe_cliente = re.sub(r'[\\/*?:"<>|]', "", str(cliente)).strip()
                buf = io.BytesIO()
                
                # FIX APPLICATO QUI: engine='openpyxl' (non io_engine)
                with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                    data_export = group[cols_presenti]
                    data_export.to_excel(writer, index=False, sheet_name='Prospetto_DDT', startrow=6)
                    
                    info = anagrafiche.get(str(cliente).upper(), "")
                    formatta_excel_valsecchi(writer, 'Prospetto_DDT', is_casati=True, cliente_nome=str(cliente), info_anagrafica=info)
                
                zip_file.writestr(f"Clienti/{safe_cliente}_{mese_str}.xlsx", buf.getvalue())

            # --- 2. FILE AUTISTI (FILTRATO AD AUTISTA ALLO SCARICO) ---
            if 'AUTISTA ALLO SCARICO' in df.columns:
                for autista, group in df.groupby('AUTISTA ALLO SCARICO'):
                    safe_autista = re.sub(r'[\\/*?:"<>|]', "", str(autista)).strip()
                    buf = io.BytesIO()
                    
                    # FIX APPLICATO QUI: engine='openpyxl' (non io_engine)
                    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                        group.to_excel(writer, index=False, sheet_name='Prospetto_Autista')
                        formatta_excel_valsecchi(writer, 'Prospetto_Autista', is_casati=False)
                        
                    zip_file.writestr(f"Autisti/Scarico_{safe_autista}_{mese_str}.xlsx", buf.getvalue())
            else:
                st.warning("⚠️ Colonna 'AUTISTA ALLO SCARICO' non trovata nel file. Sezione autisti saltata.")

        st.success("✅ Documenti aggiornati alla v7.3 e pronti!")
        st.download_button("📥 SCARICA ARCHIVIO COMPLETO", zip_buffer.getvalue(), f"Trasporti_Valsecchi_{mese_str}.zip", "application/zip")

    except Exception as e:
        st.error(f"Errore: {e}")
