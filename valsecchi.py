import streamlit as st
import pandas as pd
import io
import zipfile
import re
import os
import math
import textwrap
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# --- CONFIGURAZIONE DI SISTEMA ---
st.set_page_config(page_title="Trasporti App v8.6", layout="centered")
st.caption("Versione Sistema: 8.6 (Fix Altezza Riga Fattura e Intestazioni)")

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
    
    font_bold_standard = Font(name='Calibri', size=11, bold=True, color="000000")
    font_bold_calibri12 = Font(name='Calibri', size=12, bold=True, color="000000")
    
    align_left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right_wrap = Alignment(horizontal="right", vertical="center", wrap_text=True)
    
    fill_azzurrino = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
    )

    start_row_header = 7 if is_casati else 1

    worksheet.oddFooter.center.text = "Pagina &P"
    
    worksheet.page_setup.orientation = 'landscape' 
    worksheet.page_setup.paperSize = 9  
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = False

    if is_casati:
        worksheet.print_title_rows = '1:7' 
        worksheet.row_dimensions[1].height = 80
        
        logo_path = None
        for filename in ["logo.png", "logo.jpg", "logo.jpeg", "LOGO.jpg", "LOGO.PNG"]:
            if os.path.exists(filename):
                logo_path = filename
                break
        
        if logo_path:
            try:
                from openpyxl.drawing.image import Image as OpenpyxlImage
                img = OpenpyxlImage(logo_path)
                img.width = 147  
                img.height = 91  
                worksheet.add_image(img, 'I1')
            except:
                pass
        
        worksheet.row_dimensions[2].height = 65 
        cell_azienda = worksheet['A2']
        cell_azienda.value = "VALSECCHI TRASPORTI S.R.L.\nVIA GIUSEPPE PARINI N. 8\nCESANA BRIANZA (LC)\nP.IVA 01932020132"
        cell_azienda.font = font_bold_calibri12
        cell_azienda.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        
        testo_spettabile = info_anagrafica if info_anagrafica else cliente_nome
        if info_anagrafica:
            parti_indirizzo = [p.strip() for p in re.split(r'\s{2,}', str(info_anagrafica)) if p.strip()]
            if parti_indirizzo:
                parti_indirizzo[0] = f"SPETT.LE   {parti_indirizzo[0]}"
                testo_finito = "\n".join(parti_indirizzo)
            else:
                testo_finito = f"SPETT.LE   {info_anagrafica}"
        else:
            testo_finito = f"SPETT.LE   {cliente_nome}"

        cell_spett = worksheet['A4']
        cell_spett.value = testo_finito
        cell_spett.font = font_bold_calibri12
        cell_spett.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        worksheet.row_dimensions[4].height = 65 
        
        cell_fatt = worksheet['A5']
        cell_fatt.value = "RIF. NS FATT. N. ........ DEL  ..../..../202.."
        cell_fatt.font = font_bold_standard
        cell_fatt.alignment = align_left_wrap
        # L'altezza della riga 5 viene ora calcolata dinamicamente nel blocco sottostante
        worksheet.row_dimensions[5].height = 18

    # Calcolo larghezze colonne con Limite a 25 caratteri
    for col_idx in range(1, worksheet.max_column + 1):
        col_letter = get_column_letter(col_idx)
        header_val = str(worksheet.cell(row=start_row_header, column=col_idx).value).upper()
        
        max_len = 0
        for row_idx in range(start_row_header, worksheet.max_row + 1):
            val = worksheet.cell(row=row_idx, column=col_idx).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        
        max_allowed_width = 25 
        if col_letter == 'A' and is_casati:
            worksheet.column_dimensions[col_letter].width = 30.9
        else:
            if max_len > max_allowed_width and any(x in header_val for x in ['DITTA', 'LUOGO', 'DESTINAZIONE', 'VARIE']):
                worksheet.column_dimensions[col_letter].width = max_allowed_width
            else:
                worksheet.column_dimensions[col_letter].width = max(max_len + 4, 16)

    # Formattazione allineamenti base
    for row in worksheet.iter_rows(min_row=start_row_header):
        for cell in row:
            cell.font = font_bold_standard
            cell.border = thin_border
            if cell.column <= 4:
                cell.alignment = align_left_wrap
            else:
                cell.alignment = align_right_wrap

    # --- FIX 8.6: MOTORE DI SIMULAZIONE ALTEZZA (Esteso da riga 4 in giù) ---
    start_dynamic_row = 4 if is_casati else 1
    for row_idx in range(start_dynamic_row, worksheet.max_row + 1):
        max_lines_in_row = 0
        has_text = False
        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            if cell.value is not None:
                has_text = True
                val_str = str(cell.value)
                col_letter = get_column_letter(col_idx)
                
                w = worksheet.column_dimensions[col_letter].width or 16
                safe_width = max(5, int(w) - 2)
                
                lines_for_this_cell = 0
                for paragraph in val_str.split('\n'):
                    wrapped = textwrap.wrap(paragraph, width=safe_width)
                    lines_for_this_cell += len(wrapped) if wrapped else 1
                
                if lines_for_this_cell > max_lines_in_row:
                    max_lines_in_row = lines_for_this_cell
        
        if has_text:
            if row_idx == 4 and is_casati:
                # Protegge la riga 4 Spettabile mantenendo almeno 65 di altezza
                worksheet.row_dimensions[row_idx].height = max(65, max_lines_in_row * 16.5)
            elif row_idx < start_row_header:
                # Applica il calcolo alla riga 5 della fattura
                worksheet.row_dimensions[row_idx].height = max(18, max_lines_in_row * 16.5)
            else:
                # Applica il calcolo a tutta la tabella sottostante
                worksheet.row_dimensions[row_idx].height = max(24.9, max_lines_in_row * 16.5)
    # ---------------------------------------------------------

    # Sfondo azzurrino
    for cell in worksheet[start_row_header]:
        cell.fill = fill_azzurrino

    # Formattazione Decimali e Date
    for col_idx in range(1, worksheet.max_column + 1):
        col_letter = get_column_letter(col_idx)
        header_val = str(worksheet.cell(row=start_row_header, column=col_idx).value).upper()
        
        for row_idx in range(start_row_header + 1, worksheet.max_row + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            if cell.value is not None:
                if any(x in header_val for x in ['TARIFFA', 'TOTALE']):
                    cell.number_format = '#,##0.00'
                if any(x in header_val for x in ['DATA', 'DDT']) and not isinstance(cell.value, str):
