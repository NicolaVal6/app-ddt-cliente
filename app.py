import streamlit as st
import pandas as pd
import io
import zipfile
import re
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Generatore Excel DDT", layout="centered")

# --- 1. SICUREZZA ---
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


# --- 2. FUNZIONI DI SUPPORTO E FORMATTAZIONE ---

def sanitize_filename(name):
    if pd.isna(name): return "Sconosciuto"
    return re.sub(r'[\\/*?:"<>|]', "", str(name)).strip()

def formatta_excel_casati(writer, sheet_name, cliente_nome, info_anagrafica=""):
    """Formatta l'Excel per i Clienti con lo standard CASATI (Intestazione alta)"""
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    testo_spettabile = info_anagrafica if info_anagrafica else cliente_nome
    worksheet['A2'] = f"SPETT.LE   {testo_spettabile}"
    worksheet['A2'].font = Font(bold=True, size=11)
    
    worksheet['A3'] = "RIF. NS FATT. N. ........ DEL  ..../..../202.."
    worksheet['A3'].font = Font(bold=True)

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    
    for cell in worksheet[5]:
        cell.font = header_font
        cell.fill = header_fill

    date_cols_to_format = ['DATA DDT', 'DATA DI CARICO']

    for col_idx, col in enumerate(worksheet.columns, 1):
        max_length = 0
        column_letter = get_column_letter(col_idx)
        header_value = worksheet.cell(row=5, column=col_idx).value

        for cell in col:
            if header_value in date_cols_to_format and cell.row > 5:
                cell.number_format = 'DD/MM/YYYY'
            try:
                if cell.value and cell.row > 4:
                    val_str = str(cell.value)
                    if len(val_str) > max_length:
                        max_length = len(val_str)
            except:
                pass
        
        adjusted_width = (max_length + 2) if max_length > 8 else 12
        worksheet.column_dimensions[column_letter].width = adjusted_width

def formatta_excel_standard(writer, sheet_name):
    """Formatta l'Excel classico (usato per i file interni degli Autisti)"""
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    date_cols_to_format = ['DATA DI CARICO', 'DATA DDT']

    for col_idx, col in enumerate(worksheet.columns, 1):
        max_length = 0
        column_letter = get_column_letter(col_idx)
        header_value = col[0].value
        col[0].font = header_font
        col[0].fill = header_fill
        for cell in col:
            if header_value in date_cols_to_format and cell.row > 1:
                cell.number_format = 'DD/MM/YYYY'
            try:
                if cell.value:
                    if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
            except:
                pass
        worksheet.column_dimensions[column_letter].width = (max_length + 2) if max_length > 10 else 15


# --- 3. INTERFACCIA APP ---

try:
    st.image("logo.png", width=250)
except Exception:
    pass

st.title("🚛 Generatore File DDT - Advanced")

col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("1. File Mensile Trasporti (Dati)", type=["xlsx", "csv"])
with col2:
    uploaded_anagrafica = st.file_uploader("2. File Anagrafiche Clienti (Opzionale)", type=["xlsx", "csv"])

if uploaded_file is not None:
    if st.button("Elabora Dati", type="primary"):
        try:
            anagrafiche_dict = {}
            if uploaded_anagrafica is not None:
                df_ana = pd.read_excel(uploaded_anagrafica) if uploaded_anagrafica.name.endswith('.xlsx') else pd.read_csv(uploaded_anagrafica)
                df_ana.columns = df_ana.columns.str.strip().str.upper()
                if 'CLIENTE' in df_ana.columns and 'INTESTAZIONE COMPLETA' in df_ana.columns:
                    anagrafiche_dict = dict(zip(df_ana['CLIENTE'].astype(str).str.upper().str.strip(), df_ana['INTESTAZIONE COMPLETA']))

            # Lettura file principale
            df = pd.read_csv(uploaded_file) if uploaded_file.name.lower().endswith('.csv') else pd.read_excel(uploaded_file)
            df.columns = df.columns.str.strip().str.upper()
            
            # PULIZIA: Rimuove le colonne vuote o fantasma (es. Unnamed: 3 generato dalle doppie virgole del gestionale)
            df = df.loc[:, ~df.columns.str.contains('^UNNAMED', na=False)]

            # Formattazione Date Universale
            date_columns = ['DATA DI CARICO', 'DATA DDT']
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
            
            # Estrazione Mese per rinomina ZIP
            if 'DATA DDT' in df.columns:
                first_valid = df['DATA DDT'].dropna()
                mese_str = first_valid.iloc[0].strftime("%Y-%m") if not first_valid.empty else "MESE_IGNOTO"
            else:
                mese_str = "MESE_IGNOTO"
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                
                # --- LOGICA CLIENTI (STILE CASATI) ---
                # Mappatura Universale (vecchio file o nuovo gestionale) alle colonne Target definitive
                mappa_rinomina = {
                    'DITTA DI CARICO': 'DITTA PRESA',
                    'LUOGO DI CARICO': 'LUOGO PRESA',
                    'DESTINAZIONE': 'DESTINAZIONE FINALE',
                    'VARIE': 'VARIE',
                    'TIPO MERCE': 'VARIE',
                    'PESO': "QUANTITA'",
                    'QUANTITA': "QUANTITA'",
                    'N. DDT': 'N. DDT',
                    'DDT': 'N. DDT'
                }
                
                # Rinomina preventiva sul dataframe copiato per i clienti
                df_clienti = df.rename(columns=mappa_rinomina)
                
                # Colonne che il cliente vuole nel file finale (se esistono)
                cols_target_finali = ['DITTA PRESA', 'LUOGO PRESA', 'DESTINAZIONE FINALE', 'VARIE', "QUANTITA'", 'TARIFFA', 'TOTALE', 'N. DDT', 'DATA DDT']
                cols_effettive_esportazione = [col for col in cols_target_finali if col in df_clienti.columns]
                
                if 'CLIENTE' in df.columns:
                    for cliente, group in df.groupby('CLIENTE'):
                        safe_cliente = sanitize_filename(cliente)
                        filename = f"Clienti/{safe_cliente}_{mese_str}.xlsx"
                        
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            # Prende il gruppo mappato, isola le colonne target e esporta
                            group_mapped = df_clienti.loc[group.index]
                            group_to_export = group_mapped[cols_effettive_esportazione]
                            
                            group_to_export.to_excel(writer, index=False, sheet_name='Dati_DDT', startrow=4)
                            
                            info_ana = anagrafiche_dict.get(str(cliente).upper().strip(), "")
                            formatta_excel_casati(writer, 'Dati_DDT', cliente, info_ana)
                        
                        zip_file.writestr(filename, excel_buffer.getvalue())
                
                # --- LOGICA AUTISTI (STILE STANDARD) ---
                # Per gli autisti esportiamo tutte le colonne (senza le vuote/unnamed, già rimosse in alto)
                if 'AUTISTA AL CARICO' in df.columns:
                    for autista, group in df.groupby('AUTISTA AL CARICO'):
                        safe_autista = sanitize_filename(autista)
                        filename = f"Autisti/Autista_{safe_autista}_{mese_str}.xlsx"
                        
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            group.to_excel(writer, index=False, sheet_name='Dati_DDT')
                            formatta_excel_standard(writer, 'Dati_DDT')
                        
                        zip_file.writestr(filename, excel_buffer.getvalue())

            st.success("✅ Elaborazione (Input Gestionale) completata!")
            st.download_button(
                label="📥 Scarica Archivio ZIP",
                data=zip_buffer.getvalue(),
                file_name=f"DDT_Elaborati_{mese_str}.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"Errore di elaborazione: {e}")
