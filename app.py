import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from datetime import date
import calendar
import io

# --- KONFIGURÁCIA ---
MESIACE_SK = ["Január", "Február", "Marec", "Apríl", "Máj", "Jún", "Júl", "August", "September", "Október", "November", "December"]
DNI_SK = ["Pondelok", "Utorok", "Streda", "Štvrtok", "Piatok", "Sobota", "Nedeľa"]
SVIATKY_2026 = {1: [1, 6], 4: [3, 6], 5: [1, 8], 7: [5], 8: [29], 9: [1, 15], 11: [1, 17], 12: [24, 25, 26]}

st.set_page_config(page_title="Generátor Výkazu NR SR", layout="wide")

# Funkcia na resetovanie polí
def reset_cycle():
    for key in list(st.session_state.keys()):
        if key.startswith("d_") or key.startswith("n_") or key.startswith("a_"):
            del st.session_state[key]

st.title("📄 Generátor pracovného výkazu")

# --- BOČNÝ PANEL (NASTAVENIA) ---
with st.sidebar:
    st.header("Nastavenia")
    meno = st.text_input("Meno a Priezvisko:", placeholder="Martin Haluska")
    mesiac_meno = st.selectbox("Mesiac:", MESIACE_SK, index=4) # Predvolený Máj (index 4)
    mesiac_idx = MESIACE_SK.index(mesiac_meno) + 1
    rok = st.number_input("Rok:", value=2026)
    zmena_skupina = st.selectbox("Základný cyklus:", ["Zmena 1", "Zmena 2", "Zmena 3", "Zmena 4"], index=3) # Predvolená Zmena 4
    fond = st.number_input("Fond hodín na mesiac:", value=147.0, step=0.5)
    utvar = st.text_input("Organizačný útvar:", value="Odbor obrany, bezpečnosti a ochrany")
    
    st.divider()
    if st.button("🔄 Načítať / Resetovať cyklus", use_container_width=True, on_click=reset_cycle):
        st.toast("Cyklus bol úspešne načítaný pre vybraný mesiac.")

# --- LOGIKA VÝPOČTU ---
cykly = {"Zmena 1": "DNVDNVVV", "Zmena 2": "VVDNVDNV", "Zmena 3": "VDNVVVDN", "Zmena 4": "NVVVDNVD"}
vzor, ref = cykly[zmena_skupina], date(2026, 3, 1)
_, dni_count = calendar.monthrange(rok, mesiac_idx)

st.subheader(f"Upresnenie dní pre {mesiac_meno} ({zmena_skupina})")

# --- INTERAKTÍVNA TABUĽKA ---
vysledne_dni = []
# Hlavička tabuľky
h1, h2, h3, h4, h5 = st.columns([0.5, 0.5, 0.5, 1, 1.5])
h1.write("**Deň**")
h2.write("**D**")
h3.write("**N**")
h4.write("**Absencia**")
h5.write("**Dátum**")

for d in range(1, dni_count + 1):
    dt = date(rok, mesiac_idx, d)
    pos = (dt - ref).days % 8
    
    # Predvolené hodnoty podľa cyklu
    def_d = (vzor[pos] == "D")
    def_n = (vzor[pos] == "N")
    
    col1, col2, col3, col4, col5 = st.columns([0.5, 0.5, 0.5, 1, 1.5])
    
    with col1: st.write(f"{d}.")
    with col2: d_val = st.checkbox(" ", value=def_d, key=f"d_{d}")
    with col3: n_val = st.checkbox(" ", value=def_n, key=f"n_{d}")
    with col4: abs_val = st.text_input("Abs", key=f"a_{d}", label_visibility="collapsed", placeholder="D/KZ/PN/L")
    with col5: 
        label = f"{dt.strftime('%d.%m.')} ({DNI_SK[dt.weekday()]})"
        if dt.weekday() >= 5: st.write(f":orange[{label}]")
        else: st.write(label)
    
    vysledne_dni.append({
        "den": d, "is_d": d_val, "is_n": n_val, 
        "abs": abs_val.upper().strip(), "weekday": dt.weekday()
    })

# --- GENEROVANIE EXCELU ---
st.divider()
if st.button("💾 VYGENEROVAŤ EXCEL VÝKAZ", type="primary", use_container_width=True):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    
    f_header = Font(name='Arial', size=11, bold=True)
    f_bold = Font(name='Arial', size=9, bold=True)
    f_norm = Font(name='Arial', size=9)
    align_c = Alignment(horizontal='center', vertical='center', wrap_text=True)
    side = Side(style='thin')
    border = Border(left=side, right=side, top=side, bottom=side)
    fill_sat, fill_sun, fill_hol = PatternFill("solid", "FFFF00"), PatternFill("solid", "00FF00"), PatternFill("solid", "00CCFF")

    # Hlavička dokumentu
    ws.merge_cells('A1:J1'); ws['A1'] = "Kancelária Národnej rady Slovenskej republiky"; ws['A1'].font = f_header; ws['A1'].alignment = align_c
    ws.merge_cells('A2:J2'); ws['A2'] = f"Pracovný výkaz za mesiac {mesiac_meno} {rok}"; ws['A2'].font = f_header; ws['A2'].alignment = align_c
    ws['A4'], ws['C4'], ws['H4'], ws['I4'] = "Organizačný útvar:", utvar, "Fond:", fond
    ws['A6'], ws['C6'] = "Zamestnanec:", meno
    for c in ['C4', 'I4', 'C6']: ws[c].font = f_bold

    # Hlavička tabuľky
    ws.merge_cells('A8:A10'); ws['A8'] = "Dátum"
    ws.merge_cells('B8:B10'); ws['B8'] = "Pracovná doba"
    ws.merge_cells('C8:C10'); ws['C8'] = "Odprac.\nhod.\nspolu"
    ws.merge_cells('D8:H8'); ws['D8'] = "z toho"
    ws.merge_cells('I8:J8'); ws['I8'] = "Pracovná pohotovosť"
    ws.merge_cells('D9:D10'); ws['D9'] = "v so-ne"; ws.merge_cells('E9:E10'); ws['E9'] = "v noci"
    ws.merge_cells('F9:G9'); ws['F9'] = "nadčas"; ws.merge_cells('H9:H10'); ws['H9'] = "sviatok"
    ws['I9'], ws['J9'] = "na\npracovisk.", "mimo\npracovisk."
    ws['F10'], ws['G10'] = "Po – Pia", "So – Ne"

    for r_idx in range(8, 11):
        for c_idx in range(1, 11):
            cell = ws.cell(r_idx, c_idx); cell.font = f_bold; cell.alignment = align_c; cell.border = border

    sviatky = SVIATKY_2026.get(mesiac_idx, [])
    for row_data in vysledne_dni:
        d, dow = row_data["den"], row_data["weekday"]
        is_sviatok = d in sviatky
        p_doba, h_sp, h_sn, h_no, h_sv = "", "", "", "", ""
        
        if row_data["abs"] == "D": 
            p_doba, h_sp = "Dovolenka", 11.5
        elif row_data["abs"] == "KZ": 
            p_doba, h_sp = "KZ", 11.5
        elif row_data["abs"] in ["PN", "L"]: 
            p_doba = row_data["abs"]
        elif row_data["is_d"]:
            p_doba, h_sp = "06:00 - 18:00", 11.5
            if dow >= 5: h_sn = 11.5
            if is_sviatok: h_sv = 11.5
        elif row_data["is_n"]:
            p_doba, h_sp, h_no = "18:00 - 06:00", 11.5, 7.5
            if dow == 4: h_sn = 6.0   # Pi-So
            elif dow == 5: h_sn = 11.5 # So-Ne
            elif dow == 6: h_sn = 5.5  # Ne-Po
            # Logika: Ak zmena začína vo sviatok, celá sa ráta do sviatku
            if is_sviatok: h_sv = 11.5

        line = [d, p_doba, h_sp, h_sn, h_no, "", "", h_sv, "", ""]
        for idx, val in enumerate(line, 1):
            cell = ws.cell(11+d-1, idx, val); cell.border = border; cell.alignment = align_c; cell.font = f_norm
            if idx == 1:
                if is_sviatok: cell.fill = fill_hol
                elif dow == 5: cell.fill = fill_sat
                elif dow == 6: cell.fill = fill_sun

    # Spodný riadok súčtov
    s_row = 11 + dni_count
    ws.merge_cells(f'A{s_row}:B{s_row}')
    ws[f'A{s_row}'] = "Odpracované hodiny s p o l u:"
    ws[f'A{s_row}'].alignment = Alignment(horizontal='left'); ws[f'A{s_row}'].font = f_bold
    for col_idx in range(1, 11): ws.cell(s_row, col_idx).border = border
    
    for col_l in ["C", "D", "E", "H", "I", "J"]:
        ws[f'{col_l}{s_row}'] = f"=SUM({col_l}11:{col_l}{s_row-1})"
        ws[f'{col_l}{s_row}'].font = f_bold; ws[f'{col_l}{s_row}'].alignment = align_c

    ws.merge_cells(f'F{s_row}:G{s_row}')
    ws[f'F{s_row}'] = f"=C{s_row}-I4"; ws[f'F{s_row}'].font = f_bold; ws[f'F{s_row}'].alignment = align_c
    
    # Podpisy
    sig_r = s_row + 3
    ws.merge_cells(f'B{sig_r}:D{sig_r}'); ws[f'B{sig_r}'] = "____________________________"; ws[f'B{sig_r}'].alignment = align_c
    ws.merge_cells(f'G{sig_r}:I{sig_r}'); ws[f'G{sig_r}'] = "____________________________"; ws[f'G{sig_r}'].alignment = align_c
    ws.merge_cells(f'B{sig_r+1}:D{sig_r+1}'); ws[f'B{sig_r+1}'] = "podpis zamestnanca"; ws[f'B{sig_r+1}'].alignment = align_c
    ws.merge_cells(f'G{sig_r+1}:I{sig_r+1}'); ws[f'G{sig_r+1}'] = "podpis vedúceho"; ws[f'G{sig_r+1}'].alignment = align_c

    widths = [6, 18, 10, 10, 10, 10, 10, 10, 10, 10]
    for i, w in enumerate(widths, 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    wb.save(output)
    st.download_button(label="📥 STIAHNUŤ VYGENEROVAŤ EXCEL", data=output.getvalue(), file_name=f"Vykaz_{meno.replace(' ','_')}_{mesiac_meno}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
