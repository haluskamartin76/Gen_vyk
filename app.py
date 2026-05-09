# --- AKTUALIZOVANÁ ČASŤ V KÓDE PRE GENEROVANIE EXCELU ---

    # Definícia farieb optimalizovaných pre ČB tlač
    # Sobota: Veľmi svetlá (na papieri takmer biela/jemne sivá)
    fill_sat = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid") 
    # Nedeľa: Stredne sivá (jasne viditeľný rozdiel)
    fill_sun = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid") 
    # Sviatok: Tmavosivá (vynikne najviac)
    fill_hol = PatternFill(start_color="595959", end_color="595959", fill_type="solid") 
    
    # Fonty
    f_norm = Font(name='Arial', size=9)
    f_white = Font(name='Arial', size=9, color="FFFFFF", bold=True) # Biele písmo pre sviatok

    # ... v cykle generovania dní ...

    for idx, val in enumerate(line, 1):
        cell = ws.cell(11+d-1, idx, val)
        cell.border = border
        cell.alignment = align_c
        cell.font = f_norm # Predvolený font
        
        if idx == 1: # Farbíme len bunku s dátumom
            if is_sv:
                cell.fill = fill_hol
                cell.font = f_white # Pri sviatku prepneme na biele písmo pre max. kontrast
            elif dow == 5: # Sobota
                cell.fill = fill_sat
            elif dow == 6: # Nedeľa
                cell.fill = fill_sun
