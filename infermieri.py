import pytesseract
import streamlit as st
import numpy as np
from PIL import Image
import re
import pandas as pd
import pdfplumber
import calendar
from dateutil.easter import easter
from datetime import datetime
from io import BytesIO
import pyexcel

# Configurazione Tesseract per macOS
pytesseract.pytesseract.tesseract_cmd = '/usr/local/bin/tesseract'

# Lista fissa dei nominativi
FIXED_NAMES = [
    "TATIANA ARTENIE", "GIANLUCA MATARRESE", "SIMONA LUPU", 
    "CHIAROTTI EMMA", "RUGGIANO MASSIMO", "LIVIA BURRAI", 
    "ANDREA PALLI", "TIZNADO MARLY", "SPADA MATTEO", "CUTAIA MARCO"
]

ORE_MAP = {
    'M': 7, 'P': 7, 'N': 10, 'MP': 14,
    'PN': 16, 'REC': -6, 'F': 6, 'S': 0
}

SHIFT_COLORS = {
    'M': '#FFCCCB', 'P': '#ADD8E6', 'N': '#90EE90',
    'R': '#D3D3D3', 'S': '#FFA07A', 'F': '#FFD700',
    'PN': '#FF69B4', 'MP': '#9370DB', 'REC': '#D3D3D3'
}

MONTH_COLORS = {
    'Gennaio': '#FF6F61', 'Febbraio': '#6B5B95',
    'Marzo': '#88B04B', 'Aprile': '#F7CAC9',
    'Maggio': '#92A8D1', 'Giugno': '#955251',
    'Luglio': '#B565A7', 'Agosto': '#009B77',
    'Settembre': '#DD4124', 'Ottobre': '#D65076',
    'Novembre': '#45B8AC', 'Dicembre': '#EFC050'
}

def normalize_name(name):
    name = re.sub(r'\s+', ' ', str(name).upper().strip())
    name = re.sub(r'[^A-ZÀÈÉÌÒÙ\s]', '', name)
    return name

def is_valid_name(text):
    return re.match(r'^[A-ZÀÈÉÌÒÙ]{2,}\s+[A-ZÀÈÉÌÒÙ]{2,}(\s+[A-ZÀÈÉÌÒÙ]{2,})*$', text)

def get_italian_holidays(year):
    holidays = [
        {'month': 1, 'day': 1, 'name': 'Capodanno'},
        {'month': 1, 'day': 6, 'name': 'Epifania'},
        {'month': 4, 'day': 25, 'name': 'Liberazione'},
        {'month': 5, 'day': 1, 'name': 'Lavoro'},
        {'month': 6, 'day': 2, 'name': 'Repubblica'},
        {'month': 8, 'day': 15, 'name': 'Ferragosto'},
        {'month': 11, 'day': 1, 'name': 'Ognissanti'},
        {'month': 12, 'day': 8, 'name': 'Immacolata'},
        {'month': 12, 'day': 25, 'name': 'Natale'},
        {'month': 12, 'day': 26, 'name': 'S.Stefano'}
    ]
    
    easter_date = easter(year)
    pasquetta = easter_date + pd.DateOffset(days=1)
    holidays.append({'month': pasquetta.month, 'day': pasquetta.day, 'name': 'Pasquetta'})
    
    return holidays

def extract_from_excel(excel_file):
    """Gestisce sia XLS che XLSX con pyexcel e openpyxl"""
    try:
        content = excel_file.read()
        excel_file.seek(0)
        
        # Controlla se è un file XLS (formato binario)
        if content.startswith(b'\xD0\xCF\x11\xE0'):
            sheet = pyexcel.get_sheet(file_type='xls', file_stream=BytesIO(content))
            df = pd.DataFrame(sheet.array)
        else:
            df = pd.read_excel(excel_file, header=None, engine='openpyxl')
            
        people_shifts = {}
        current_name = None
        
        for idx, row in df.iterrows():
            for cell in row:
                if pd.isna(cell):
                    continue
                clean_cell = normalize_name(str(cell))
                if is_valid_name(clean_cell):
                    fixed_match = next((fn for fn in FIXED_NAMES if fn in clean_cell), None)
                    current_name = fixed_match if fixed_match else clean_cell
                elif current_name and re.match(r'^[MPNRECSF]{1,2}$', str(cell).strip().upper()):
                    if current_name not in people_shifts:
                        people_shifts[current_name] = []
                    people_shifts[current_name].append(str(cell).strip().upper())
        
        return people_shifts
        
    except Exception as e:
        st.error(f"Errore lettura Excel: {str(e)}")
        return {}

def extract_from_pdf(pdf_file):
    try:
        people_shifts = {}
        current_name = None
        
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                tables = page.extract_tables()
                
                for line in text.split('\n'):
                    clean_line = normalize_name(line)
                    if is_valid_name(clean_line):
                        fixed_match = next((fn for fn in FIXED_NAMES if fn in clean_line), None)
                        current_name = fixed_match if fixed_match else clean_line
                
                for table in tables:
                    for row in table:
                        for cell in row:
                            if not cell:
                                continue
                            clean_cell = normalize_name(str(cell))
                            if is_valid_name(clean_line):
                                fixed_match = next((fn for fn in FIXED_NAMES if fn in clean_line), None)
                                current_name = fixed_match if fixed_match else clean_line
                            elif current_name and re.match(r'^[MPNRECSF]{1,2}$', str(cell).strip().upper()):
                                if current_name not in people_shifts:
                                    people_shifts[current_name] = []
                                people_shifts[current_name].append(str(cell).strip().upper())
        
        return people_shifts
        
    except Exception as e:
        st.error(f"Errore lettura PDF: {str(e)}")
        return {}

def calculate_metrics(shifts, month, year):
    try:
        month_num = [
            'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
            'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'
        ].index(month) + 1
        
        days_in_month = calendar.monthrange(year, month_num)[1]
        
        holidays = get_italian_holidays(year)
        festivita = [h for h in holidays if h['month'] == month_num]
        festivita_count = len(festivita)
        festivita_nomi = [h['name'] for h in festivita]
        
        valid_shifts = shifts[:days_in_month]
        shift_counts = {s: valid_shifts.count(s) for s in ORE_MAP}
        ore_totali = {s: count * ORE_MAP[s] for s, count in shift_counts.items()}
        
        ore_mensili = sum(ore for s, ore in ore_totali.items() if s not in ['R', 'S', 'REC'])
        r_count = shift_counts.get('R', 0)
        target_ore = (days_in_month - (r_count + festivita_count)) * 6
        ore_mancanti = target_ore - ore_mensili
        
        cal = calendar.Calendar(firstweekday=0)
        month_weeks = cal.monthdayscalendar(year, month_num)
        shifts_per_day = (valid_shifts + [''] * days_in_month)[:days_in_month]
        
        weeks = []
        for week in month_weeks:
            week_data = []
            for day in week:
                if day == 0:
                    week_data.append((0, ''))
                else:
                    shift = shifts_per_day[day-1] if (day-1) < len(shifts_per_day) else ''
                    week_data.append((day, shift))
            weeks.append(week_data)
        
        return {
            'days_in_month': days_in_month,
            'festivita_count': festivita_count,
            'festivita_nomi': festivita_nomi,
            'shift_counts': shift_counts,
            'ore_totali': ore_totali,
            'ore_mensili': ore_mensili,
            'target_ore': target_ore,
            'ore_mancanti': ore_mancanti,
            'weeks': weeks
        }
        
    except Exception as e:
        st.error(f"Errore nei calcoli: {str(e)}")
        return None

def display_month(month, year, festivita_nomi):
    month_color = MONTH_COLORS.get(month, '#000000')
    st.markdown(
        f"<h2 style='text-align: center; color: {month_color};'>"
        f"{month} {year}</h2>",
        unsafe_allow_html=True
    )
    
    if festivita_nomi:
        st.markdown(
            f"<div style='text-align: center; margin-top: -15px; color: #666666;'>"
            f"({', '.join(festivita_nomi)})</div>",
            unsafe_allow_html=True
        )

def main():
    st.title("⏰ Analizzatore Turni Professionale")
    
    col1, col2 = st.columns(2)
    with col1:
        month = st.selectbox(
            "Seleziona il mese:",
            ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
             'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre']
        )
    with col2:
        year = st.number_input(
            "Seleziona l'anno:",
            min_value=2000,
            max_value=2100,
            value=datetime.now().year
        )
    
    uploaded_file = st.file_uploader("Carica il planning turni", type=['pdf', 'xlsx', 'xls'])
    
    if uploaded_file:
        with st.spinner('Elaborazione in corso...'):
            if uploaded_file.type == "application/pdf":
                people_shifts = extract_from_pdf(uploaded_file)
            else:
                people_shifts = extract_from_excel(uploaded_file)
        
        if people_shifts:
            valid_people = {k: v for k, v in people_shifts.items() if len(v) > 0}
            
            if not valid_people:
                st.error("❌ Nessun dato rilevato nel documento!")
                return
                
            all_names = list(
                dict.fromkeys(
                    [name for name in FIXED_NAMES if name in valid_people] +
                    [name for name in people_shifts.keys() if name not in FIXED_NAMES and name in valid_people]
                )
            )
            
            st.subheader("👤 Selezione Personale")
            selected_name = st.selectbox("Seleziona un collaboratore:", options=all_names)
            
            shifts = valid_people.get(selected_name, [])
            
            if shifts:
                metrics = calculate_metrics(shifts, month, year)
                
                if metrics:
                    display_month(month, year, metrics['festivita_nomi'])
                
                st.subheader(f"📅 Turni di {selected_name}")
                
                weekdays = ['Lun', 'Mar', 'Mer', 'Gio', 'Ven', 'Sab', 'Dom']
                cols = st.columns(7)
                for i, day in enumerate(weekdays):
                    with cols[i]:
                        st.markdown(f"**{day}**")

                for week in metrics['weeks']:
                    cols = st.columns(7)
                    for i, (day_num, shift) in enumerate(week):
                        with cols[i]:
                            if day_num != 0:
                                bg_color = SHIFT_COLORS.get(shift, '#FFFFFF')
                                st.markdown(
                                    f"<div style='text-align: center; margin: 2px; padding: 8px; "
                                    f"border-radius: 5px; background-color: {bg_color}; "
                                    f"min-height: 50px; display: flex; flex-direction: column; justify-content: center;'>"
                                    f"<div style='font-size: 0.8em; color: #666;'>{day_num}</div>"
                                    f"<div style='font-size: 1.2em;'>{shift}</div>"
                                    f"</div>",
                                    unsafe_allow_html=True
                                )
                            else:
                                st.markdown("<div style='min-height:50px'></div>", unsafe_allow_html=True)
                
                if metrics:
                    st.subheader("📊 Riepilogo Ore")
                    
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Totale Ore", f"{metrics['ore_mensili']} ore")
                    col2.metric("Ore Previste", f"{metrics['target_ore']} ore")
                    col3.metric("Ore Mancanti", f"{metrics['ore_mancanti']} ore")
                    
                    st.write("---")
                    st.markdown("**Dettaglio Turni:**")
                    cols = st.columns(3)
                    shift_items = list(metrics['shift_counts'].items())
                    
                    for i in range(0, len(shift_items), 3):
                        with cols[0]:
                            if i < len(shift_items):
                                s, count = shift_items[i]
                                ore = metrics['ore_totali'][s]
                                st.write(f"**{s}:** {count} ({ore} ore)")
                        with cols[1]:
                            if i+1 < len(shift_items):
                                s, count = shift_items[i+1]
                                ore = metrics['ore_totali'][s]
                                st.write(f"**{s}:** {count} ({ore} ore)")
                        with cols[2]:
                            if i+2 < len(shift_items):
                                s, count = shift_items[i+2]
                                ore = metrics['ore_totali'][s]
                                st.write(f"**{s}:** {count} ({ore} ore)")
                    
                    st.write("---")
                    st.write(f"**Giorni totali:** {metrics['days_in_month']}")
                    st.write(f"**Festività:** {metrics['festivita_count']}")
                    if metrics['festivita_nomi']:
                        st.write(f"**Nomi festività:** {', '.join(metrics['festivita_nomi'])}")
        else:
            st.error("❌ Nessun dato rilevato nel documento!")

if __name__ == '__main__':
    main()
