import pytesseract
import streamlit as st
import numpy as np
from PIL import Image
import re
import pandas as pd
import calendar
from dateutil.easter import easter
from datetime import datetime
from io import BytesIO
import pyexcel
from icalendar import Calendar, Event

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
    'PN': 17, 'REC': -6, 'F': 6, 'S': 0,
    'MAL': 6  # Aggiunto MAL per malattia
}

SHIFT_COLORS = {
    'M': '#ADD8E6',  # Azzurro per Mattina
    'P': '#0000FF',  # Blu per Pomeriggio
    'N': '#800080',  # Viola per Notte
    'S': '#FFA07A',  # Colore attuale per Smonto
    'R': '#90EE90',  # Verde per Riposo
    'F': '#FFD700',  # Oro per Ferie
    'MAL': '#FF4444'  # Rosso per Malattia
}

MONTH_COLORS = {
    'Gennaio': '#FF6F61', 'Febbraio': '#6B5B95',
    'Marzo': '#88B04B', 'Aprile': '#F7CAC9',
    'Maggio': '#92A8D1', 'Giugno': '#955251',
    'Luglio': '#B565A7', 'Agosto': '#009B77',
    'Settembre': '#DD4124', 'Ottobre': '#D65076',
    'Novembre': '#45B8AC', 'Dicembre': '#EFC050'
}

# Custom CSS styling
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@500&family=Rajdhani:wght@500&display=swap');
        
        .main {
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
            color: #e6e6e6;
        }
        
        h1 {
            font-family: 'Orbitron', sans-serif;
            color: #00ff9d !important;
            text-shadow: 0 0 10px #00ff9d88;
            text-align: center;
        }
        
        .subheader {
            font-family: 'Rajdhani', sans-serif;
            color: #7f8fa6 !important;
            text-align: center;
            margin-top: -15px !important;
        }
        
        .stSelectbox div div div input {
            background-color: #16213e !important;
            color: #e6e6e6 !important;
        }
        
        .footer {
            position: fixed;
            bottom: 10px;
            right: 20px;
            color: #7f8fa6;
            font-family: 'Rajdhani', sans-serif;
            font-size: 0.8em;
        }
        
        .negative {
            color: #ff4444 !important;
            font-weight: bold;
        }
        
        .positive {
            color: #00ff9d !important;
            font-weight: bold;
        }
        
        .instructions {
            background-color: #16213e;
            padding: 15px;
            border-radius: 10px;
            margin-bottom: 20px;
            color: #e6e6e6;
        }
        
        .instructions strong {
            color: #00ff9d;
        }
    </style>
""", unsafe_allow_html=True)

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

def extract_from_ics(ics_file):
    try:
        shifts = []
        cal = Calendar.from_ical(ics_file.read())
        
        for component in cal.walk():
            if component.name == "VEVENT":
                summary = component.get('summary', '').upper()
                if "MATTINA" in summary:
                    shifts.append('M')
                elif "POMERIGGIO" in summary:
                    shifts.append('P')
                elif "NOTTE" in summary:
                    shifts.append('N')
                elif "SMONTO" in summary:
                    shifts.append('S')
                elif "ASSENZA" in summary:
                    shifts.append('ASS')
        
        return shifts
        
    except Exception as e:
        st.error(f"Errore lettura ICS: {str(e)}")
        return []

def adjust_shifts(shifts):
    """Gestisce le assenze e le regole sul turno di notte e smonto"""
    adjusted_shifts = []
    previous_shift = None
    
    for shift in shifts:
        if shift == 'ASS':
            adjusted_shifts.append('ASS')
            previous_shift = None  # Ignora il turno precedente e successivo
        elif shift == 'N':  # Se è una Notte
            if previous_shift != 'N':  # Se non è preceduta da un'altra Notte
                adjusted_shifts.append('N')
                adjusted_shifts.append('S')  # Aggiungi Smonto dopo la Notte
                previous_shift = 'S'  # La S è il turno successivo alla Notte
            else:
                adjusted_shifts.append('N')
                previous_shift = 'N'  # Continua con Notte
        else:
            adjusted_shifts.append(shift)
            previous_shift = shift
    
    return adjusted_shifts

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

        cal = calendar.Calendar()
        sundays = 0
        for week in cal.monthdays2calendar(year, month_num):
            for day, weekday in week:
                if day != 0 and weekday == 6:
                    sundays += 1
        
        valid_shifts = shifts[:days_in_month]
        adjusted_shifts = adjust_shifts(valid_shifts)  # Regola i turni
        shift_counts = {s: adjusted_shifts.count(s) for s in ORE_MAP}
        ore_totali = {s: count * ORE_MAP[s] for s, count in shift_counts.items()}
        
        ore_mensili = sum(ore for s, ore in ore_totali.items() if s not in ['R', 'S', 'REC'])
        target_ore = (days_in_month - sundays - festivita_count) * 6
        
        differenza = ore_mensili - target_ore
        ore_mancanti = max(-differenza, 0)
        ore_straordinario = max(differenza, 0)

        cal = calendar.Calendar(firstweekday=0)
        month_weeks = cal.monthdayscalendar(year, month_num)
        shifts_per_day = (adjusted_shifts + [''] * days_in_month)[:days_in_month]
        
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
            'ore_straordinario': ore_straordinario,
            'weeks': weeks,
            'sundays': sundays
        }
        
    except Exception as e:
        st.error(f"Errore nei calcoli: {str(e)}")
        return None

def display_metrics(metrics):
    if metrics:
        st.header("Riepilogo Mensile")
        st.subheader(f"Festività: {metrics['festivita_count']} giorni ({', '.join(metrics['festivita_nomi'])})")
        st.subheader(f"Domeniche: {metrics['sundays']}")
        st.subheader(f"Ore Totali Lavorate: {metrics['ore_mensili']} ore")
        st.subheader(f"Ore Target: {metrics['target_ore']} ore")
        st.subheader(f"Ore Straordinario: {metrics['ore_straordinario']} ore")
        st.subheader(f"Ore Mancanti: {metrics['ore_mancanti']} ore")
        st.write("Turni Settimanali:")
        for week in metrics['weeks']:
            st.write(week)
def display_metrics(metrics):
    if metrics:
        st.header("Riepilogo Mensile")
        st.subheader(f"Festività: {metrics['festivita_count']} giorni ({', '.join(metrics['festivita_nomi'])})")
        st.subheader(f"Domeniche: {metrics['sundays']}")
        st.subheader(f"Ore Totali Lavorate: {metrics['ore_mensili']} ore")
        st.subheader(f"Ore Target: {metrics['target_ore']} ore")
        st.subheader(f"Ore Straordinario: {metrics['ore_straordinario']} ore")
        st.subheader(f"Ore Mancanti: {metrics['ore_mancanti']} ore")
        st.write("Turni Settimanali:")
        for week in metrics['weeks']:
            st.write(week)

def main():
    st.title("Gestione Turni Personale")
    
    uploaded_file = st.file_uploader("Carica un file Excel o ICS", type=["xls", "xlsx", "ics"])
    
    if uploaded_file:
        if uploaded_file.name.endswith(('.xls', '.xlsx')):
            shifts = extract_from_excel(uploaded_file)
        elif uploaded_file.name.endswith('.ics'):
            shifts = extract_from_ics(uploaded_file)
        
        if shifts:
            month = st.selectbox("Seleziona il mese", [
                "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
                "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"
            ])
            year = st.number_input("Seleziona l'anno", min_value=2020, max_value=2100, value=datetime.now().year)

            metrics = calculate_metrics(shifts, month, year)
            display_metrics(metrics)

if __name__ == "__main__":
    main()
