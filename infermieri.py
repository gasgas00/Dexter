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
import plotly.express as px

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
    'PN': 17, 'REC': -6, 'F': 6, 'S': 0, 'MAL': 6, '-': 0
}

SHIFT_COLORS = {
    'M': '#00BFFF',  # Blu acceso
    'P': '#FFA500',  # Arancione
    'N': '#800080',  # Viola
    'R': '#00FF00', 
    'S': '#FFA07A',
    'F': '#FFD700',
    'PN': '#FF69B4',
    'MP': '#9370DB',
    'REC': '#D3D3D3',
    'MAL': '#FF4444',
    '-': '#333333'  # Grigio scuro per celle vuote
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
            background: linear-gradient(135deg, #0a0a0a 0%, #1a1a1a 100%);
            color: #e6e6e6;
        }
        
        h1 {
            font-family: 'Orbitron', sans-serif;
            color: #00ff9d !important;
            text-shadow: 0 0 10px #00ff9d88;
            text-align: center;
            font-size: 8em;
        }
        
        .calendar-day {
            margin: 5px;
            padding: 10px;
            border-radius: 5px;
            text-align: center;
            font-size: 1.2em;
        }
        
        .shift-select {
            margin-top: 5px;
            width: 100%;
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
        cal = Calendar.from_ical(ics_file.read())
        shifts = []
        absences = []

        for component in cal.walk():
            if component.name == "VEVENT":
                summary = component.get('summary', '').upper()
                if 'ASSENZA' in summary:
                    absences.append({
                        'date': component.get('dtstart').dt,
                        'type': None
                    })
                elif any(turno in summary for turno in ['MATTINA', 'POMERIGGIO', 'NOTTE', 'SMONTO']):
                    shifts.append({
                        'date': component.get('dtstart').dt,
                        'turno': 'M' if 'MATTINA' in summary else 'P' if 'POMERIGGIO' in summary else 'N' if 'NOTTE' in summary else 'S'
                    })

        return shifts, absences

    except Exception as e:
        st.error(f"Errore lettura ICS: {str(e)}")
        return [], []

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
        
        # Filtra i turni validi escludendo '-'
        valid_shifts = [s for s in shifts[:days_in_month] if s != '-']
        shift_counts = {s: valid_shifts.count(s) for s in ORE_MAP if s != '-'}
        ore_totali = {s: count * ORE_MAP[s] for s, count in shift_counts.items()}
        
        ore_mensili = sum(ore for s, ore in ore_totali.items() if s not in ['R', 'S', 'REC'])
        target_ore = (days_in_month - sundays - festivita_count) * 6
        
        differenza = ore_mensili - target_ore
        ore_mancanti = max(-differenza, 0)
        ore_straordinario = max(differenza, 0)

        cal = calendar.Calendar(firstweekday=0)
        month_weeks = cal.monthdayscalendar(year, month_num)
        shifts_per_day = (shifts[:days_in_month] + ['-'] * (days_in_month - len(shifts)))[:days_in_month]
        
        weeks = []
        for week in month_weeks:
            week_data = []
            for day in week:
                if day == 0:
                    week_data.append((0, ''))
                else:
                    shift = shifts_per_day[day-1] if (day-1) < len(shifts_per_day) else '-'
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
    st.markdown("<h1>Eureka!</h1>", unsafe_allow_html=True)
    st.markdown("<div class='subheader'>L'App di analisi turni dell'UTIC</div>", unsafe_allow_html=True)
    
    st.markdown("""
        <div class='instructions'>
            <strong>ISTRUZIONI PER ZUCCHETTI:</strong>
            <ol>
                <li>Entrare su Zucchetti</li>
                <li>Cliccare in alto a sinistra sul menu rappresentato dai quadratini ed entrare su Zscheduling</li>
                <li>In alto comparirà la dicitura "Calendario Operatore", cliccare su essa</li>
                <li><strong>Cambiare la visualizzazione da "Settimanale" a "Mensile"</strong></li>
                <li>Cliccare su "Esporta" e scaricare il file in formato ICS</li>
            </ol>
        </div>
    """, unsafe_allow_html=True)
    
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
    
    uploaded_file = st.file_uploader("Carica il planning turni", type=['xlsx', 'xls', 'ics'])
    
    if uploaded_file:
        with st.spinner('Elaborazione in corso...'):
            if uploaded_file.type == "text/calendar":
                shifts, absences = extract_from_ics(uploaded_file)
                if shifts:
                    st.info("Nota: Se i turni non coincidono con quelli effettivi o ci sono modifiche, clicca su una cella del calendario per modificare il turno e aggiornare i dati.")
                    st.subheader("📅 Turni")
                    display_month(month, year, [])
                    
                    if absences:
                        st.warning("⚠️ Sono presenti delle assenze. Seleziona il tipo di assenza per ciascuna.")
                        for absence in absences:
                            absence_date = absence['date'].strftime('%Y-%m-%d')
                            absence_type = st.selectbox(
                                f"Tipo di assenza per il giorno {absence_date}:",
                                ['Ferie (F)', 'Malattia (MAL)']
                            )
                            absence['type'] = 'F' if 'Ferie' in absence_type else 'MAL'
                    
                    shifts_list = [s['turno'] for s in shifts]
                    month_num = list(MONTH_COLORS.keys()).index(month) + 1
                    days_in_month = calendar.monthrange(year, month_num)[1]
                    shifts_list = (shifts_list[:days_in_month] + ['-'] * (days_in_month - len(shifts_list)))[:days_in_month]
                    
                    key = ('ics', month, year)
                    if 'edited_shifts' not in st.session_state:
                        st.session_state.edited_shifts = {}
                    if key not in st.session_state.edited_shifts:
                        st.session_state.edited_shifts[key] = shifts_list.copy()
                    current_shifts = st.session_state.edited_shifts[key]
                    
                    st.subheader("📅 Calendario Turni")
                    cal = calendar.Calendar(firstweekday=0)
                    month_weeks = cal.monthdayscalendar(year, list(MONTH_COLORS.keys()).index(month) + 1)
                    
                    for week in month_weeks:
                        cols = st.columns(7)
                        for i, day in enumerate(week):
                            if day == 0:
                                cols[i].write("")
                            else:
                                shift_index = day - 1
                                current_shift = current_shifts[shift_index] if shift_index < len(current_shifts) else '-'
                                color = SHIFT_COLORS.get(current_shift, '#FFFFFF')
                                
                                with cols[i]:
                                    st.markdown(
                                        f"<div class='calendar-day' style='background-color: {color};'>"
                                        f"<strong>{day}</strong></div>",
                                        unsafe_allow_html=True
                                    )
                                    options = ['-'] + [k for k in ORE_MAP.keys() if k != '-']
                                    default_index = options.index(current_shift) if current_shift in options else 0
                                    new_shift = st.selectbox(
                                        label=f"Turno {day}",
                                        options=options,
                                        index=default_index,
                                        key=f"shift_ics_{day}",
                                        label_visibility="collapsed"
                                    )
                                    if new_shift != current_shift:
                                        st.session_state.edited_shifts[key][shift_index] = new_shift
                    
                    metrics = calculate_metrics(current_shifts, month, year)
                    
                    if metrics:
                        st.write("---")
                        st.subheader("📊 Grafico a Torta - Dettaglio Turni")
                        shift_counts = metrics['shift_counts']
                        df_pie = pd.DataFrame({
                            'Turno': list(shift_counts.keys()),
                            'Conteggio': list(shift_counts.values())
                        })
                        fig = px.pie(df_pie, values='Conteggio', names='Turno', title='Distribuzione dei Turni')
                        st.plotly_chart(fig)
                        
                        st.write("---")
                        st.subheader("📊 Riepilogo Ore")
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Totale Ore Lavorate", f"{metrics['ore_mensili']} ore")
                        col2.metric("Ore Previste", f"{metrics['target_ore']} ore")
                        
                        if metrics['ore_mancanti'] > 0:
                            col3.markdown(f"<div class='negative'>🟡 Ore Mancanti: {metrics['ore_mancanti']}h</div>", unsafe_allow_html=True)
                        elif metrics['ore_straordinario'] > 0:
                            col3.markdown(f"<div class='positive'>🟢 Ore Straordinario: {metrics['ore_straordinario']}h</div>", unsafe_allow_html=True)

                        st.write("---")
                        st.markdown("**Dettaglio Turni:**", unsafe_allow_html=True)
                        shift_items = list(metrics['shift_counts'].items())
                        num_columns = 3
                        num_rows = (len(shift_items) + num_columns - 1) // num_columns

                        for row in range(num_rows):
                            cols = st.columns(num_columns)
                            for col in range(num_columns):
                                index = row * num_columns + col
                                if index < len(shift_items):
                                    s, count = shift_items[index]
                                    ore = metrics['ore_totali'][s]
                                    color = SHIFT_COLORS.get(s, '#FFFFFF')
                                    with cols[col]:
                                        st.markdown(
                                            f"<div style='background-color: {color}; padding: 10px; border-radius: 5px; margin: 5px; text-align: center;'>"
                                            f"<strong style='font-size: 1.2em;'>{s}</strong><br>"
                                            f"Turni: {count}<br>"
                                            f"Ore: {ore}</div>",
                                            unsafe_allow_html=True
                                        )
                        
                        st.write("---")
                        st.write(f"**Giorni totali:** {metrics['days_in_month']}")
                        st.write(f"**Domeniche:** {metrics['sundays']}")
                        st.write(f"**Festività:** {metrics['festivita_count']}")
                        if metrics['festivita_nomi']:
                            st.write(f"**Festività:** {', '.join(metrics['festivita_nomi'])}")
            
            else:
                people_shifts = extract_from_excel(uploaded_file)
                if people_shifts:
                    st.info("Nota: Se i turni non coincidono con quelli effettivi o ci sono modifiche, clicca su una cella del calendario per modificare il turno e aggiornare i dati.")
                    st.subheader("📅 Turni Rilevati")
                    display_month(month, year, [])
                    
                    selected_operator = st.selectbox("Seleziona l'operatore:", list(people_shifts.keys()))
                    original_shifts = people_shifts[selected_operator]
                    
                    month_num = list(MONTH_COLORS.keys()).index(month) + 1
                    days_in_month = calendar.monthrange(year, month_num)[1]
                    original_shifts_padded = (original_shifts[:days_in_month] + ['-'] * (days_in_month - len(original_shifts)))[:days_in_month]
                    
                    key = (selected_operator, month, year)
                    if 'edited_shifts' not in st.session_state:
                        st.session_state.edited_shifts = {}
                    if key not in st.session_state.edited_shifts:
                        st.session_state.edited_shifts[key] = original_shifts_padded.copy()
                    current_shifts = st.session_state.edited_shifts[key]
                    
                    st.subheader("📅 Calendario Turni")
                    cal = calendar.Calendar(firstweekday=0)
                    month_weeks = cal.monthdayscalendar(year, list(MONTH_COLORS.keys()).index(month) + 1)
                    
                    for week in month_weeks:
                        cols = st.columns(7)
                        for i, day in enumerate(week):
                            if day == 0:
                                cols[i].write("")
                            else:
                                shift_index = day - 1
                                current_shift = current_shifts[shift_index] if shift_index < len(current_shifts) else '-'
                                color = SHIFT_COLORS.get(current_shift, '#FFFFFF')
                                
                                with cols[i]:
                                    st.markdown(
                                        f"<div class='calendar-day' style='background-color: {color};'>"
                                        f"<strong>{day}</strong></div>",
                                        unsafe_allow_html=True
                                    )
                                    options = ['-'] + [k for k in ORE_MAP.keys() if k != '-']
                                    default_index = options.index(current_shift) if current_shift in options else 0
                                    new_shift = st.selectbox(
                                        label=f"Turno {day}",
                                        options=options,
                                        index=default_index,
                                        key=f"shift_{key}_{day}",
                                        label_visibility="collapsed"
                                    )
                                    if new_shift != current_shift:
                                        st.session_state.edited_shifts[key][shift_index] = new_shift
                    
                    metrics = calculate_metrics(current_shifts, month, year)
                    
                    if metrics:
                        st.write("---")
                        st.subheader("📊 Grafico a Torta - Dettaglio Turni")
                        shift_counts = metrics['shift_counts']
                        df_pie = pd.DataFrame({
                            'Turno': list(shift_counts.keys()),
                            'Conteggio': list(shift_counts.values())
                        })
                        fig = px.pie(df_pie, values='Conteggio', names='Turno', title='Distribuzione dei Turni')
                        st.plotly_chart(fig)
                        
                        st.write("---")
                        st.subheader("📊 Riepilogo Ore")
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Totale Ore Lavorate", f"{metrics['ore_mensili']} ore")
                        col2.metric("Ore Previste", f"{metrics['target_ore']} ore")
                        
                        if metrics['ore_mancanti'] > 0:
                            col3.markdown(f"<div class='negative'>🟡 Ore Mancanti: {metrics['ore_mancanti']}h</div>", unsafe_allow_html=True)
                        elif metrics['ore_straordinario'] > 0:
                            col3.markdown(f"<div class='positive'>🟢 Ore Straordinario: {metrics['ore_straordinario']}h</div>", unsafe_allow_html=True)
                        
                        st.write("---")
                        st.markdown("**Dettaglio Turni:**", unsafe_allow_html=True)
                        shift_items = list(metrics['shift_counts'].items())
                        num_columns = 3
                        num_rows = (len(shift_items) + num_columns - 1) // num_columns

                        for row in range(num_rows):
                            cols = st.columns(num_columns)
                            for col in range(num_columns):
                                index = row * num_columns + col
                                if index < len(shift_items):
                                    s, count = shift_items[index]
                                    ore = metrics['ore_totali'][s]
                                    color = SHIFT_COLORS.get(s, '#FFFFFF')
                                    with cols[col]:
                                        st.markdown(
                                            f"<div style='background-color: {color}; padding: 10px; border-radius: 5px; margin: 5px; text-align: center;'>"
                                            f"<strong style='font-size: 1.2em;'>{s}</strong><br>"
                                            f"Turni: {count}<br>"
                                            f"Ore: {ore}</div>",
                                            unsafe_allow_html=True
                                        )
                        
                        st.write("---")
                        st.write(f"**Giorni totali:** {metrics['days_in_month']}")
                        st.write(f"**Domeniche:** {metrics['sundays']}")
                        st.write(f"**Festività:** {metrics['festivita_count']}")
                        if metrics['festivita_nomi']:
                            st.write(f"**Festività:** {', '.join(metrics['festivita_nomi'])}")

if __name__ == "__main__":
    main()
