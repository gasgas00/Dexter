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
    'PN': 17, 'REC': -6, 'F': 6, 'S': 0, 'MAL': 6, '-': 0, 'R': 0
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
            position: relative;
        }
        
        .day-number {
            position: absolute;
            top: 5px;
            left: 5px;
            font-size: 0.8em;
            color: #ffffff;
        }
        
        .weekday-header {
            font-weight: bold;
            text-align: center;
            padding: 10px;
            background-color: #444444;
            color: white;
            border-radius: 5px;
            margin: 2px;
        }
        
        .shift-select {
            margin-top: 20px;
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
                elif any(turno in summary for turno in ['MATTINA', 'POMERIGGIO', 'NOTTE', 'SMONTO', 'RECUPERO', 'RIPOSO']):
                    shift = 'M' if 'MATTINA' in summary else \
                            'P' if 'POMERIGGIO' in summary else \
                            'N' if 'NOTTE' in summary else \
                            'S' if 'SMONTO' in summary else \
                            'R' if 'RECUPERO' in summary or 'RIPOSO' in summary else '-'
                    shifts.append({
                        'date': component.get('dtstart').dt,
                        'turno': shift
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

def display_calendar(month, year, shifts, festivita_nomi):
    month_num = list(MONTH_COLORS.keys()).index(month) + 1
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
    
    # Intestazioni giorni della settimana
    weekdays = ['Lunedì', 'Martedì', 'Mercoledì', 'Giovedì', 'Venerdì', 'Sabato', 'Domenica']
    cols = st.columns(7)
    for i, day in enumerate(weekdays):
        cols[i].markdown(f"<div class='weekday-header'>{day}</div>", unsafe_allow_html=True)
    
    # Calendario
    cal = calendar.Calendar(firstweekday=0)
    month_weeks = cal.monthdayscalendar(year, month_num)
    
    for week in month_weeks:
        cols = st.columns(7)
        for i, day in enumerate(week):
            if day == 0:
                cols[i].write("")
            else:
                shift_index = day - 1
                current_shift = shifts[shift_index] if shift_index < len(shifts) else '-'
                color = SHIFT_COLORS.get(current_shift, '#FFFFFF')
                
                with cols[i]:
                    st.markdown(
                        f"<div class='calendar-day' style='background-color: {color};'>"
                        f"<div class='day-number'>{day}</div>",
                        unsafe_allow_html=True
                    )
                    options = ['-', 'M', 'P', 'N', 'MP', 'PN', 'REC', 'F', 'S', 'MAL', 'R']
                    default_index = options.index(current_shift) if current_shift in options else 0
                    key = f"shift_{year}_{month_num}_{day}"
                    new_shift = st.selectbox(
                        label=f"Turno {day}",
                        options=options,
                        index=default_index,
                        key=key,
                        label_visibility="collapsed",
                        on_change=lambda: st.experimental_rerun()
                    )
                    if new_shift != current_shift:
                        shifts[shift_index] = new_shift

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
    
    uploaded_file = st.file_uploader("Carica il planning turni", type=['ics'])
    
    if uploaded_file:
        with st.spinner('Elaborazione in corso...'):
            if uploaded_file.type == "text/calendar":
                shifts, absences = extract_from_ics(uploaded_file)
                if shifts:
                    st.info("Nota: Se i turni non coincidono con quelli effettivi o ci sono modifiche, clicca su una cella del calendario per modificare il turno e aggiornare i dati.")
                    
                    month_num = list(MONTH_COLORS.keys()).index(month) + 1
                    days_in_month = calendar.monthrange(year, month_num)[1]
                    
                    # Crea un dizionario giorno -> turno
                    shifts_dict = {}
                    for s in shifts:
                        if s['date'].month == month_num and s['date'].year == year:
                            shifts_dict[s['date'].day] = s['turno']
                    # Riempie la lista per tutti i giorni del mese
                    shifts_list = [shifts_dict.get(day, '-') for day in range(1, days_in_month + 1)]
                    
                    key = ('ics', month, year)
                    if 'edited_shifts' not in st.session_state:
                        st.session_state.edited_shifts = {}
                    if key not in st.session_state.edited_shifts:
                        st.session_state.edited_shifts[key] = shifts_list.copy()
                    current_shifts = st.session_state.edited_shifts[key]
                    
                    display_calendar(month, year, current_shifts, [])
                    
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
                        st.subheader("📋 Dettaglio Turni e Ore")
                        
                        num_cols = 3
                        cols = st.columns(num_cols)
                        
                        shift_types = [k for k in ORE_MAP.keys() if k != '-']
                        for i, shift in enumerate(shift_types):
                            count = metrics['shift_counts'].get(shift, 0)
                            total_hours = metrics['ore_totali'].get(shift, 0)
                            
                            if count > 0:
                                with cols[i % num_cols]:
                                    st.markdown(f"""
                                        <div style="padding: 15px; background: {SHIFT_COLORS[shift]}; 
                                            border-radius: 10px; margin: 5px; color: {'white' if shift in ['N', 'PN', 'MP'] else 'black'}">
                                            <h4>{shift}</h4>
                                            <p>Turni: {count}</p>
                                            <p>Ore totali: {total_hours}h</p>
                                        </div>
                                    """, unsafe_allow_html=True)
                        
                        st.write("---")
                        st.subheader("📊 Riepilogo Ore")
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Totale Ore Lavorate", f"{metrics['ore_mensili']} ore")
                        col2.metric("Ore Previste", f"{metrics['target_ore']} ore")
                        
                                                if metrics['ore_mancanti'] > 0:
                            col3.markdown(f"<div class='negative'>🟡 Ore Mancanti: {metrics['ore_mancanti']}h</div>", unsafe_allow_html=True)
                        elif metrics['ore_straordinario'] > 0:
                            col3.markdown(f"<div class='positive'>🟢 Ore Straordinario: {metrics['ore_straordinario']}h</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
