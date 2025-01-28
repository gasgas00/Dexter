import pytesseract
import streamlit as st
import numpy as np
from PIL import Image
import re
import pandas as pd
import icalendar
from dateutil import parser
from datetime import datetime
from io import BytesIO
import pyexcel
import icalendar

# Configurazione Tesseract per macOS
pytesseract.pytesseract.tesseract_cmd = '/usr/local/bin/tesseract'

# Lista fissa dei nominativi (ora non necessaria)
FIXED_NAMES = [
    "TATIANA ARTENIE", "GIANLUCA MATARRESE", "SIMONA LUPU", 
    "CHIAROTTI EMMA", "RUGGIANO MASSIMO", "LIVIA BURRAI", 
    "ANDREA PALLI", "TIZNADO MARLY", "SPADA MATTEO", "CUTAIA MARCO"
]

# Mappa dei turni con nuove modifiche
SHIFT_COLORS = {
    'M': '#ADD8E6',  # Azzurro
    'P': '#0000FF',  # Blu
    'N': '#8A2BE2',  # Viola
    'S': '#FFA07A',  # Colore attuale per SMONTO
    'R': '#32CD32',  # Verde per Riposo
    'F': '#FFD700',  # Giallo per Ferie
    'MAL': '#FF6347',  # Rosso per Malattia
}

# Istruzioni per il caricamento del file Zucchetti
def show_instructions():
    st.markdown("""
    **ISTRUZIONI per caricare il file ZUCCHETTI:**

    1. Entra su Zucchetti
    2. Clicca in alto a sinistra sul menu rappresentato dai quadratini ed entra su **Zscheduling**
    3. In alto comparirà la dicitura **Calendario operatore**, clicca su essa
    4. **Ora in alto a destra va cambiata la dicitura da Settimanale a Mensile**
    5. Una volta selezionato il calendario mensile, clicca sopra di esso per scaricare il file
    """, unsafe_allow_html=True)

def extract_from_ics(ics_file):
    try:
        calendar_data = icalendar.Calendar.from_ical(ics_file.read())
        shifts = {}
        current_date = None

        for component in calendar_data.walk():
            if component.name == "VEVENT":
                start = component.get('DTSTART').dt
                end = component.get('DTEND').dt
                summary = component.get('SUMMARY')

                if start and summary:
                    shift_type = summary.strip().upper()
                    # Aggiungi il turno al dizionario con la data
                    if start.date() not in shifts:
                        shifts[start.date()] = []
                    if shift_type in SHIFT_COLORS:
                        shifts[start.date()].append(shift_type)

        return shifts

    except Exception as e:
        st.error(f"Errore nell'elaborazione del file ICS: {str(e)}")
        return {}

def calculate_metrics_from_ics(shifts, month, year):
    try:
        # Logica per calcolare le ore come prima (senza tenere conto delle ore nel calendario)
        month_num = [
            'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
            'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'
        ].index(month) + 1
        
        days_in_month = calendar.monthrange(year, month_num)[1]
        
        # Calcola i turni per ogni giorno
        valid_shifts = [shifts.get(datetime(year, month_num, day), []) for day in range(1, days_in_month+1)]

        shift_counts = {s: sum(shift.count(s) for shift in valid_shifts) for s in SHIFT_COLORS}
        ore_totali = {s: count * 6 for s, count in shift_counts.items()}  # Ogni turno vale 6 ore

        ore_mensili = sum(ore_totali.values())
        target_ore = (days_in_month - len([1 for i in range(1, days_in_month+1) if calendar.weekday(year, month_num, i) == 6])) * 6

        # Gestisci il caso delle assenze
        ore_mancanti = max(0, target_ore - ore_mensili)
        ore_straordinario = max(0, ore_mensili - target_ore)

        return {
            'ore_mensili': ore_mensili,
            'ore_mancanti': ore_mancanti,
            'ore_straordinario': ore_straordinario,
            'shift_counts': shift_counts,
            'ore_totali': ore_totali
        }
        
    except Exception as e:
        st.error(f"Errore nei calcoli: {str(e)}")
        return None

def main():
    st.markdown("<h1>Eureka!</h1>", unsafe_allow_html=True)
    st.markdown("<div class='subheader'>L'App di analisi turni dell'UTIC</div>", unsafe_allow_html=True)
    
    # Mostra le istruzioni
    show_instructions()

    # Caricamento file
    uploaded_file = st.file_uploader("Carica il planning turni", type=['xlsx', 'xls', 'ics'])

    if uploaded_file:
        with st.spinner('Elaborazione in corso...'):
            if uploaded_file.type == "application/vnd.ms-excel" or uploaded_file.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                people_shifts = extract_from_excel(uploaded_file)
            elif uploaded_file.type == "text/calendar":
                people_shifts = extract_from_ics(uploaded_file)
            else:
                st.error("Formato non supportato!")
                return

            if people_shifts:
                # Gestisci i turni come per l'Excel, senza la selezione del nominativo (per ICS)
                shifts = people_shifts
                metrics = calculate_metrics_from_ics(shifts, 'Gennaio', 2025)
                
                if metrics:
                    # Visualizza il calendario e il riepilogo ore
                    display_month('Gennaio', 2025, [])
                    st.subheader("📊 Riepilogo Ore")
                    col1, col2 = st.columns(2)
                    col1.metric("Totale Ore Lavorate", f"{metrics['ore_mensili']} ore")
                    col2.metric("Ore Previste", f"{metrics['ore_mancanti']} ore")

                    # Dettaglio turni
                    st.write("Dettaglio turni:")
                    for shift, count in metrics['shift_counts'].items():
                        st.write(f"{shift}: {count} turni")

if __name__ == '__main__':
    main()
