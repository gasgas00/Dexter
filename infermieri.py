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
import icalendar

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
    'PN': 17, 'REC': -6, 'F': 6, 'S': 0
}

SHIFT_COLORS = {
    'M': '#ADD8E6',  # Azzurro
    'P': '#0000FF',  # Blu
    'N': '#9370DB',  # Viola
    'R': '#90EE90',  # Verde
    'S': '#FFA07A',  # Arancione
    'F': '#FFD700',
    'PN': '#FF69B4',
    'MP': '#9370DB',
    'REC': '#D3D3D3'
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
        ics_content = ics_file.read()
        calendar = icalendar.Calendar.from_ical(ics_content)

        people_shifts = {}

        for component in calendar.walk():
            if component.name == "VEVENT":
                summary = component.get("SUMMARY")
                name, shift = None, None

                if summary:
                    summary = normalize_name(summary)
                    for fixed_name in FIXED_NAMES:
                        if fixed_name in summary:
                            name = fixed_name
                            break

                    if name:
                        shift_match = re.search(r'(MATTINA|POMERIGGIO|NOTTE|RIPOSO|FERIE)', summary)
                        if shift_match:
                            shift = shift_match.group(0)[0]  # Usa solo l'iniziale del turno

                    if name and shift:
                        if name not in people_shifts:
                            people_shifts[name] = []
                        people_shifts[name].append(shift)

        return people_shifts
    except Exception as e:
        st.error(f"Errore lettura ICS: {str(e)}")
        return {}

def calculate_metrics(shifts, month, year):
    try:
        month_num = [
            'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
            'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'
        ].index(month) + 1
        
        days_in_month = calendar.monthrange(year, month_num)[1]
import React, { useState } from "react";
import { Card, CardContent } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { motion } from "framer-motion";

const App = () => {
  const [items, setItems] = useState(["Item 1", "Item 2", "Item 3"]);
  const [newItem, setNewItem] = useState("");

  const addItem = () => {
    if (newItem.trim()) {
      setItems([...items, newItem.trim()]);
      setNewItem("");
    }
  };

  const removeItem = (index) => {
    const updatedItems = items.filter((_, i) => i !== index);
    setItems(updatedItems);
  };

  return (
    <div className="min-h-screen bg-gray-100 flex flex-col items-center py-10">
      <h1 className="text-3xl font-bold mb-6">Item List Manager</h1>
      <Card className="w-full max-w-md p-4">
        <CardContent>
          <div className="flex gap-2 mb-4">
            <Input
              value={newItem}
              onChange={(e) => setNewItem(e.target.value)}
              placeholder="Add a new item"
              className="flex-1"
            />
            <Button onClick={addItem} className="bg-blue-500 text-white">
              Add
            </Button>
          </div>
          <ul className="space-y-2">
            {items.map((item, index) => (
              <motion.li
                key={index}
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="flex justify-between items-center bg-white p-2 rounded shadow"
              >
                <span>{item}</span>
                <Button
                  onClick={() => removeItem(index)}
                  className="bg-red-500 text-white px-2 py-1 text-sm"
                >
                  Remove
                </Button>
              </motion.li>
            ))}
          </ul>
        </CardContent>
      </Card>
    </div>
  );
};

export default App;
