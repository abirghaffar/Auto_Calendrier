# -*- coding: utf-8 -*-
"""
Created on Sat Feb 17 12:30:28 2024

@author: abir
"""
import pandas as pd
from ics import Calendar, Event

# Étape 1: Extraction des données Excel
df = pd.read_excel("C:/Users/abir/Documents/Projet _ Conception numerique/EDT_SIGMA.xlsx",
                   sheet_name='Export CSV')
df.head ()
# Convertion le DataFrame en CSV
df.to_csv('mon_fichier.csv', index=False)
df.head()
# Contrôle des données 
from datetime import datetime
for index, row in df.iterrows():
    date_debut = datetime.strptime(row['Start Date'].strftime('%Y-%m-%d') 
                                   + ' ' + row['Start Time'].strftime('%H:%M:%S'),
                                   '%Y-%m-%d %H:%M:%S')

    date_fin = datetime.strptime(row['End Date'].strftime('%Y-%m-%d') 
                                 + ' ' + row['End Time'].strftime('%H:%M:%S'),
                                 '%Y-%m-%d %H:%M:%S')
    T=True 
    if date_fin < date_debut:
        T=False 
        print(f'erreur dans la ligne {index}')
    
# Créer un nouveau calendrier
cal = Calendar()
# Parcourir chaque ligne du DataFrame
for index, row in df.iterrows():
    event = Event()
    event.name = row['Subject']
    event.begin = row['Start Date'].strftime('%Y-%m-%d') 
    + ' ' + row['Start Time'].strftime('%H:%M:%S')
    event.end = row['End Date'].strftime('%Y-%m-%d')
    + ' ' + row['End Time'].strftime('%H:%M:%S')
    event.location = str(row ['Location'])
    event.description = row ['Description']
    cal.events.add(event)
# Écrire le calendrier dans un fichier .ics
with open('mon_calendrier.ics', 'w') as my_file:
    my_file.writelines(cal)
### Autrement     
with open('mon_calendrier.ics', 'w') as my_file:
    my_file.write(cal.serialize())

