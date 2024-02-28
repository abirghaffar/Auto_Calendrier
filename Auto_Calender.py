# -*- coding: utf-8 -*-
"""
Created on Fri Feb 23 21:45:04 2024

@author: abir
"""
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime

# Chemin vers votre fichier Excel
excel_path = "C:/Users/abir/Documents/Auto_Calendrier/Calandrier.xlsx"

# Charger le fichier Excel
df = pd.read_excel(excel_path)
df.index
df.Start_Date = pd.Series(df.Start_Date).fillna(method='ffill')
df.End_Date = pd.Series(df.End_Date).fillna(method='ffill')

# Chemin vers le fichier credentials.json
creds = None
creds_path = "C:/Users/abir/Documents/Auto_Calendrier/credentials.json.json"

# Si nécessaire, changez le scope selon vos besoins
SCOPES = ['https://www.googleapis.com/auth/calendar']

# Authentification et création du service
flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
creds = flow.run_local_server(port=0)
service = build('calendar', 'v3', credentials=creds)

# Création des événements à partir du DataFrame
for i in range(len(df)):
    row = df.iloc[i]
    print(row['Start_Date'])
    print(datetime.strptime(str(row['Start_Date']), 
                                       '%Y-%m-%d %H:%M:%S'))
    print(row['End_Time'])
    print(datetime.strptime(str(row['End_Time']), 
                                       '%Y-%m-%d %H:%M:%S'))
# Conversion des chaînes de dates en objets datetime
start_datetime = datetime.strptime(str(row['Start_Date']).strftime('%Y-%m-%d') 
                                   + ' ' 
                                   + str(row['Start_Time']).strftime('%H:%M:%S'),
                                   '%Y-%m-%d %H:%M:%S')
                                   
end_datetime = datetime.strptime(str(row['End_Date']).strftime('%Y-%m-%d') 
                                 + ' ' + 
                                 str(row['End_Time']).strftime('%H:%M:%S'),
                                 '%Y-%m-%d %H:%M:%S')
                                 
event = {
      'summary': row['Unit'],
      'location': row['Location'],
      'description': str(row['Description'])+ str(row['Intervenant']),
      'start': {
        'dateTime': start_datetime.isoformat()+'-'+ end_datetime.isoformat(),
        'timeZone': 'Europe/Paris',
      },
      'end': {
        'dateTime': end_datetime.isoformat(),
        'timeZone': 'Europe/Paris',
      },
    }

event = service.events().insert(calendarId='c519ec700f2eef68a4a7e75608be062c9bd7430c55c2fa2a2517f190647e78f0@group.calendar.google.com', body=event).execute()

print('Événement créé : %s' % (event.get('htmlLink')))

    
