Nom du projet : Auto_Agenda 

Sujet du projet CNUM : Automatisation du formatage du calendrier des cours de master SIGMA de format Excel vers le format ical/ics

Produit à développer : Un code réalisé en langage Python qui permet de créer et synchroniser un Agenda Google contenant le calandrier des cours de master SIGMA avec la possibilité de faire une sélection par intervenant ou par UE
Démarche du travail: 

* Etape 1 : formatage de base des données intiale,qui réprésente un fichier Excel contenant les calandriers des cours de Master SIGMA pour l'année 2023/2024, et création d'un nouveau fichier de la même format .xls mais plus structuré en utilisant le code python copié 
çi-dessous pour automatiser l'étape : 
###############
import os
import pandas as pd
import openpyxl
from openpyxl.styles import Border, Side
# Load the Excel file into a DataFrame from the second sheet named "Export CSV"
df = pd.read_excel("C:/Users/abir/Desktop/EDT_SIGMA_1.xlsx",
                   sheet_name="Export CSV")

# Display the first few rows of the DataFrame
print(df.head())

# Rename columns with correct names
df.columns = ['Subject', 'Start_Date', 'Start_Time', 'End_Date', 'End_Time', 'Location', 'Description','Catégorie']

# Swap columns 3 and 4 (indexing starts from 0)
df = df[['Subject', 'Start_Date', 'End_Date', 'Start_Time', 'End_Time', 'Location', 'Description','Catégorie']]

# Define lighter colors for each column
column_colors = ['#FFCCFF', '#FFFFCC', '#FFFFCC', '#CCFFFF', '#CCFFFF','#CCFFCC', '#FFD700','#FFD700']  # Lighter versions of the original colors, with green instead of red

# Apply the color formatting to each column
styled_df = df.style.apply(lambda x: ['background-color: %s' % color for color in column_colors], axis=1)

# Set all cells content to be centered
styled_df = styled_df.set_properties(**{'text-align': 'center'})

# Make the headers bold
styled_df = styled_df.set_table_styles([{
    'selector': 'th',
    'props': [('font-weight', 'bold')]
}])

# Set columns 1 and 3 as italic
styled_df = styled_df.set_properties(subset=['Start_Date', 'End_Date'], **{'font-style': 'italic'})

# Define border styles with black color
border_style = {
    'selector': '',
    'props': [('border', '1px solid black')]  # Border style with black color
}

# Set border styles
styled_df = styled_df.set_table_styles([border_style], overwrite=False)

# Set cell wrap text to True
styled_df = styled_df.set_properties(**{'white-space': 'pre-wrap'})

# Convert 'Start Date' and 'End Date' columns to datetime
date_columns = ['Start_Date', 'End_Date']
for col in date_columns:
    df[col] = pd.to_datetime(df[col], errors='coerce').dt.date

# Save the styled DataFrame to a new Excel file
excel_writer = pd.ExcelWriter('modified_agenda.xlsx', engine='openpyxl')
styled_df.to_excel(excel_writer, index=False)

# Get the ExcelWriter workbook object
workbook = excel_writer.book

# Set fixed column width for all columns except column 7
default_column_width = 20  # Default width for other columns
specific_column_width = 50  # Width for column 7
for idx, column in enumerate(df.columns, start=1):
    column_letter = openpyxl.utils.get_column_letter(idx)
    column_width = specific_column_width if idx == 6  else default_column_width
    workbook['Sheet1'].column_dimensions[column_letter].width = column_width

# Add borders to all cells
ws = workbook.active
max_row = ws.max_row
max_col = ws.max_column
for row in ws.iter_rows():
    for cell in row:
        cell.border = Border(left=Side(border_style='thin', color='000000'),
                             right=Side(border_style='thin', color='000000'),
                             top=Side(border_style='thin', color='000000'),
                             bottom=Side(border_style='thin', color='000000'))

# Save the workbook
excel_writer.save()
################################

* Etape 2: Passage du format xls vers un format ics qui sera importer automatiquement dans Agenda Google après avoir realisé une étape d'activation des API Google Calander et télechargement du fichier .json qu'est disponible en ligne via le scipt çi-dessous : 
################################
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime

# Chemin vers votre fichier Excel
excel_path = "C:/Users/abir/Documents/Projet _ Conception numerique/Calebdrier_M1_M2.xlsx"

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

event = service.events().insert(calendarId='c519ec700f2eef68a4a7e75608be062c9bd7430c55c2fa2a2517f190647e78f0@group.calendar.google.com',
                                body=event).execute()

print('Événement créé : %s' % (event.get('htmlLink')))

############################################

* Etape 3 : Exportation du calandrier pour retourner à la base de données initiale  
  
 : 22
