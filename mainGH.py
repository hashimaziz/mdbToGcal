import pyodbc
import datetime
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Select first time running or repalcement because of a date shift
ans = ""
while((ans != "f") and (ans != "r")):
    ans = str(input("Enter 'f' if this is your first time running or enter 'r' if you're replacing future dates because of a shift: "))
    ans = ans.lower()

# Access microsoft database
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=F:<path for microsoft access DB>;') # Path for database goes here
cursor = conn.cursor()

# Collect Data from database
cursor.execute('<SQL Query to get julianDates>') # SQL commands to query the database
julianDates = [row[0] for row in cursor.fetchall()]

cursor.execute('<SQL Query to get events for each day>')
events = [row[0] for row in cursor.fetchall()]

cursor.execute('<SQL Query to get hijri dates>')
hijriDates = [row[0] for row in cursor.fetchall()]

i = 0
future = False
today = datetime.date.today()

for date in julianDates:
    dateString = str(date.strftime("%Y-%m-%d")) # Create formatted date for JSON

    # Figures out if we're looking at a future date if replacement is being used
    if(not future and ans == "r"):
        if(dateString == str(today)):
            future = True
        else:
            continue
        
    # Create G-Calendar Event
    event = {
        'summary': hijriDates[i],
        'location': 'Austin, Texas',
        'description': events[i],
        'start': {
            'date': dateString,
            'timezone': 'America/Chicago',
        },
        'end': {
            'date': dateString,
            'timezone': 'America/Chicago',
        }
    }

    SCOPES = ["https://www.googleapis.com/auth/calendar"]

    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("calendar", "v3", credentials=creds)

        event = service.events().insert(calendarId='primary', body=event).execute()
        print('Event created: %s' % (event.get('htmlLink')))

    except HttpError as error:
        print(f"An error occurred: {error}")

    i += 1

print("All events created")