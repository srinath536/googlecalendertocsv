from __future__ import print_function
import datetime
import csv
import os
from collections import defaultdict
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
EXPORT_FOLDER = 'calendar_event_exports'

def authenticate_google_calendar():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def fetch_events(service):
    now = datetime.datetime.now(datetime.timezone.utc)
    time_min = (now - datetime.timedelta(days=30)).isoformat()
    events_result = service.events().list(
        calendarId='primary',
        timeMin=time_min,
        maxResults=250,
        singleEvents=True,
        orderBy='startTime'
    ).execute()
    return events_result.get('items', [])

def group_events_by_date(events):
    grouped = defaultdict(list)
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        date_only = start[:10]
        grouped[date_only].append(event)
    return grouped

def save_events_by_day(grouped_events):
    os.makedirs(EXPORT_FOLDER, exist_ok=True)
    for date, events in grouped_events.items():
        filename = os.path.join(EXPORT_FOLDER, f'{date}_events.csv')
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Summary', 'Start', 'End', 'Description'])
            for event in events:
                summary = event.get('summary', '')
                start = event['start'].get('dateTime', event['start'].get('date'))
                end = event['end'].get('dateTime', event['end'].get('date'))
                description = event.get('description', '')
                writer.writerow([summary, start, end, description])
        print(f"{filename} saved with {len(events)} event(s).")

def main():
    creds = authenticate_google_calendar()
    service = build('calendar', 'v3', credentials=creds)
    events = fetch_events(service)

    if not events:
        print("No events found.")
        return

    print(f'{len(events)} events fetched.')
    grouped = group_events_by_date(events)
    save_events_by_day(grouped)

if __name__ == '__main__':
    main()
