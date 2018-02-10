from __future__ import print_function
import httplib2
import os
import pytz

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from datetime import datetime, date, time

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Calendar API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

    
def main(dataSum, dataDesc, dataLoc, dataStart, dataEnd, dataEndRec, dataID):
    
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    event = {
        "kind" : "calendar#event",
        "summary" : dataSum,                # CHANGE THE SUMMARY HERE
        "description" : dataDesc,           # CHANGE THE DESCRIPTION HERE
        "location" : dataLoc,               # CHANGE THE LOCATION HERE
        "start" : {
            "dateTime" : dataStart,
            "timeZone" : "Australia/Melbourne"
            },
        "end" : {
            "dateTime" : dataEnd,
            "timeZone" : "Australia/Melbourne"
            },
        "recurrence": [
            "RRULE:FREQ=WEEKLY;UNTIL=" + dataEndRec + "T240000Z",
            ]
        }
            
    event = service.events().insert(calendarId=dataID, body=event).execute()
                             
if __name__ == '__main__':
    print("Start Improved Scraper.py instead")

