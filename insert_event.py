from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from import_my_xlsx import ImportMyXlsx


# シフト表のExcelファイルから自分の予定をimport
return_list = ImportMyXlsx()
y_m = return_list[0]
working_list = return_list[1]

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']

def main():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    for i in range(len(working_list)):
        index = str(i + 1)
        day = working_list[i][0]
        start = working_list[i][1]
        end = working_list[i][2]
        event = {
            'summary': '出勤' + index,
            'location': 'Tokyo',
            'description': 'sample',
            'start': {
                'dateTime': y_m[0]+'-'+y_m[1]+'-'+day+'T'+start+':00',
                'timeZone': 'Japan',
            },
            'end': {
                'dateTime': y_m[0]+'-'+y_m[1]+'-'+day+'T'+end+':00',
                'timeZone': 'Japan',
            },
        }

        event = service.events().insert(calendarId='your calendarId',
                                        body=event).execute()
        print (event['id'])


if __name__ == '__main__':
    main()