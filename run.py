
from __future__ import print_function
import httplib2
import os
import time
import datetime
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    credential_dir = os.path.join('./.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

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

def main():
    
    today = datetime.datetime.now()
    month = today.month
    year = today.year
    
    if (month == 12):
        month = 1
        year += 1
    else:
        month += 1
    
    creation = datetime.date(year, month, 1)
    sheetName = creation.strftime("%B") + ' ' + str(year)
    #Generate all Mondays in the next month
    mondays = []
    while (creation.month == month):
        if creation.weekday() == 0:
            mondays.append(creation.day)
        creation = creation + datetime.timedelta(days=1)
    

    #Generate List of Users
    file = open('users.txt')
    users = file.readlines()
    for i in range(len(users)):
        users[i] = users[i].rstrip()
        
    
    #Google API Setup
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,discoveryServiceUrl=discoveryUrl)

    spreadsheet_id = '1ETDw0OFi0Xlo6h_QibkeBocYUvY8vVYSMJsAVElCQ40';
	
    
    
    requests = []
    
    #Add a new Sheet
    requests.append({
		"addSheet": {
        "properties": {
          "title": sheetName,
          "gridProperties": {
            "rowCount": len(mondays) * (4 + len(users)) + 8,
            "columnCount": 12
          }
        }
      }
    })
	
    
    
    body = {
		'requests': requests
	}
	
	#Send the Request to the server
    response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,body=body).execute()
    
    #Get the Id of the new sheet, we need this for later 
    newSheetId = response['replies'][0]['addSheet']['properties']['sheetId']

    #Empty the Requests, we need to use 
    requests = []
    
    #Set Column Sizes    
    requests.append([{"updateDimensionProperties": {
        "range": {
          "sheetId": newSheetId,
          "dimension": "COLUMNS",
          "startIndex": 0,
          "endIndex": 1
        },
        "properties": {
          "pixelSize": 20
        },
        "fields": "pixelSize"
      }
    },
    {
        "updateDimensionProperties": {
        "range": {
          "sheetId": newSheetId,
          "dimension": "COLUMNS",
          "startIndex": 2,
          "endIndex": 12
        },
        "properties": {
          "pixelSize": 159
        },
        "fields": "pixelSize"
      }
    },
    {
        "updateDimensionProperties": {
        "range": {
          "sheetId": newSheetId,
          "dimension": "COLUMNS",
          "startIndex": 1,
          "endIndex": 2
        },
        "properties": {
          "pixelSize": 166
        },
        "fields": "pixelSize"
      }
    }]
    )
    
    #Add the Header
    for i in range(len(mondays)):
        requests.append({
          "copyPaste": {
            "source": {
              "sheetId": 1370351810,
              "startRowIndex": 0,
              "endRowIndex": 4,
              "startColumnIndex": 0,
              "endColumnIndex": 12
            },
            "destination": {
              "sheetId": newSheetId,
              "startRowIndex": i * (4+ len(users)),
              "endRowIndex": i * (4+ len(users)) + 4,
              "startColumnIndex": 0,
              "endColumnIndex": 12
            },
            "pasteType": "PASTE_NORMAL",
          }
        })
        
        #Add Row Formatting
        requests.append([{"updateDimensionProperties": {
                #Green Header Row
                "range": {
                  "sheetId": newSheetId,
                  "dimension": "ROWS",
                  "startIndex": i * (4+ len(users)),
                  "endIndex": i * (4+ len(users)) + 1
                },
                "properties": {
                  "pixelSize": 8
                },
                "fields": "pixelSize"
              }
            },
            {
                "updateDimensionProperties": {
                "range": {
                  "sheetId": newSheetId,
                  "dimension": "ROWS",
                  "startIndex": i * (4+ len(users)) + 1,
                  "endIndex": i * (4+ len(users)) + 4 + len(users)
                },
                "properties": {
                  "pixelSize": 30
                },
                "fields": "pixelSize"
              }
            },
            {"updateDimensionProperties": {
                "range": {
                  "sheetId": newSheetId,
                  "dimension": "ROWS",
                  "startIndex": 1,
                  "endIndex": 2
                },
                "properties": {
                  "pixelSize": 48
                },
                "fields": "pixelSize"
              }
            },
            {"updateDimensionProperties": {
                "range": {
                  "sheetId": newSheetId,
                  "dimension": "ROWS",
                  "startIndex": len(mondays) * (4+ len(users)),
                  "endIndex": len(mondays) * (4+ len(users)) + 8,
                },
                "properties": {
                  "pixelSize": 30
                },
                "fields": "pixelSize"
              }
            }
        ])
        
        #Add the User Rows
        for j in range(len(users)):
            requests.append([{
              "copyPaste": {
                "source": {
                  "sheetId": 1370351810,
                  "startRowIndex": 4,
                  "endRowIndex": 5,
                  "startColumnIndex": 0,
                  "endColumnIndex": 12
                },
                "destination": {
                  "sheetId": newSheetId,
                  "startRowIndex": i * (4+ len(users)) + 4 + j,
                  "endRowIndex": i * (4+ len(users)) + 4 + j + 1,
                  "startColumnIndex": 0,
                  "endColumnIndex": 12
                },
                "pasteType": "PASTE_NORMAL",
              }
            }
            ])
    
    #Addition of Comments / Legend
    requests.append({
              "copyPaste": {
                "source": {
                  "sheetId": 1370351810,
                  "startRowIndex": 28,
                  "endRowIndex": 36,
                  "startColumnIndex": 0,
                  "endColumnIndex": 12
                },
                "destination": {
                  "sheetId": newSheetId,
                  "startRowIndex": len(mondays) * (4+ len(users)),
                  "endRowIndex": len(mondays) * (4+ len(users)) + 8,
                  "startColumnIndex": 0,
                  "endColumnIndex": 12
                },
                "pasteType": "PASTE_NORMAL",
              }
            })
    body = {
		'requests': requests
	}
    
    response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,body=body).execute()
    
    #Filling in Date and Names
    requests = []
    
    for i in xrange(len(mondays)):
        requests.append([{
          "range": sheetName+"!C" + str(i * (4+ len(users)) + 1),
          "majorDimension": "ROWS",
          "values": [
            [str(month) + "/" + str(mondays[i]) + "/" + str(year)]
          ],
        }])
        for j in xrange(len(users)):
            requests.append([{
              "range": sheetName+"!A" + str(i * (4+ len(users)) + 5 + j) + ":B" + str(i * (4+ len(users)) + 5 + j) ,
              "majorDimension": "ROWS",
              "values": [
                [str(j+1), users[j]]
              ],
            }])
    
    body = {
        "valueInputOption": "USER_ENTERED",
		'data': requests
        }
    
    response = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id,body=body).execute()
    
if __name__ == '__main__':
    main()