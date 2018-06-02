from __future__ import print_function
import os, time, datetime
import webapp2
from google.auth import app_engine
from apiclient.discovery import build

class MainPage(webapp2.RequestHandler):
    def get(self):
        #Google API Setup
        credentials = app_engine.Credentials()
        service = build('sheets', 'v4', credentials=credentials)
        spreadsheet_id = '1Mlw-vHaiMcAN7OJZpFVzw0vd1lbv05zC_4mgzhUP64Q'; #Replace this
        
        #Collect list of users for next month 
        RANGE_NAME = 'Staff List!A2:11'
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,range=RANGE_NAME).execute()
        values = result.get('values', [])
          
        users = []
        for row in values:
            users.append(str(row[0]))
        
        #Sheet Name Details
        today = datetime.datetime.now()
        month = today.month
        year = today.year
        
        if (month == 12):
            nextMonth = 1
            nextYear = year + 1
        else:
            nextMonth = month + 1
            nextYear = year
        
        if (month == 1):
            lastMonth = 12
            lastYear = year - 1
        else:
            lastMonth = month - 1
            lastYear = year
        
        nextCreation = datetime.date(nextYear, nextMonth, 1)
        lastCreation = datetime.date(lastYear, lastMonth, 1)
        nextSheetName = nextCreation.strftime("%B") + ' ' + str(nextYear)
        lastSheetName = lastCreation.strftime("%B") + ' ' + str(lastYear)
        
        
        #Generate all Mondays in the next month
        mondays = []
        while (nextCreation.month == nextMonth):
            if nextCreation.weekday() == 0:
                mondays.append(nextCreation.day)
            nextCreation = nextCreation + datetime.timedelta(days=1)
        
        #Get Last month's Sheet ID to hide it
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets')
        for sheet in sheets:
            properties = sheet.get('properties')
            if (properties.get('title') == lastSheetName):
                lastSheetId = properties.get('sheetId')
        
        #Create Requests object
        #Two objectives for this request, hide last month's sheet and create next month's sheet
        #Example: Program runs in June
        #Hides: May. Creates: July
        requests = []
        
        #Add a new Sheet
        requests.append({
            "addSheet": {
            "properties": {
              "title": nextSheetName,
              "gridProperties": {
                "rowCount": len(mondays) * (4 + len(users)) + 8,
                "columnCount": 7
              }
            }
          }
        })
        
        #Hide last month's sheet
        requests.append({
                "updateSheetProperties": {
                    "properties":{
                        "sheetId": lastSheetId,
                        "hidden": 'true',
                    },
                    "fields": 'hidden'
                }
            })
        
        body = {
            'requests': requests
        }
        
        #Send the Request to the server
        response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,body=body).execute()
        
        #Get the Id of the new sheet, we need this to put all the information in later 
        newSheetId = response['replies'][0]['addSheet']['properties']['sheetId']
        
        #Empty the Requests, we need to use it again
        #Objective: Populate next month's sheet with information
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
              "endIndex": 7
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
                  "endColumnIndex": 7
                },
                "destination": {
                  "sheetId": newSheetId,
                  "startRowIndex": i * (4+ len(users)),
                  "endRowIndex": i * (4+ len(users)) + 4,
                  "startColumnIndex": 0,
                  "endColumnIndex": 7
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
                      "endColumnIndex": 7
                    },
                    "destination": {
                      "sheetId": newSheetId,
                      "startRowIndex": i * (4+ len(users)) + 4 + j,
                      "endRowIndex": i * (4+ len(users)) + 4 + j + 1,
                      "startColumnIndex": 0,
                      "endColumnIndex": 7
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
                      "startRowIndex": 5,
                      "endRowIndex": 13,
                      "startColumnIndex": 0,
                      "endColumnIndex": 7
                    },
                    "destination": {
                      "sheetId": newSheetId,
                      "startRowIndex": len(mondays) * (4+ len(users)),
                      "endRowIndex": len(mondays) * (4+ len(users)) + 8,
                      "startColumnIndex": 0,
                      "endColumnIndex": 7
                    },
                    "pasteType": "PASTE_NORMAL",
                  }
                })
        
        #Moving of sheet to front
        
        requests.append({
                "updateSheetProperties": {
                    "properties":{
                        "sheetId": newSheetId,
                        "index": 0,
                        "gridProperties": {
                            "hideGridlines": 'true'
                        }
                    },
                    "fields": 'index, gridProperties.hideGridlines'
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
              "range": nextSheetName+"!C" + str(i * (4+ len(users)) + 1),
              "majorDimension": "ROWS",
              "values": [
                [str(nextMonth) + "/" + str(mondays[i]) + "/" + str(nextYear)]
              ],
            }])
            for j in xrange(len(users)):
                requests.append([{
                  "range": nextSheetName+"!A" + str(i * (4+ len(users)) + 5 + j) + ":B" + str(i * (4+ len(users)) + 5 + j) ,
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

app = webapp2.WSGIApplication([
    ('/', MainPage),
], debug=True)