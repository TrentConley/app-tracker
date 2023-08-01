import calendar
import datetime
from datetime import timedelta
from AppKit import NSWorkspace
import argparse
from oauth2client.service_account import ServiceAccountCredentials
import httplib2
from googleapiclient import discovery
from pprint import pprint
import AppKit
import WebKit
import subprocess

import pygsheets


def update_app_stats(app_stats, start_time, end_time, duration, previous_app_name):
  if previous_app_name not in app_stats:
    app_stats[previous_app_name] = {
      'duration': duration,
      'periods': [(start_time, end_time)]
    }
  else:
    app_stats[previous_app_name]['duration'] += duration
    app_stats[previous_app_name]['periods'].append((start_time, end_time))
  
  return app_stats

def update_sheet(wks, app_stats):
  data = []
  for app, info in app_stats.items():
    for period in info['periods']:
      data.append([app, info['duration'], period[0].isoformat(), period[1].isoformat()]) 

  wks.insert_rows(row=1, values=data)
  print("updated sheet")




def run():
  gc = pygsheets.authorize(service_file='calendar-automation-393821-7b81b3d907b7.json')

  spreadsheet_id = "1LwWx8zP9HOHGJ-nXEK1DmYWYhMikn5CWdU_ZY486Kko"
  sheet = gc.open_by_key(spreadsheet_id)

  wks = sheet.worksheet_by_title('Sheet1')

  app_stats = {}

  print("got worksheet")

  # Set up Google Calendar API credentials
  SCOPES = ['https://www.googleapis.com/auth/calendar']
  SERVICE_ACCOUNT_FILE = 'calendar-automation-393821-7b81b3d907b7.json'

  credentials = ServiceAccountCredentials.from_json_keyfile_name(
          SERVICE_ACCOUNT_FILE, SCOPES)

  http = credentials.authorize(httplib2.Http())
  service = discovery.build('calendar', 'v3', http=http)

  # Get current datetime
  now = datetime.datetime.now() 

  # Get list of running applications
  workspace = NSWorkspace.sharedWorkspace()
  apps = workspace.runningApplications()
  app_names = [app.localizedName() for app in apps]

  import time

  # Get the name of the current application
  previous_app_name = AppKit.NSWorkspace.sharedWorkspace().activeApplication()['NSApplicationName']
  start_time = datetime.datetime.now()

  while True:
    print('Loop')
    new_app_name = AppKit.NSWorkspace.sharedWorkspace().activeApplication()['NSApplicationName']
    print(new_app_name)

    if new_app_name != previous_app_name:
      print('Using new app')

      end_time = datetime.datetime.now()
      duration = (end_time - start_time).total_seconds() / 60.0  # duration in minutes
      app_stats = update_app_stats(app_stats, start_time, end_time, duration, previous_app_name)
      update_sheet(wks, app_stats)

      if duration >= 1:  # only create event if duration is at least 1 minute
        # Create calendar event for the previous application
        event = {
            'summary': previous_app_name,
            'start': {
                'dateTime': start_time.isoformat(),
                'timeZone': 'America/Los_Angeles',
            },
            'end': {
                'dateTime': end_time.isoformat(),
                'timeZone': 'America/Los_Angeles',
            }
        }
        # Add event to calendar
        event = service.events().insert(calendarId='0518785f3057575ea04ecf6eab8c4e9f7af05fd2a3ec92b07ffb79981ed2662b@group.calendar.google.com', body=event).execute()
        print('Event created: %s' % (event.get('htmlLink')))
      # Update current application and start time
      previous_app_name = new_app_name
      start_time = datetime.datetime.now()
    time.sleep(5)  # check every minute
  # Add event to calendar
  event = service.events().insert(calendarId='0518785f3057575ea04ecf6eab8c4e9f7af05fd2a3ec92b07ffb79981ed2662b@group.calendar.google.com', body=event).execute()

  print('Event created: %s' % (event.get('htmlLink')))




run()