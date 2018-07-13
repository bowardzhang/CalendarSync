#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
    Synchronize MS Outlook calendar to Google calendar
    Author: Bo Zhang <bowardzhang@gmail.com>
    Created on 13.07.2018
"""

from time import time
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import datetime
import dateutil.parser
import outlookCalReader

HIDE_SUBJECT = True
ORGANIZATION_NAME = "ABC"

RecurrenceTypeDict = {0:"DAILY", 1:"WEEKLY", 2:"MONTHLY",}

# class
class GoogleCalendar:
    def __init__(self):
        # Setup the Calendar API
        SCOPES = 'https://www.googleapis.com/auth/calendar'
        store = file.Storage('credentials.json')
        creds = store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
            creds = tools.run_flow(flow, store)
        self.service = build('calendar', 'v3', http=creds.authorize(Http()))
        self.eventIds = []
        
    def readCalEvents(self, aheadDays=90, eventMax = 100):
        dtNow = datetime.datetime.now()
        dtMax = dtNow + datetime.timedelta(days=aheadDays)
        events_result = self.service.events().list(
            calendarId='primary',
            singleEvents=True,
            orderBy='startTime',
            maxResults=eventMax, 
            timeMin=dtNow.strftime('%Y-%m-%dT%H:%M:%S-00:00'),
            timeMax=dtMax.strftime('%Y-%m-%dT%H:%M:%S-00:00'),
            showDeleted=True,
            ).execute()
        events = events_result.get('items', [])
        print('found %d google events in %d days' % (len(events), aheadDays))
        
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            self.eventIds.append(event['id'])
        
    def eventExist(self, outlookEvent):
        if outlookEvent.GlobalAppointmentID.lower() in self.eventIds:
            return True
        return False
    
    def addOutlookCalEvent(self, outlookEvent):
        eId = outlookEvent.GlobalAppointmentID.lower()
        if HIDE_SUBJECT:
            summary = "%s Busy" % ORGANIZATION_NAME
        else:
            summary = outlookEvent.Subject
            
        location = outlookEvent.Location
        start = {}
        end = {}
        if outlookEvent.AllDayEvent:
            start['date'] = str(outlookEvent.Start)[:10]
            end['date'] = str(outlookEvent.End)[:10]
        else:
            start['dateTime'] = str(outlookEvent.Start).replace(' ', 'T')
            end['dateTime'] = str(outlookEvent.End).replace(' ', 'T')
            
            # fix incorrect timezone of outlook events
            if str(outlookEvent.StartTimeZone) == "W. Europe Standard Time":
                start['dateTime'] = start['dateTime'].replace("+00:00", "+02:00")
                start['timeZone'] = "Europe/Berlin"
            if str(outlookEvent.EndTimeZone) == "W. Europe Standard Time":
                end['dateTime'] = end['dateTime'].replace("+00:00", "+02:00")
                end['timeZone'] = "Europe/Berlin"
        
        event = {
              'id': eId,
              'summary': summary,
              'location': location,
              #'description': outlookEvent.Body,
              'start': start,
              'end': end,
        }
        
        oRecPattern = outlookEvent.GetRecurrencePattern()
        if outlookEvent.IsRecurring:
            freq = RecurrenceTypeDict[oRecPattern.Interval]
            recEndDate = oRecPattern.PatternEndDate
            recEndDate = str(recEndDate).replace('-', '').replace(' ', 'T').replace(':', '')[:-5]
            rule = "RRULE:FREQ=%s;UNTIL=%sZ" % (freq, recEndDate)
            event['recurrence'] = [rule]
            event['summary'] += " (series)"

        eventExists = True
        try:
            self.service.events().get(calendarId='primary', eventId=eId).execute()
        except:
            eventExists = False
        
        startStr = start.get('dateTime', start.get('date'))
        if eventExists:
            self.service.events().update(calendarId='primary', eventId=eId, body=event).execute()
            print('\nEvent [%s] on %s exists and is updated' % (outlookEvent.Subject, startStr))
        else:
            event = self.service.events().insert(calendarId='primary', body=event).execute()
            print('\nEvent %s on %s is created' % (summary, startStr))
        
        # update event instances in a recurrence series according to outlook
        if outlookEvent.IsRecurring:            
            gInstances = self.service.events().instances(calendarId='primary', eventId=eId, showDeleted=True).execute()['items']
            # must order the gInstances by time in order to keep the indexes consistent
            gInstances = sorted(gInstances, key=lambda x: x['start']['dateTime'])
            
            for i in range(oRecPattern.Exceptions.count):
                gInstance = gInstances[i]
                if oRecPattern.Exceptions.Item(i+1).Deleted: # deactivate exceptions
                    if gInstance['status'] != 'cancelled':
                        gInstance['status'] = 'cancelled'
                        self.service.events().update(calendarId='primary', eventId=gInstance['id'], body=gInstance).execute()
                        print("removed exception event on %s in a recurrence series" % str(oRecPattern.Exceptions.Item(i+1))[:-15])
                else: # update the location of non-exception instances
                    oInstance = oRecPattern.Exceptions.Item(i+1).AppointmentItem
                    gInstance['location'] = oInstance.Location
                    self.service.events().update(calendarId='primary', eventId=gInstance['id'], body=gInstance).execute()
                    print("updated instance on %s at %s" % (str(oInstance.Start)[:-15], oInstance.Location))
            
    def syncFromOutlook(self):
        outlookEvents = outlookCalReader.getOutlookCalEvents()
        for oe in outlookEvents:
            self.addOutlookCalEvent(oe)

if __name__ == "__main__":
    tstart = time()
    
    cal = GoogleCalendar()
    cal.readCalEvents()
    cal.syncFromOutlook()

    print("\nFinished in %d seconds" % (time()-tstart))
