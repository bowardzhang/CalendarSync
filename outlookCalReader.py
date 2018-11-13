#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
    Read outlook calendar
    Author: Bo Zhang <bowardzhang@gmail.com>
    Created on 13.07.2018
"""

from time import time
import win32com.client, datetime

RecurrenceTypeDict = {0:"DAILY", 1:"WEEKLY", 2:"MONTHLY",}

def getOutlookCalEvents(dayMax = 360, eventMax = 100):
    Outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = Outlook.GetNamespace("MAPI")

    events = namespace.GetDefaultFolder(9).Items
    
    events.Sort("[Start]")
    events.IncludeRecurrences = "False"
    
    begin = datetime.date.today()
    end = begin + datetime.timedelta(days = dayMax)
    restriction = "[Start] >= '" + begin.strftime("%d/%m/%Y") + "' AND [End] <= '" +end.strftime("%d/%m/%Y") + "'"
    
    restrictedEvents = list(events.Restrict(restriction))
    
    eventCount = len(restrictedEvents)
    print("found %d events in Outlook" % eventCount)
    
    if eventCount > eventMax:
        restrictedEvents = restrictedEvents[:eventMax]
        print("too many events. Only read the first %d" % eventMax)
    return restrictedEvents

if __name__ == "__main__":
    tstart = time()

    # Iterate through restricted AppointmentItems and print them
    for event in getOutlookCalEvents():
        print("\n%s \nLocation: %s \nStart: %s \nStartTimeZone: %s \nEnd: %s \nEndTimeZone: %s \nOrganizer: %s" % (
              event.Subject, event.Location, event.Start, 
              event.StartTimeZone, event.End, event.EndTimeZone, event.Organizer))
        if event.AllDayEvent:
            print("All Day Event\n")
            print(event.GlobalAppointmentID.lower())
            
        if event.IsRecurring:
            freq = RecurrenceTypeDict[event.GetRecurrencePattern().Interval]
            recEndDate = event.GetRecurrencePattern().PatternEndDate
            recEndDate = str(recEndDate).replace('-', '').replace(' ', 'T').replace(':', '')[:-7]
            rule = "PRULE:FREQ=%s;UNTIL=%sZ" % (freq, recEndDate)
            print(rule)
            
            # print(event.GetRecurrencePattern().RecurrenceType)
            # print("interval: %s" % event.GetRecurrencePattern().Interval)
            # print("instance: %s" % event.GetRecurrencePattern().Instance)
            # print(event.GetRecurrencePattern().Occurrences)
            # print("%d exceptions: " % event.GetRecurrencePattern().Exceptions.count)
            exceptionIndexes = [i for i in range(event.GetRecurrencePattern().Exceptions.count) if event.GetRecurrencePattern().Exceptions.Item(i+1).Deleted]
            for i in range(event.GetRecurrencePattern().Exceptions.count):
                if event.GetRecurrencePattern().Exceptions.Item(i+1).Deleted:
                    print("exception: %s" % event.GetRecurrencePattern().Exceptions.Item(i+1))
                else:
                    instance = event.GetRecurrencePattern().Exceptions.Item(i+1).AppointmentItem
                    print("instance: %s at %s" % (instance.Start, instance.Location))
                    
            # print(event.GetRecurrencePattern().PatternStartDate)
            # print(event.GetRecurrencePattern().PatternEndDate)
            # print((event.RecurrenceState))
              
    print("Finished in %d seconds" % (time()-tstart))
