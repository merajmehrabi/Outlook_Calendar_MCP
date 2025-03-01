' findFreeSlots.vbs - Finds available time slots in the calendar
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Constants for working hours
Const DEFAULT_WORK_DAY_START = 9 ' 9 AM
Const DEFAULT_WORK_DAY_END = 17 ' 5 PM
Const DEFAULT_SLOT_DURATION = 30 ' 30 minutes

' Main function
Sub Main()
    ' Get command line arguments
    Dim startDateStr, endDateStr, durationStr, workDayStartStr, workDayEndStr, calendarName
    Dim startDate, endDate, duration, workDayStart, workDayEnd
    
    ' Get and validate arguments
    startDateStr = GetArgument("startDate")
    endDateStr = GetArgument("endDate")
    durationStr = GetArgument("duration")
    workDayStartStr = GetArgument("workDayStart")
    workDayEndStr = GetArgument("workDayEnd")
    calendarName = GetArgument("calendar")
    
    ' Require start date
    RequireArgument "startDate"
    
    ' Parse dates
    startDate = ParseDate(startDateStr)
    
    ' If end date is not provided, use 7 days from start date
    If endDateStr = "" Then
        endDate = DateAdd("d", 7, startDate)
    Else
        endDate = ParseDate(endDateStr)
    End If
    
    ' Ensure end date is not before start date
    If endDate < startDate Then
        OutputError "End date cannot be before start date"
        WScript.Quit 1
    End If
    
    ' Parse duration (in minutes)
    If durationStr = "" Then
        duration = DEFAULT_SLOT_DURATION
    Else
        If Not IsNumeric(durationStr) Then
            OutputError "Duration must be a number (minutes)"
            WScript.Quit 1
        End If
        duration = CInt(durationStr)
    End If
    
    ' Parse work day start/end hours
    If workDayStartStr = "" Then
        workDayStart = DEFAULT_WORK_DAY_START
    Else
        If Not IsNumeric(workDayStartStr) Then
            OutputError "Work day start hour must be a number (0-23)"
            WScript.Quit 1
        End If
        workDayStart = CInt(workDayStartStr)
        If workDayStart < 0 Or workDayStart > 23 Then
            OutputError "Work day start hour must be between 0 and 23"
            WScript.Quit 1
        End If
    End If
    
    If workDayEndStr = "" Then
        workDayEnd = DEFAULT_WORK_DAY_END
    Else
        If Not IsNumeric(workDayEndStr) Then
            OutputError "Work day end hour must be a number (0-23)"
            WScript.Quit 1
        End If
        workDayEnd = CInt(workDayEndStr)
        If workDayEnd < 0 Or workDayEnd > 23 Then
            OutputError "Work day end hour must be between 0 and 23"
            WScript.Quit 1
        End If
    End If
    
    ' Ensure work day end is after work day start
    If workDayEnd <= workDayStart Then
        OutputError "Work day end hour must be after work day start hour"
        WScript.Quit 1
    End If
    
    ' Find free slots
    Dim freeSlots
    freeSlots = FindFreeTimeSlots(startDate, endDate, duration, workDayStart, workDayEnd, calendarName)
    
    ' Output free slots as JSON
    OutputSuccess freeSlots
End Sub

' Finds free time slots in the calendar
Function FindFreeTimeSlots(startDate, endDate, duration, workDayStart, workDayEnd, calendarName)
    On Error Resume Next
    
    ' Create Outlook objects
    Dim outlookApp, calendar, events, busySlots, freeSlots
    Dim currentDate, currentSlotStart, currentSlotEnd, i, event, isSlotFree
    
    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()
    
    ' Get calendar folder
    If calendarName = "" Then
        Set calendar = GetDefaultCalendar(outlookApp)
    Else
        Set calendar = GetCalendarByName(outlookApp, calendarName)
    End If
    
    ' Get all events in the date range
    Set events = GetCalendarEvents(startDate, endDate, calendarName)
    
    ' Create an array to store busy time slots
    busySlots = CreateBusySlots(events)
    
    ' Create an array to store free time slots
    freeSlots = CreateObject("System.Collections.ArrayList")
    
    ' Loop through each day in the date range
    currentDate = startDate
    Do While currentDate <= endDate
        ' Skip weekends (Saturday = 7, Sunday = 1)
        If Weekday(currentDate) <> 1 And Weekday(currentDate) <> 7 Then
            ' Loop through each potential slot in the work day
            currentSlotStart = DateAdd("h", workDayStart, currentDate)
            
            Do While DateAdd("n", duration, currentSlotStart) <= DateAdd("h", workDayEnd, currentDate)
                currentSlotEnd = DateAdd("n", duration, currentSlotStart)
                
                ' Check if the slot is free
                isSlotFree = True
                
                For i = 0 To UBound(busySlots, 2)
                    ' If the slot overlaps with a busy slot, it's not free
                    If (currentSlotStart < busySlots(1, i)) And (currentSlotEnd > busySlots(0, i)) Then
                        isSlotFree = False
                        Exit For
                    End If
                Next
                
                ' If the slot is free, add it to the free slots array
                If isSlotFree Then
                    freeSlots.Add Array(FormatDateTime(currentSlotStart), FormatDateTime(currentSlotEnd))
                End If
                
                ' Move to the next slot
                currentSlotStart = DateAdd("n", 30, currentSlotStart) ' 30-minute increments
            Loop
        End If
        
        ' Move to the next day
        currentDate = DateAdd("d", 1, currentDate)
    Loop
    
    ' Convert free slots to JSON
    Dim json, slot
    
    json = "["
    
    For i = 0 To freeSlots.Count - 1
        If i > 0 Then json = json & ","
        
        slot = freeSlots(i)
        json = json & "{""start"":""" & slot(0) & """,""end"":""" & slot(1) & """}"
    Next
    
    json = json & "]"
    
    FindFreeTimeSlots = json
    
    ' Clean up
    Set calendar = Nothing
    Set outlookApp = Nothing
End Function

' Creates an array of busy time slots from calendar events
Function CreateBusySlots(events)
    Dim i, busySlots, event, count
    
    ' Count the number of events
    count = events.Count
    
    ' Create a 2D array to store busy slots (start and end times)
    ReDim busySlots(1, count - 1)
    
    ' Fill the array with event start and end times
    For i = 1 To count
        Set event = events.Item(i)
        
        ' Only consider events marked as Busy or Out of Office
        If event.BusyStatus = olBusy Or event.BusyStatus = olOutOfOffice Then
            busySlots(0, i - 1) = event.Start
            busySlots(1, i - 1) = event.End
        End If
    Next
    
    CreateBusySlots = busySlots
End Function

' Run the main function
Main
