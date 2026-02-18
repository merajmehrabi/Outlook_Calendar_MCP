' listEvents.vbs - Lists calendar events within a specified date range
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Main function
Sub Main()
    ' Get command line arguments
    Dim startDateStr, endDateStr, calendarName
    Dim startDate, endDate
    
    ' Get and validate arguments
    startDateStr = GetArgument("startDate")
    endDateStr = GetArgument("endDate")
    calendarName = GetArgument("calendar")
    
    ' Require start date
    RequireArgument "startDate"
    
    ' Parse dates
    startDate = ParseDate(startDateStr)
    
    ' If end date is not provided, use start date (single day)
    If endDateStr = "" Then
        endDate = startDate
    Else
        endDate = ParseDate(endDateStr)
    End If
    
    ' Ensure end date is not before start date
    If endDate < startDate Then
        OutputError "End date cannot be before start date"
        WScript.Quit 1
    End If
    
    ' Get calendar events
    Dim events
    Set events = GetCalendarEvents(startDate, endDate, calendarName)
    
    ' Output events as JSON
    OutputSuccess AppointmentsToJSON(events)
End Sub

' Gets calendar events within the specified date range
Function GetCalendarEvents(startDate, endDate, calendarName)
    On Error Resume Next
    
    ' Create Outlook objects
    Dim outlookApp, calendar, filter, events, allItems
    
    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()
    
    ' Get calendar folder
    If calendarName = "" Then
        Set calendar = GetDefaultCalendar(outlookApp)
    Else
        Set calendar = GetCalendarByName(outlookApp, calendarName)
    End If
    
    ' CRITICAL: Must be in this exact order for recurring events:
    ' 1. Get Items collection
    ' 2. Sort by [Start]
    ' 3. Set IncludeRecurrences = True
    ' 4. Then filter with Restrict
    Set allItems = calendar.Items
    allItems.Sort "[Start]"
    allItems.IncludeRecurrences = True
    
    ' Create filter for date range
    filter = "[Start] >= '" & FormatDate(startDate) & " 12:00 AM' AND [Start] <= '" & FormatDate(DateAdd("d", 1, endDate)) & " 12:00 AM'"
    
    ' Get events matching the filter
    Set events = allItems.Restrict(filter)
    
    If Err.Number <> 0 Then
        OutputError "Failed to get calendar events: " & Err.Description
        WScript.Quit 1
    End If
    
    ' Return events
    Set GetCalendarEvents = events
    
    ' Clean up
    Set calendar = Nothing
    Set outlookApp = Nothing
End Function

' Run the main function
Main
