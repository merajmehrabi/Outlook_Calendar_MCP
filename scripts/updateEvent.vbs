' updateEvent.vbs - Updates an existing calendar event
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Main function
Sub Main()
    ' Get command line arguments
    Dim eventId, subject, startDateStr, startTimeStr, endDateStr, endTimeStr, location, body, calendarName
    Dim startDateTime, endDateTime
    
    ' Get and validate arguments
    eventId = GetArgument("eventId")
    subject = GetArgument("subject")
    startDateStr = GetArgument("startDate")
    startTimeStr = GetArgument("startTime")
    endDateStr = GetArgument("endDate")
    endTimeStr = GetArgument("endTime")
    location = GetArgument("location")
    body = GetArgument("body")
    calendarName = GetArgument("calendar")
    
    ' Require event ID
    RequireArgument "eventId"
    
    ' Parse date/time if provided
    If startDateStr <> "" And startTimeStr <> "" Then
        startDateTime = ParseDateTime(startDateStr, startTimeStr)
    End If
    
    If endDateStr <> "" And endTimeStr <> "" Then
        endDateTime = ParseDateTime(endDateStr, endTimeStr)
    ElseIf startDateTime <> Empty And endDateStr = "" And endTimeStr = "" Then
        ' If only start time is provided, default end time to 30 minutes after start
        endDateTime = DateAdd("n", 30, startDateTime)
    End If
    
    ' Ensure end time is not before start time if both are provided
    If Not IsEmpty(startDateTime) And Not IsEmpty(endDateTime) Then
        If endDateTime <= startDateTime Then
            OutputError "End time cannot be before or equal to start time"
            WScript.Quit 1
        End If
    End If
    
    ' Update the event
    Dim result
    result = UpdateCalendarEvent(eventId, subject, startDateTime, endDateTime, location, body, calendarName)
    
    ' Output success
    OutputSuccess "{""success"":" & LCase(CStr(result)) & "}"
End Sub

' Parses a date and time string into a DateTime object
Function ParseDateTime(dateStr, timeStr)
    Dim dateObj, timeObj, dateTimeStr
    
    ' Parse date
    dateObj = ParseDate(dateStr)
    
    ' Combine date and time
    dateTimeStr = FormatDate(dateObj) & " " & timeStr
    
    ' Parse combined date/time
    If Not IsDate(dateTimeStr) Then
        OutputError "Invalid time format: " & timeStr
        WScript.Quit 1
    End If
    
    ParseDateTime = CDate(dateTimeStr)
End Function

' Updates an existing calendar event with the specified properties
Function UpdateCalendarEvent(eventId, subject, startDateTime, endDateTime, location, body, calendarName)
    On Error Resume Next
    
    ' Create Outlook objects
    Dim outlookApp, namespace, calendar, appointment
    
    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()
    
    ' Get MAPI namespace
    Set namespace = outlookApp.GetNamespace("MAPI")
    
    ' Get calendar folder
    If calendarName = "" Then
        Set calendar = GetDefaultCalendar(outlookApp)
    Else
        Set calendar = GetCalendarByName(outlookApp, calendarName)
    End If
    
    ' Try to get the appointment by EntryID
    Set appointment = namespace.GetItemFromID(eventId)
    
    ' Check if appointment was found
    If appointment Is Nothing Then
        OutputError "Event not found with ID: " & eventId
        UpdateCalendarEvent = False
        Exit Function
    End If
    
    ' Update appointment properties if provided
    If subject <> "" Then appointment.Subject = subject
    If Not IsEmpty(startDateTime) Then appointment.Start = startDateTime
    If Not IsEmpty(endDateTime) Then appointment.End = endDateTime
    If location <> "" Then appointment.Location = location
    If body <> "" Then appointment.Body = body
    
    ' Save the changes
    appointment.Save
    
    If Err.Number <> 0 Then
        OutputError "Failed to update calendar event: " & Err.Description
        UpdateCalendarEvent = False
    Else
        UpdateCalendarEvent = True
    End If
    
    ' Clean up
    Set appointment = Nothing
    Set calendar = Nothing
    Set namespace = Nothing
    Set outlookApp = Nothing
End Function

' Run the main function
Main
