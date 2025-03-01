' createEvent.vbs - Creates a new calendar event
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Main function
Sub Main()
    ' Get command line arguments
    Dim subject, startDateStr, startTimeStr, endDateStr, endTimeStr, location, body, isMeeting, attendeesStr, calendarName
    Dim startDateTime, endDateTime, attendees
    
    ' Get and validate arguments
    subject = GetArgument("subject")
    startDateStr = GetArgument("startDate")
    startTimeStr = GetArgument("startTime")
    endDateStr = GetArgument("endDate")
    endTimeStr = GetArgument("endTime")
    location = GetArgument("location")
    body = GetArgument("body")
    isMeeting = LCase(GetArgument("isMeeting")) = "true"
    attendeesStr = GetArgument("attendees")
    calendarName = GetArgument("calendar")
    
    ' Require subject and start date/time
    RequireArgument "subject"
    RequireArgument "startDate"
    RequireArgument "startTime"
    
    ' Parse start date/time
    startDateTime = ParseDateTime(startDateStr, startTimeStr)
    
    ' Parse end date/time (if not provided, default to 30 minutes after start)
    If endDateStr = "" Then endDateStr = startDateStr
    If endTimeStr = "" Then
        endDateTime = DateAdd("n", 30, startDateTime)
    Else
        endDateTime = ParseDateTime(endDateStr, endTimeStr)
    End If
    
    ' Ensure end time is not before start time
    If endDateTime <= startDateTime Then
        OutputError "End time cannot be before or equal to start time"
        WScript.Quit 1
    End If
    
    ' Parse attendees (if provided and it's a meeting)
    If isMeeting And attendeesStr <> "" Then
        attendees = Split(attendeesStr, ";")
    Else
        attendees = Array()
    End If
    
    ' Create the event
    Dim eventId
    eventId = CreateCalendarEvent(subject, startDateTime, endDateTime, location, body, isMeeting, attendees, calendarName)
    
    ' Output success with the event ID
    OutputSuccess "{""eventId"":""" & eventId & """}"
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

' Creates a new calendar event with the specified properties
Function CreateCalendarEvent(subject, startDateTime, endDateTime, location, body, isMeeting, attendees, calendarName)
    On Error Resume Next
    
    ' Create Outlook objects
    Dim outlookApp, calendar, appointment, i, recipient
    
    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()
    
    ' Get calendar folder
    If calendarName = "" Then
        Set calendar = GetDefaultCalendar(outlookApp)
    Else
        Set calendar = GetCalendarByName(outlookApp, calendarName)
    End If
    
    ' Create new appointment item
    Set appointment = calendar.Items.Add(olAppointmentItem)
    
    ' Set appointment properties
    appointment.Subject = subject
    appointment.Start = startDateTime
    appointment.End = endDateTime
    appointment.Location = location
    appointment.Body = body
    
    ' If it's a meeting, add attendees
    If isMeeting Then
        appointment.MeetingStatus = olMeeting
        
        ' Add attendees
        For i = LBound(attendees) To UBound(attendees)
            If Trim(attendees(i)) <> "" Then
                Set recipient = appointment.Recipients.Add(Trim(attendees(i)))
                recipient.Type = 1 ' Required attendee
            End If
        Next
        
        ' Send the meeting request
        appointment.Send
    Else
        ' Save the appointment
        appointment.Save
    End If
    
    If Err.Number <> 0 Then
        OutputError "Failed to create calendar event: " & Err.Description
        WScript.Quit 1
    End If
    
    ' Return the EntryID as the event ID
    CreateCalendarEvent = appointment.EntryID
    
    ' Clean up
    Set appointment = Nothing
    Set calendar = Nothing
    Set outlookApp = Nothing
End Function

' Run the main function
Main
