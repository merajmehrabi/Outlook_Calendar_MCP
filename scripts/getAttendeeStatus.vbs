' getAttendeeStatus.vbs - Checks the response status of meeting attendees
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Main function
Sub Main()
    ' Get command line arguments
    Dim eventIdStr, calendarName
    
    ' Get and validate arguments
    eventIdStr = GetArgument("eventId")
    calendarName = GetArgument("calendar")
    
    ' Require event ID
    RequireArgument "eventId"
    
    ' Get attendee status
    Dim attendeeStatus
    attendeeStatus = GetAttendeeStatus(eventIdStr, calendarName)
    
    ' Output attendee status as JSON
    OutputSuccess attendeeStatus
End Sub

' Gets the response status of meeting attendees
Function GetAttendeeStatus(eventId, calendarName)
    On Error Resume Next
    
    ' Create Outlook objects
    Dim outlookApp, calendar, appointment, recipients, recipient, i
    Dim attendees, attendeeStatus, responseStatus
    
    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()
    
    ' Get calendar folder
    If calendarName = "" Then
        Set calendar = GetDefaultCalendar(outlookApp)
    Else
        Set calendar = GetCalendarByName(outlookApp, calendarName)
    End If
    
    ' Find the appointment by EntryID
    Set appointment = outlookApp.GetNamespace("MAPI").GetItemFromID(eventId)
    
    If Err.Number <> 0 Then
        OutputError "Failed to find event with ID: " & eventId & " - " & Err.Description
        WScript.Quit 1
    End If
    
    ' Check if it's a meeting
    If appointment.MeetingStatus <> olMeeting Then
        OutputError "The specified event is not a meeting"
        WScript.Quit 1
    End If
    
    ' Get recipients
    Set recipients = appointment.Recipients
    
    ' Create JSON array for attendees
    attendees = "["
    
    For i = 1 To recipients.Count
        Set recipient = recipients.Item(i)
        
        If i > 1 Then attendees = attendees & ","
        
        attendees = attendees & "{"
        attendees = attendees & """name"":""" & EscapeJSON(recipient.Name) & ""","
        attendees = attendees & """email"":""" & EscapeJSON(recipient.Address) & ""","
        
        ' Response status
        Select Case recipient.MeetingResponseStatus
            Case olResponseAccepted
                responseStatus = "Accepted"
            Case olResponseDeclined
                responseStatus = "Declined"
            Case olResponseTentative
                responseStatus = "Tentative"
            Case olResponseNotResponded
                responseStatus = "Not Responded"
            Case Else
                responseStatus = "Unknown"
        End Select
        
        attendees = attendees & """responseStatus"":""" & responseStatus & """"
        attendees = attendees & "}"
    Next
    
    attendees = attendees & "]"
    
    ' Create JSON object with meeting details and attendees
    Dim json
    
    json = "{"
    json = json & """subject"":""" & EscapeJSON(appointment.Subject) & ""","
    json = json & """start"":""" & FormatDateTime(appointment.Start) & ""","
    json = json & """end"":""" & FormatDateTime(appointment.End) & ""","
    json = json & """location"":""" & EscapeJSON(appointment.Location) & ""","
    json = json & """organizer"":""" & EscapeJSON(appointment.Organizer) & ""","
    json = json & """attendees"":" & attendees
    json = json & "}"
    
    GetAttendeeStatus = json
    
    ' Clean up
    Set appointment = Nothing
    Set calendar = Nothing
    Set outlookApp = Nothing
End Function

' Run the main function
Main
