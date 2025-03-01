' deleteEvent.vbs - Deletes a calendar event by its EntryID
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Main function
Sub Main()
    ' Get command line arguments
    Dim eventId, calendarName
    
    ' Get and validate arguments
    eventId = GetArgument("eventId")
    calendarName = GetArgument("calendar")
    
    ' Require event ID
    RequireArgument "eventId"
    
    ' Delete the event
    Dim result
    result = DeleteCalendarEvent(eventId, calendarName)
    
    ' Output success
    OutputSuccess "{""success"":" & LCase(CStr(result)) & "}"
End Sub

' Deletes a calendar event with the specified EntryID
Function DeleteCalendarEvent(eventId, calendarName)
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
        DeleteCalendarEvent = False
        Exit Function
    End If
    
    ' Delete the appointment
    appointment.Delete
    
    If Err.Number <> 0 Then
        OutputError "Failed to delete calendar event: " & Err.Description
        DeleteCalendarEvent = False
    Else
        DeleteCalendarEvent = True
    End If
    
    ' Clean up
    Set appointment = Nothing
    Set calendar = Nothing
    Set namespace = Nothing
    Set outlookApp = Nothing
End Function

' Run the main function
Main
