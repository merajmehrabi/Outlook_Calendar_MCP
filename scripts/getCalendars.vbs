' getCalendars.vbs - Lists available calendars
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Main function
Sub Main()
    ' Get calendars
    Dim calendars
    calendars = GetAvailableCalendars()
    
    ' Output calendars as JSON
    OutputSuccess calendars
End Sub

' Gets available calendars
Function GetAvailableCalendars()
    On Error Resume Next
    
    ' Create Outlook objects
    Dim outlookApp, namespace, folders, folder, calendarFolder, i, j
    Dim calendars, calendarName, calendarOwner
    
    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()
    
    ' Get MAPI namespace
    Set namespace = outlookApp.GetNamespace("MAPI")
    
    If Err.Number <> 0 Then
        OutputError "Failed to get MAPI namespace: " & Err.Description
        WScript.Quit 1
    End If
    
    ' Get all folders
    Set folders = namespace.Folders
    
    ' Create JSON array for calendars
    calendars = "["
    
    ' Add default calendar
    calendars = calendars & "{""name"":""Default"",""owner"":""" & EscapeJSON(namespace.CurrentUser.Name) & """,""isDefault"":true}"
    
    ' Loop through all folders (accounts)
    For i = 1 To folders.Count
        Set folder = folders.Item(i)
        
        ' Try to get calendar folder
        On Error Resume Next
        Set calendarFolder = folder.Folders("Calendar")
        
        ' If calendar folder exists
        If Not calendarFolder Is Nothing Then
            ' Get calendar name and owner
            calendarName = folder.Name & " - Calendar"
            calendarOwner = folder.Name
            
            ' Add to JSON array
            calendars = calendars & ",{""name"":""" & EscapeJSON(calendarName) & """,""owner"":""" & EscapeJSON(calendarOwner) & """,""isDefault"":false}"
        End If
        
        ' Reset error handling
        On Error GoTo 0
    Next
    
    ' Close JSON array
    calendars = calendars & "]"
    
    GetAvailableCalendars = calendars
    
    ' Clean up
    Set namespace = Nothing
    Set outlookApp = Nothing
End Function

' Run the main function
Main
