' utils.vbs - Shared utility functions for Outlook Calendar operations
Option Explicit

' Constants
Const olFolderCalendar = 9
Const olAppointmentItem = 1
Const olMeeting = 1
Const olBusy = 2
Const olTentative = 1
Const olFree = 0
Const olOutOfOffice = 3
Const olResponseAccepted = 3
Const olResponseDeclined = 4
Const olResponseTentative = 2
Const olResponseNotResponded = 5

' Error handling constants
Const ERROR_PREFIX = "ERROR:"
Const SUCCESS_PREFIX = "SUCCESS:"

' ===== Outlook Application Management =====

' Creates and returns an Outlook Application object
Function CreateOutlookApplication()
    On Error Resume Next
    Dim outlookApp
    Set outlookApp = CreateObject("Outlook.Application")
    
    If Err.Number <> 0 Then
        WScript.Echo ERROR_PREFIX & "Failed to create Outlook Application: " & Err.Description
        WScript.Quit 1
    End If
    
    Set CreateOutlookApplication = outlookApp
End Function

' Gets the default calendar folder from Outlook
Function GetDefaultCalendar(outlookApp)
    On Error Resume Next
    Dim namespace, calendar
    
    Set namespace = outlookApp.GetNamespace("MAPI")
    If Err.Number <> 0 Then
        WScript.Echo ERROR_PREFIX & "Failed to get MAPI namespace: " & Err.Description
        WScript.Quit 1
    End If
    
    Set calendar = namespace.GetDefaultFolder(olFolderCalendar)
    If Err.Number <> 0 Then
        WScript.Echo ERROR_PREFIX & "Failed to get default calendar: " & Err.Description
        WScript.Quit 1
    End If
    
    Set GetDefaultCalendar = calendar
End Function

' Gets a specific calendar folder by name
Function GetCalendarByName(outlookApp, calendarName)
    On Error Resume Next
    Dim namespace, folders, folder, i
    
    Set namespace = outlookApp.GetNamespace("MAPI")
    If Err.Number <> 0 Then
        WScript.Echo ERROR_PREFIX & "Failed to get MAPI namespace: " & Err.Description
        WScript.Quit 1
    End If
    
    ' Get default calendar if no name specified
    If calendarName = "" Then
        Set GetCalendarByName = GetDefaultCalendar(outlookApp)
        Exit Function
    End If
    
    ' Try to find the specified calendar
    Set folders = namespace.Folders
    For i = 1 To folders.Count
        Set folder = folders.Item(i)
        If folder.Name = calendarName Then
            Set GetCalendarByName = folder.GetDefaultFolder(olFolderCalendar)
            Exit Function
        End If
    Next
    
    ' Calendar not found
    WScript.Echo ERROR_PREFIX & "Calendar not found: " & calendarName
    WScript.Quit 1
End Function

' ===== Date Handling =====

' Converts a date string in MM/DD/YYYY format to a Date object
Function ParseDate(dateStr)
    On Error Resume Next
    
    If IsDate(dateStr) Then
        ParseDate = CDate(dateStr)
    Else
        ' Try to parse MM/DD/YYYY format
        Dim parts, month, day, year
        parts = Split(dateStr, "/")
        
        If UBound(parts) = 2 Then
            month = parts(0)
            day = parts(1)
            year = parts(2)
            
            If IsNumeric(month) And IsNumeric(day) And IsNumeric(year) Then
                ParseDate = DateSerial(year, month, day)
            Else
                WScript.Echo ERROR_PREFIX & "Invalid date format. Expected MM/DD/YYYY: " & dateStr
                WScript.Quit 1
            End If
        Else
            WScript.Echo ERROR_PREFIX & "Invalid date format. Expected MM/DD/YYYY: " & dateStr
            WScript.Quit 1
        End If
    End If
    
    If Err.Number <> 0 Then
        WScript.Echo ERROR_PREFIX & "Failed to parse date: " & dateStr & " - " & Err.Description
        WScript.Quit 1
    End If
End Function

' Formats a Date object to MM/DD/YYYY format
Function FormatDate(dateObj)
    FormatDate = Month(dateObj) & "/" & Day(dateObj) & "/" & Year(dateObj)
End Function

' Formats a Date object to MM/DD/YYYY HH:MM AM/PM format
Function FormatDateTime(dateTimeObj)
    FormatDateTime = FormatDate(dateTimeObj) & " " & FormatTime(dateTimeObj)
End Function

' Formats a time to HH:MM AM/PM format
Function FormatTime(dateTimeObj)
    Dim hours, minutes, ampm
    
    hours = Hour(dateTimeObj)
    minutes = Minute(dateTimeObj)
    
    If hours >= 12 Then
        ampm = "PM"
        If hours > 12 Then hours = hours - 12
    Else
        ampm = "AM"
        If hours = 0 Then hours = 12
    End If
    
    FormatTime = Right("0" & hours, 2) & ":" & Right("0" & minutes, 2) & " " & ampm
End Function

' ===== JSON Handling =====

' Escapes a string for JSON with proper Unicode handling
Function EscapeJSON(str)
    Dim result, i, char, charCode
    
    If IsNull(str) Or str = "" Then
        EscapeJSON = ""
        Exit Function
    End If
    
    result = ""
    For i = 1 To Len(str)
        char = Mid(str, i, 1)
        charCode = AscW(char)
        
        Select Case char
            Case "\"
                result = result & "\\"
            Case """"
                result = result & "\"""
            Case vbCrLf
                result = result & "\n"
            Case vbCr
                result = result & "\n"
            Case vbLf
                result = result & "\n"
            Case vbTab
                result = result & "\t"
            Case Else
                ' Handle Unicode characters properly
                If charCode > 127 Then
                    result = result & "\u" & Right("0000" & Hex(charCode), 4)
                Else
                    result = result & char
                End If
        End Select
    Next
    
    EscapeJSON = result
End Function

' Converts a VBScript array to a JSON array
Function ArrayToJSON(arr)
    Dim i, result
    
    result = "["
    For i = LBound(arr) To UBound(arr)
        If i > LBound(arr) Then result = result & ","
        
        If IsNull(arr(i)) Then
            result = result & "null"
        ElseIf IsArray(arr(i)) Then
            result = result & ArrayToJSON(arr(i))
        ElseIf IsObject(arr(i)) Then
            result = result & "null" ' Objects not supported in this simple implementation
        ElseIf VarType(arr(i)) = vbString Then
            result = result & """" & EscapeJSON(arr(i)) & """"
        ElseIf VarType(arr(i)) = vbBoolean Then
            If arr(i) Then
                result = result & "true"
            Else
                result = result & "false"
            End If
        Else
            result = result & arr(i)
        End If
    Next
    
    result = result & "]"
    ArrayToJSON = result
End Function

' ===== Outlook Item Conversion =====

' Converts an Outlook appointment item to a JSON string
Function AppointmentToJSON(appointment)
    Dim json, recipients, recipient, i, attendees, attendeeStatus
    
    ' Start building the JSON object
    json = "{"
    
    ' Include EntryID for event identification
    json = json & """id"":""" & EscapeJSON(appointment.EntryID) & ""","
    
    ' Basic properties
    json = json & """subject"":""" & EscapeJSON(appointment.Subject) & ""","
    json = json & """start"":""" & FormatDateTime(appointment.Start) & ""","
    json = json & """end"":""" & FormatDateTime(appointment.End) & ""","
    json = json & """location"":""" & EscapeJSON(appointment.Location) & ""","
    json = json & """body"":""" & EscapeJSON(appointment.Body) & ""","
    json = json & """organizer"":""" & EscapeJSON(appointment.Organizer) & ""","
    json = json & """isRecurring"":" & LCase(CStr(appointment.IsRecurring)) & ","
    
    ' Meeting status
    json = json & """isMeeting"":" & LCase(CStr(appointment.MeetingStatus = olMeeting)) & ","
    
    ' Busy status
    Select Case appointment.BusyStatus
        Case olBusy
            json = json & """busyStatus"":""Busy"","
        Case olTentative
            json = json & """busyStatus"":""Tentative"","
        Case olFree
            json = json & """busyStatus"":""Free"","
        Case olOutOfOffice
            json = json & """busyStatus"":""Out of Office"","
        Case Else
            json = json & """busyStatus"":""Unknown"","
    End Select
    
    ' Attendees (if it's a meeting)
    If appointment.MeetingStatus = olMeeting Then
        Set recipients = appointment.Recipients
        attendees = ""
        
        For i = 1 To recipients.Count
            Set recipient = recipients.Item(i)
            
            If i > 1 Then attendees = attendees & ","
            
            attendees = attendees & "{"
            attendees = attendees & """name"":""" & EscapeJSON(recipient.Name) & ""","
            attendees = attendees & """email"":""" & EscapeJSON(recipient.Address) & ""","
            
            ' Response status
            Select Case recipient.MeetingResponseStatus
                Case olResponseAccepted
                    attendeeStatus = "Accepted"
                Case olResponseDeclined
                    attendeeStatus = "Declined"
                Case olResponseTentative
                    attendeeStatus = "Tentative"
                Case olResponseNotResponded
                    attendeeStatus = "Not Responded"
                Case Else
                    attendeeStatus = "Unknown"
            End Select
            
            attendees = attendees & """responseStatus"":""" & attendeeStatus & """"
            attendees = attendees & "}"
        Next
        
        json = json & """attendees"":[" & attendees & "]"
    Else
        json = json & """attendees"":[]"
    End If
    
    ' Close the JSON object
    json = json & "}"
    
    AppointmentToJSON = json
End Function

' Converts a collection of Outlook appointment items to a JSON array
Function AppointmentsToJSON(appointments)
    Dim i, json
    
    json = "["
    
    For i = 1 To appointments.Count
        If i > 1 Then json = json & ","
        json = json & AppointmentToJSON(appointments.Item(i))
    Next
    
    json = json & "]"
    
    AppointmentsToJSON = json
End Function

' ===== Command Line Argument Handling =====

' Gets a command line argument by name
Function GetArgument(name)
    Dim args, i, arg, parts
    
    Set args = WScript.Arguments
    
    For i = 0 To args.Count - 1
        arg = args(i)
        
        If Left(arg, 1) = "/" Or Left(arg, 1) = "-" Then
            parts = Split(Mid(arg, 2), ":", 2)
            
            If UBound(parts) >= 0 Then
                If LCase(parts(0)) = LCase(name) Then
                    If UBound(parts) = 1 Then
                        GetArgument = parts(1)
                    Else
                        GetArgument = "true"
                    End If
                    Exit Function
                End If
            End If
        End If
    Next
    
    GetArgument = ""
End Function

' Checks if a required argument is present
Sub RequireArgument(name)
    Dim value
    
    value = GetArgument(name)
    
    If value = "" Then
        WScript.Echo ERROR_PREFIX & "Missing required argument: " & name
        WScript.Quit 1
    End If
End Sub

' ===== Output Formatting =====

' Outputs a success message with JSON data
Sub OutputSuccess(jsonData)
    WScript.Echo SUCCESS_PREFIX & jsonData
End Sub

' Outputs an error message
Sub OutputError(message)
    WScript.Echo ERROR_PREFIX & message
End Sub
