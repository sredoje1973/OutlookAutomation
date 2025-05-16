
'  ---------------------------------
' Macros to help find available timeslots in the calendar
' author: Srdoje Vakareskov May 16, 2025
'
' Written with a little bit of help from Copilot
' ... lots of human refinements required.
'
'
'' Step-by-Step: Add a Macro to a Button in Outlook
' 1. Create or Open Your Macro
' Press ALT + F11 to open the VBA Editor.
' Insert your macro into a module:
' Go to Insert > Module
' Paste your macro code there.
' Save and close the editor.
' 2. Add the Macro to the Ribbon or Quick Access Toolbar
' Option A: Add to the Quick Access Toolbar
' In Outlook, click the down arrow at the top-left (Quick Access Toolbar).
' Choose More Commands.
' In the dropdown, select Macros.
' Select your macro (e.g., Project1.Module1.EmailFreeTimeSlotsWithCounts).
' Click Add >>.
' (Optional) Click Modify to choose an icon and rename the button.
' Click OK.
' Option B: Add to the Ribbon
' Go to File > Options > Customize Ribbon.
' On the right, choose a tab (e.g., Home) and click New Group.
' With the new group selected, click Choose commands from > Macros.
' Select your macro and click Add >>.
' (Optional) Click Rename to customize the button.
' Click OK.
'
'--------------------------------------



Function RoundUpToNextHalfHour(dt As Date) As Date
    Dim minutes As Integer
    minutes = Minute(dt)
    
    If minutes = 0 Or minutes = 30 Then
        RoundUpToNextHalfHour = dt
    ElseIf minutes < 30 Then
        RoundUpToNextHalfHour = DateAdd("n", 30 - minutes, dt)
    Else
        RoundUpToNextHalfHour = DateAdd("n", 60 - minutes, dt)
    End If
End Function

Sub EmailFreeTimeSlotsWithCounts()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olCalendar As Outlook.Folder
    Dim olItems As Outlook.Items
    Dim dailyItems As Outlook.Items
    Dim appt As Outlook.AppointmentItem
    Dim StartDate As Date, EndDate As Date
    Dim WorkStart As Date, WorkEnd As Date
    Dim i As Integer
    Dim currentDate As Date
    Dim FreeSlots As String
    Dim roundedStart As Date
    Dim minDurationMinutes As Integer: minDurationMinutes = 30
    Dim numberOfBizDays As Integer

    ' Set working hours locally
    WorkStart = TimeValue("09:00:00")
    WorkEnd = TimeValue("16:30:00")

    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olCalendar = olNS.GetDefaultFolder(olFolderCalendar)

    ' Define date range (next 10 business days)
    
    StartDate = Date
    EndDate = StartDate + 14 ' Includes weekends, we'll skip them
    numberOfBizDays = 10
    ' Loop through each day
    For i = 0 To 13
        currentDate = StartDate + i
        If Weekday(currentDate, vbMonday) <= 5 Then ' Skip weekends
            Set olItems = olCalendar.Items
            olItems.Sort "[Start]"
            olItems.IncludeRecurrences = True

            Dim filter As String
            filter = "[Start] >= '" & Format(currentDate, "ddddd h:nn AMPM") & "' AND [Start] < '" & Format(currentDate + 1, "ddddd h:nn AMPM") & "'"
            Set dailyItems = olItems.Restrict(filter)
            
            ' Skip days with a full day event
            Dim skipDay As Boolean: skipDay = False
            For Each appt In dailyItems
                If appt.AllDayEvent = True Then
                    skipDay = True
                        Exit For
                End If
            Next
            

                Dim busyTimes() As Variant
                Dim count As Integer: count = 0
    
                ' Collect busy/tentative appointments
                For Each appt In dailyItems
                
                '    If appt.BusyStatus = olBusy Or appt.BusyStatus = olTentative Then
                    If (appt.BusyStatus = olBusy Or appt.BusyStatus = olTentative) And appt.BusyStatus <> olOutOfOffice Then
                        ReDim Preserve busyTimes(1 To 2, 1 To count + 1)
                        busyTimes(1, count + 1) = appt.Start
                        busyTimes(2, count + 1) = appt.End
                        count = count + 1
                    End If
                Next
                
         
                
                
                
    
                ' Sort busy times
                Dim j As Integer, k As Integer
                For j = 1 To count - 1
                    For k = j + 1 To count
                        If busyTimes(1, j) > busyTimes(1, k) Then
                            Dim tempStart As Variant, tempEnd As Variant
                            tempStart = busyTimes(1, j)
                            tempEnd = busyTimes(2, j)
                            busyTimes(1, j) = busyTimes(1, k)
                            busyTimes(2, j) = busyTimes(2, k)
                            busyTimes(1, k) = tempStart
                            busyTimes(2, k) = tempEnd
                        End If
                    Next k
                Next j
    
                ' Store free slots
                Dim dailyFreeSlots() As String
                Dim slotCount As Integer: slotCount = 0
                Dim slotStart As Date, slotEnd As Date
                slotStart = currentDate + WorkStart
    
                For j = 1 To count
                    slotEnd = busyTimes(1, j)
                    If slotStart < slotEnd Then
                        roundedStart = RoundUpToNextHalfHour(slotStart)
                        If roundedStart < currentDate + WorkStart Then roundedStart = currentDate + WorkStart
                        If slotEnd > currentDate + WorkEnd Then slotEnd = currentDate + WorkEnd
    
                        If roundedStart < slotEnd Then
                            If DateDiff("n", roundedStart, slotEnd) >= minDurationMinutes Then
                                slotCount = slotCount + 1
                                ReDim Preserve dailyFreeSlots(1 To slotCount)
                                dailyFreeSlots(slotCount) = Format(roundedStart, "hh:mm AM/PM") & " - " & Format(slotEnd, "hh:mm AM/PM")
                            End If
                        End If
                    End If
                    If busyTimes(2, j) > slotStart Then slotStart = busyTimes(2, j)
                Next j
    
                ' Check after last appointment
                If slotStart < currentDate + WorkEnd Then
                    roundedStart = RoundUpToNextHalfHour(slotStart)
                    If roundedStart < currentDate + WorkStart Then roundedStart = currentDate + WorkStart
                    If currentDate + WorkEnd > roundedStart Then
                        If DateDiff("n", roundedStart, currentDate + WorkEnd) >= minDurationMinutes Then
                            slotCount = slotCount + 1
                            ReDim Preserve dailyFreeSlots(1 To slotCount)
                            dailyFreeSlots(slotCount) = Format(roundedStart, "hh:mm AM/PM") & " - " & Format(currentDate + WorkEnd, "hh:mm AM/PM")
                        End If
                    End If
                End If
    
                ' Add to output
                If skipDay Then slotCount = 0
                If slotCount > 0 Then
                    FreeSlots = FreeSlots & vbCrLf & "**" & Format(currentDate, "dddd, mmmm d, yyyy") & ", " & " slots: " & slotCount & vbCrLf
                Else
                    FreeSlots = FreeSlots & vbCrLf & "**" & Format(currentDate, "dddd, mmmm d, yyyy") & ", " & " slots: none " & vbCrLf
                End If
                
                
                If slotCount > 0 Then
                    FreeSlots = FreeSlots & "• "
                    For j = 1 To slotCount - 1
                        FreeSlots = FreeSlots & dailyFreeSlots(j) & ", "
                    Next j
                    FreeSlots = FreeSlots & dailyFreeSlots(j) & vbCrLf
                End If
                
            End If
        
    Next

    ' Create and display email
    Dim mail As Outlook.MailItem
    Set mail = olApp.CreateItem(olMailItem)
    mail.Display
    
    mail.Subject = "My Availability for the next " & numberOfBizDays & " Business Days"
    mail.Body = "Hi," & vbCrLf & vbCrLf & _
                "Here are my available time slots:" & vbCrLf & vbCrLf & FreeSlots & vbCrLf & _
                "Let me know what works best for you." & vbCrLf & vbCrLf & "Best regards,"
    
    
    

End Sub


