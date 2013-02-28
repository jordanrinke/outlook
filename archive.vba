Private WithEvents colReminders As Outlook.Reminders
Private Sub Application_Startup()
    Set colReminders = Application.Reminders
End Sub
Sub colReminders_BeforeReminderShow(Cancel As Boolean)
    Dim objOutlook As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objTasks As Outlook.Items
    Dim objTask As Outlook.TaskItem
    Dim objTasksFolder As Outlook.Folder
    Dim objFilteredTasks As Outlook.Items
    ' Create an object for the Outlook application.
    Set objOutlook = Application
    ' Retrieve an object for the MAPI namespace.
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objTasksFolder = objNamespace.GetDefaultFolder(olFolderTasks)
    Set objTasks = objTasksFolder.Items
    Debug.Print "reminder event fired"
    For Each objRem In colReminders
        Debug.Print objRem.Caption
            If objRem.Caption = "MOVEOLDEMAILREMINDER" Then
                If objRem.IsVisible Then
                    objRem.Dismiss
                    Cancel = True
                    Set objFilteredTasks = objTasks.Restrict("[Subject] = 'MOVEOLDEMAILREMINDER'")
                    For i = objFilteredTasks.Count To 1 Step -1
                        Set objTask = objFilteredTasks.Item(i)
                        If objTask.Subject = "MOVEOLDEMAILREMINDER" Then
                            With objTask
                                .ReminderTime = DateAdd("d", 1, objTask.ReminderTime)
                                .StartDate = DateAdd("d", 1, Date)
                                .Save
                                .DueDate = DateAdd("d", 2, Date)
                                .Save
                            End With
                            MoveOldEmails
                        End If
                    Next i
                End If
                Exit For
            End If
        Next objRem
End Sub
Sub MoveOldEmails()
'NOTE: This does Sent Items, Inbox, and 1 folder below inbox. It will not do more than 1 folder deep in the inbox.
    ' Declare all variables.
    Dim objOutlook As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objInbox As Outlook.Folder
    Dim objSentbox As Outlook.Folder
    Dim objDestFolder As Outlook.Folder
    Dim objMail As Variant
    Dim intCount As Integer
    Dim intDateDiff As Integer
    Dim intAge As Integer

    'Move anything older than this date.
    intAge = 30

    ' Create an object for the Outlook application.
    Set objOutlook = Application
    ' Retrieve an object for the MAPI namespace.
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    ' Retrieve a folder object for the inbox folder
    Set objInbox = objNamespace.GetDefaultFolder(olFolderInbox)
    'Retrieve a folder object for sent box folder
    Set objSentbox = objNamespace.GetDefaultFolder(olFolderSentMail)
    ' Note: Using cached mode with exchange is much faster, non-cached mode will take 1-2 seconds per email
    ' since a request is made for every object vs using a local cache.

    'Move sent items first.

    For intCount = objSentbox.Items.Count To 1 Step -1
        ' Loop through the items in the folder. NOTE: This has to
        ' be done backwards; if you process forwards you have to
        ' re-run the macro an inverese exponential number of times.
        Set objMail = objSentbox.Items.Item(intCount)
        If objMail.Class = olMail Or objMail.Class = olMeetingRequest Then
            intDateDiff = DateDiff("d", objMail.SentOn, Now)
            Debug.Print "SENT: " & intDateDiff & ":" & objMail.SentOn & ":" & objMail.Subject
            'Move anything older than intAge
            If intDateDiff > intAge Then
                'Set blnFound to False, it will be set to true by logic if the folder does not need to be created
                blnFound = False
                
                For Each objFolder In objNamespace.Folders(CStr(Year(objMail.SentOn))).Folders
                    'This is sloppy since we will loop through all of the folders to see if the folder exists
                    'for every mail item, I could make an array and do a search on it but this should be run in
                    'cached mode so there is not much of a difference. This is however, very much not optimal
                    If objFolder.Name = objSentbox.Name Then
                        'If the current folder matches our search folder set blnFound to true to skip folder creation
                        Debug.Print "Folder Exists: " & objFolder.Name
                        blnFound = True
                    End If
                Next
                
                If blnFound = False Then
                    'Create the folder if it was not found
                    Debug.Print "Creating Folder: " & objSentbox.Name
                    Debug.Print CStr(Year(objMail.SentOn))
                    objNamespace.Folders(CStr(Year(objMail.SentOn))).Folders.Add objSentbox.Name
                    Debug.Print "created"
                End If
                
                
    
                Debug.Print objMail.SentOn & ":" & objMail.Subject
                'Set the destination to the same structure as the source folder, i.e. "2010Sent Items"
                Set objDestFolder = objNamespace.Folders(CStr(Year(objMail.SentOn))).Folders(objSentbox.Name)
                'Move the object - technically mail is not the best name since this can be calendar items etc but I liked it more than "variant"
                objMail.Move objDestFolder
                'Destroy object for clarity
                Set objDestFolder = Nothing
            End If
        End If
    Next intCount

    'Next move everything in the root of the default Inbox
    For intCount = objInbox.Items.Count To 1 Step -1
        Set objMail = objInbox.Items.Item(intCount)
        If objMail.Class = olMail Or objMail.Class = olMeetingRequest Then
            Debug.Print objMail.Subject
            intDateDiff = DateDiff("d", objMail.SentOn, Now)
            Debug.Print "ROOT: " & intDateDiff & ":" & objMail.SentOn & ":" & objMail.Subject
            If intDateDiff > intAge Then
                blnFound = False
                For Each objFolder In objNamespace.Folders(CStr(Year(objMail.SentOn))).Folders
                    If objFolder.Name = objInbox.Name Then
                        Debug.Print "Folder Exists: " & objFolder.Name
                        blnFound = True
                    End If
                Next
                If blnFound = False Then
                    Debug.Print "Creating Folder: " & objInbox.Name
                    Debug.Print CStr(Year(objMail.SentOn))
                    objNamespace.Folders(CStr(Year(objMail.SentOn))).Folders.Add objInbox.Name
                End If
    
                Debug.Print objMail.SentOn & ":" & objMail.Subject
                ' folder structure i.e. "2010Inbox"
                Set objDestFolder = objNamespace.Folders(CStr(Year(objMail.SentOn))).Folders(objInbox.Name)
                objMail.Move objDestFolder
                Set objDestFolder = Nothing
            End If
        End If
    Next intCount

   'Loop through all the folders in the inbox
    For intFolderCount = 1 To objInbox.Folders.Count
        For intCount = objInbox.Folders(intFolderCount).Items.Count To 1 Step -1
            DoEvents
            Set objMail = objInbox.Folders(intFolderCount).Items.Item(intCount)
            If objMail.Class = olMail Or objMail.Class = olMeetingRequest Then
                intDateDiff = DateDiff("d", objMail.SentOn, Now)
                If intDateDiff > intAge Then
                    blnFound = False
                    For Each objFolder In objNamespace.Folders(CStr(Year(objMail.SentOn))).Folders(objInbox.Name).Folders
                        If objFolder.Name = objInbox.Folders(intFolderCount).Name Then
                            Debug.Print "Folder Exists: " & objFolder.Name
                            blnFound = True
                        End If
                    Next
                    If blnFound = False Then
                        Debug.Print "Creating Folder: " & objInbox.Folders(intFolderCount).Name
                        objNamespace.Folders(CStr(Year(objMail.SentOn))).Folders(objInbox.Name).Folders.Add objInbox.Folders(intFolderCount).Name
                    End If
    
                    Debug.Print objInbox.Folders(intFolderCount).Name & ":" & objMail.SentOn & ":" & objMail.Subject
                    ' folder structure i.e. "2010InboxsubFolder"
                    Set objDestFolder = objNamespace.Folders(CStr(Year(objMail.SentOn))).Folders(objInbox.Name).Folders(objInbox.Folders(intFolderCount).Name)
                    If objMail.Class = olMail Or objMail.Class = olMeetingRequest Then
                        objMail.Move objDestFolder
                    End If
                    Set objDestFolder = Nothing
                End If
            End If
        Next intCount
    Next intFolderCount

    Debug.Print "Done"

End Sub
