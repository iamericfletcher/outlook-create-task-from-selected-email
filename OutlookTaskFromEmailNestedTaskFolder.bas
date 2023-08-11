Sub CreateTaskFromEmail()

    Dim objMail As MailItem
    Dim objTask As TaskItem
    Dim objNamespace As NameSpace
    Dim objTaskFolder As Folder
    Dim strCategory As String
    
    ' Specify the target Task Folder name
    Const TARGET_FOLDER_NAME As String = "INSERT TASK FOLDER NAME"
    
    ' Specify the category to apply
    strCategory = "INSERT CATEGORY"
    
    ' Handle errors while working with objects
    On Error Resume Next
    
    ' Get the currently selected item in Outlook
    Set objMail = Application.ActiveExplorer.Selection.Item(1)
    
    ' Check if an email is selected
    If objMail Is Nothing Then
        MsgBox "Please select an email.", vbInformation
        Exit Sub
    End If
    
    ' Check if the selected item is a mail item
    If objMail.Class = olMail Then
        ' Get the MAPI namespace
        Set objNamespace = Application.GetNamespace("MAPI")
        
        ' Locate the target task folder
        Set objTaskFolder = objNamespace.GetDefaultFolder(olFolderTasks).Folders(TARGET_FOLDER_NAME)
        
        ' Check if the target task folder exists
        If objTaskFolder Is Nothing Then
            MsgBox "The specified task folder does not exist. Please check the folder name.", vbInformation
            Exit Sub
        End If
        
        ' Create a new task item in the target task folder
        Set objTask = objTaskFolder.Items.Add(olTaskItem)
        With objTask
            .Subject = objMail.Subject ' Set the task subject to the email subject
            .StartDate = objMail.ReceivedTime ' Set the start date of the task
            .Categories = strCategory ' Set the task category
            .Display ' Open the task window for editing
        End With
        
        MsgBox "A new task has been created in the specified folder.", vbInformation
    Else
        MsgBox "Please select an email.", vbInformation
    End If
    
    ' Release the object variables
    Set objMail = Nothing
    Set objTask = Nothing
    Set objNamespace = Nothing
    Set objTaskFolder = Nothing

End Sub
