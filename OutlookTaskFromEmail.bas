Sub CreateTaskFromEmail()

    Dim objTask As Outlook.TaskItem
    Dim objMail As Outlook.MailItem
    Dim objSelection As Selection

    ' Check if the selected item is an email
    Set objSelection = Outlook.Application.ActiveExplorer.Selection

    If objSelection.Count = 0 Then
        MsgBox "Please select an email."
        Exit Sub
    End If

    If Not TypeOf objSelection.Item(1) Is MailItem Then
        MsgBox "Please select an email."
        Exit Sub
    End If

    Set objMail = objSelection.Item(1)

    ' Create a new task item
    Set objTask = Application.CreateItem(olTaskItem)
    
    With objTask
        ' Set the subject of the task
        .Subject = objMail.Subject

        ' Set the start date of the task
        .StartDate = objMail.ReceivedTime

        ' Assign the category
        .Categories = "INSERT YOUR TASK CATEGORY HERE" 

        ' Open the task window for editing
        .Display
    End With

End Sub
