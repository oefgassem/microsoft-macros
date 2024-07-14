Sub SendTaskReminderEmails()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim t As Task
    Dim currentDate As Date
    Dim taskTable As String
    Dim emailCount As Long

    ' Set up Outlook
    Set OutApp = CreateObject("Outlook.Application")
    
    ' Next week
    currentDate = Date + 8
    
    ' Initialize email count
    emailCount = 0
    
    ' Loop through each task in the active project
    For Each t In ActiveProject.Tasks
        ' Check if the task has an assigned resource, is not completed, and the planned finish date is past the current date
        If Not t Is Nothing Then
            If t.resourceNames <> "" And t.PercentComplete < 100 And t.Start <= currentDate Then
                ' Prepare task table as HTML
                taskTable = "<table border=1><tr><th>Task Name</th><th>Start Date</th><th>Finish Date</th><th>Resource Names</th></tr>"
                taskTable = taskTable & "<tr><td>" & t.Name & "</td><td>" & t.Start & "</td><td>" & t.Finish & "</td><td>" & t.resourceNames & "</td></tr></table>"
                
                ' Prepare email content
                Dim mailSubject As String
                Dim mailBody As String
                
                mailSubject = "Task Reminder"
                mailBody = "Dear " & t.resourceNames & "," & "<br><br>" & _
                           "Please find below the pending tasks assigned to you:" & "<br><br>" & _
                           taskTable & "<br>" & _
                           "Please update the status or provide an update on the progress. Thank you."
                
                ' Create a new email
                Set OutMail = OutApp.CreateItem(0)
                With OutMail
                    .To = GetResourceEmail(t.resourceNames)
                    .Subject = mailSubject
                    .HTMLBody = "<html><body>" & mailBody & "</body></html>"
                    .Send  ' Change to .Display to manually review before sending
                End With
                
                ' Increase email count
                emailCount = emailCount + 1
                
                ' Clean up
                Set OutMail = Nothing
            End If
        End If
    Next t
    
    ' Display a message box when all emails have been sent
    MsgBox emailCount & " emails have been sent successfully.", vbInformation
    
    ' Clean up
    Set OutApp = Nothing
End Sub

' Function to get the email of the resource
Function GetResourceEmail(resourceNames As String) As String
    Dim resourceName As String
    Dim resourceEmail As String
    Dim emailList As String
    Dim resource As resource
    
    emailList = ""
    
    For Each resource In ActiveProject.Resources
        resourceName = Trim(resource.Name)
        If InStr(resourceNames, resourceName) > 0 Then
            If resource.EMailAddress <> "" Then
                If emailList = "" Then
                    emailList = resource.EMailAddress
                Else
                    emailList = emailList & ";" & resource.EMailAddress
                End If
            End If
        End If
    Next resource
    
    GetResourceEmail = emailList
End Function