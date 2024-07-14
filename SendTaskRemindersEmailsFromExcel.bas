Sub SendTaskReminderEmails()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim currentDate As Date
    Dim consultantDict As Object
    Dim consultant As Variant
    Dim consultantEmail As String
    Dim mailSubject As String
    Dim mailBody As String
    Dim taskTable As String
    Dim filteredRange As Range
    Dim emailCount As Long

    ' Set up Outlook
    Set OutApp = CreateObject("Outlook.Application")
    ' Ensure the correct worksheet name
    Set ws = ThisWorkbook.Sheets("ProjectTimeline") ' Change "Sheet1" to the actual name of your worksheet

    ' Check if the table exists
    On Error Resume Next
    Set lo = ws.ListObjects("Jalons43525")
    On Error GoTo 0

    If lo Is Nothing Then
        MsgBox "The table 'Jalons43525' does not exist on the sheet '" & ws.Name & "'.", vbCritical
        Exit Sub
    End If

    ' Get the current date
    currentDate = Date

    ' Create a dictionary to store consultant emails
    Set consultantDict = CreateObject("Scripting.Dictionary")

    ' Loop through each row in the task list to populate the dictionary
    For i = 12 To ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
        consultant = ws.Cells(i, "J").Value
        consultantEmail = ws.Cells(i, "K").Value
        If Not consultantDict.Exists(consultant) And consultantEmail <> "" Then
            consultantDict.Add consultant, consultantEmail
        End If
    Next i

    ' Initialize email count
    emailCount = 0

    ' Loop through each consultant in the dictionary
    For Each consultant In consultantDict.Keys
        consultantEmail = consultantDict(consultant)

        ' Apply filters to the table
        With lo.Range
            .AutoFilter Field:=9, Criteria1:=consultant ' Consultant (Column J, Field 10 in the table)
            .AutoFilter Field:=12, Criteria1:="Started" ' Status (Column M, Field 13 in the table)
        End With

        ' Set the filtered range
        On Error Resume Next
        Set filteredRange = lo.DataBodyRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not filteredRange Is Nothing Then
            ' Convert the filtered range to HTML table with headers
            taskTable = RangetoHTML(filteredRange, lo)

            ' Prepare email content
            mailSubject = "Task Reminder"
            mailBody = "Dear " & consultant & "," & vbCrLf & vbCrLf & _
                       "Please find below the pending tasks assigned to you:" & vbCrLf & vbCrLf & _
                       "Task Table:" & vbCrLf & vbCrLf & _
                       "Please update the status or provide an update on the progress. Thank you."

            ' Create a new email
            Set OutMail = OutApp.CreateItem(0)
            On Error Resume Next
            With OutMail
                .To = consultantEmail
                .Subject = mailSubject
                .HTMLBody = "<html><body>" & mailBody & "<br><br>" & taskTable & "</body></html>"
                .Send  ' Change to .Display to manually review before sending
            End With
            On Error GoTo 0

            If Err.Number = 0 Then
                ' Increase email count
                emailCount = emailCount + 1
            Else
                MsgBox "Error sending email to " & consultantEmail & ".", vbExclamation
            End If

            ' Clean up
            Set OutMail = Nothing
        End If

        ' Remove the filter
        lo.AutoFilter.ShowAllData
    Next consultant

    ' Display a message box when all emails have been sent
    MsgBox emailCount & " emails have been sent successfully.", vbInformation

    ' Clean up
    Set OutApp = Nothing
    Set consultantDict = Nothing
End Sub


Function RangetoHTML(rng As Range, lo As ListObject) As String
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
    Dim headerRange As Range
    Dim headerCell As Range
    Dim htmlTable As String

    TempFile = Environ("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    ' Copy the headers and data to a new workbook to convert to HTML
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        ' Copy headers from ListObject
        Set headerRange = lo.HeaderRowRange
        For Each headerCell In headerRange.Cells
            .Cells(1, headerCell.Column).Value = headerCell.Value
        Next headerCell

        ' Copy data
        rng.Copy
        .Cells(2, 1).PasteSpecial Paste:=8 ' Paste column widths
        .Cells(2, 1).PasteSpecial Paste:=xlPasteValues
        .Cells(2, 1).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False

        ' Format table
        .UsedRange.Columns.AutoFit

        ' Convert to HTML
        htmlTable = .UsedRange.HTMLBody
    End With

    ' Save the temporary workbook as HTML
    TempFile = Environ("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
    TempWB.PublishObjects.Add(SourceType:=xlSourceRange, _
     Filename:=TempFile, _
     Sheet:=TempWB.Sheets(1).Name, _
     Source:=TempWB.Sheets(1).UsedRange.Address, _
     HtmlType:=xlHtmlStatic).Publish True

    ' Read the HTML file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close

    ' Delete the temporary file
    TempWB.Close SaveChanges:=False
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function