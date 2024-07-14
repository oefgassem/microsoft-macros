Sub ExportTasksToExcel()
    ' Déclaration des variables
    Dim pjProj As Project
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim i As Integer
    Dim row As Integer
    Dim ProgressForm As New UserForm1
    Dim TotalTasks As Long
    Dim CurrentTask As Long
    Dim ProgressPercentage As Double
    Dim filePath As String
    
    ' Initialisation de l'application Excel via late binding
    Set xlApp = CreateObject("Excel.Application")
    
    ' Ouvrir le projet actif dans MS Project
    Set pjProj = ActiveProject
    
    ' Créer un nouveau classeur Excel
    Set xlBook = xlApp.Workbooks.Add
    
    ' Référencer la première feuille de calcul du classeur Excel
    Set xlSheet = xlBook.Sheets(1)
    
    ' Nommer la feuille Excel
    xlSheet.Name = "Exported Tasks"
    
    ' Écrire les en-têtes dans Excel
    xlSheet.Cells(1, 1).Value = "ID"
    xlSheet.Cells(1, 2).Value = "Process"
    xlSheet.Cells(1, 3).Value = "Task Name"
    xlSheet.Cells(1, 4).Value = "Duration"
    xlSheet.Cells(1, 5).Value = "Start Date"
    xlSheet.Cells(1, 6).Value = "Finish Date"
    xlSheet.Cells(1, 7).Value = "Baseline Duration"
    xlSheet.Cells(1, 8).Value = "Baseline Start Date"
    xlSheet.Cells(1, 9).Value = "Baseline Finish Date"
    xlSheet.Cells(1, 10).Value = "Resource Names"
    xlSheet.Cells(1, 11).Value = "Resource Groups"
    xlSheet.Cells(1, 12).Value = "% complete"
    xlSheet.Cells(1, 13).Value = "Notes"
    
    ' Initialiser le compteur de lignes Excel
    row = 2
    
    ' Calculate total tasks
    TotalTasks = pjProj.Tasks.Count

    ' Initialize progress bar
    ProgressForm.Frame1.Width = 216
    ProgressForm.Label1.Width = 0
    ProgressForm.Show vbModeless
    
    ' Parcourir les tâches du projet
    For i = 1 To pjProj.Tasks.Count
        If Not pjProj.Tasks(i) Is Nothing Then
            If pjProj.Tasks(i).Summary = False Then
                ' Exporter les détails de la tâche dans Excel
                xlSheet.Cells(row, 1).Value = pjProj.Tasks(i).ID
                xlSheet.Cells(row, 2).Value = pjProj.Tasks(i).Text3
                xlSheet.Cells(row, 3).Value = pjProj.Tasks(i).Name
                xlSheet.Cells(row, 4).Value = pjProj.Tasks(i).Duration
                xlSheet.Cells(row, 5).Value = pjProj.Tasks(i).Start
                xlSheet.Cells(row, 6).Value = pjProj.Tasks(i).Finish
                xlSheet.Cells(row, 7).Value = pjProj.Tasks(i).BaselineDuration
                xlSheet.Cells(row, 8).Value = pjProj.Tasks(i).BaselineStart
                xlSheet.Cells(row, 9).Value = pjProj.Tasks(i).BaselineFinish
                xlSheet.Cells(row, 10).Value = pjProj.Tasks(i).resourceNames
                xlSheet.Cells(row, 11).Value = pjProj.Tasks(i).ResourceGroup
                xlSheet.Cells(row, 12).Value = pjProj.Tasks(i).PercentComplete
                xlSheet.Cells(row, 13).Value = pjProj.Tasks(i).Notes
                
                ' Incrémenter le compteur de lignes Excel
                row = row + 1
                
                ' Update progress bar
                CurrentTask = CurrentTask + 1
                ProgressPercentage = CurrentTask / TotalTasks * 100
                ProgressForm.Label1.Width = ProgressForm.Frame1.Width * ProgressPercentage / 100
                DoEvents
            End If
        End If
    Next i
    
    ' Hide progress bar
    ProgressForm.Hide
    
    ' Delete Export if exists
    filePath = "Chemin vers le fichier\ExportProject.xlsx"
    DeleteFileIfExists filePath
    
    ' Sauvegarder le classeur Excel
    xlBook.SaveAs "Chemin vers le fichier\ExportProject.xlsx"
    
    ' Fermer le classeur Excel
    xlBook.Close SaveChanges:=False
    
    ' Fermer l'application Excel
    xlApp.Quit
    
    ' Libérer la mémoire
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    
    ' Message de confirmation
    MsgBox "Les tâches ont été exportées avec succès vers ExportProject.xlsx."
End Sub

Sub DeleteFileIfExists(filePath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(filePath) Then
        fso.DeleteFile filePath
    End If

    Set fso = Nothing
End Sub