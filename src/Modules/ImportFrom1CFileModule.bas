Attribute VB_Name = "ImportFrom1CFileModule"
' Автор: Шибанов Ф.В.
' mail: shibanovfv@yandex.ru

Public Sub ImportFrom1CFile()
    

    If Not isWorkbookSaved() Then
         MsgBox "Для продолжения импорта необходимо сохранить книгу. Сохраните книгу и повторите импорт", vbInformation
         Exit Sub
    End If
    Dim tableWriter As New tableDataWriter
    Dim targetTable As ListObject
    Set targetTable = tableWriter.getTableOfCell(ActiveCell)
    On Error Resume Next
    If Not targetTable Is Nothing Then
        If tableWriter.VerifyTableHeaders(targetTable) Then
                    
                    Dim parser As New Parser1C
                    parser.Run
                    
                    Dim parsedFile As parsed1CData
                    
                    Dim ErrorsForm As New Import1CFileErrorsForm
                    Dim errors As String
                    Application.ScreenUpdating = False
                    For Each parsedFile In parser.parsedFiles
                        If parsedFile.errors.Count > 0 Then
                            ErrorsForm.Show vbModeless
                            
                            Dim error As Variant
                            For Each error In parsedFile.errors
                                ErrorsForm.txtbox_output.Text = ErrorsForm.txtbox_output.Text & Chr(10) & Chr(13) & error
                            Next error
                        End If
                        
                        Dim account As Object
                        For Each account In parsedFile.accountSections
                            DoEvents
                            Call tableWriter.markForDeletion(account, targetTable)
                        Next account
                        Call tableWriter.addNewRecords(parsedFile.docSections, targetTable)
                        DoEvents
                  
                        
                    Next parsedFile
                    Application.ScreenUpdating = True
                    MsgBox "Импорт завершен!", vbApplicationModal, "Импорт из файлов 1С"
                    
        End If
        
    End If
End Sub
Public Sub removeMarked()
    Dim tableWriter As New tableDataWriter
    Call tableWriter.removeMarked(tableWriter.getTableOfCell(ActiveCell))
End Sub
Private Function isWorkbookSaved() As Boolean
    isWorkbookSaved = True
    On Error Resume Next
    Dim result As Variant
    result = ActiveWorkbook.BuiltinDocumentProperties(12).value
    If Err.Number <> 0 Then
        isWorkbookSaved = False
    End If
    On Error GoTo 0
End Function
