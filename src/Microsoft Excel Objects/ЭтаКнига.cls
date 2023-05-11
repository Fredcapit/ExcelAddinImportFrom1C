VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_AddinInstall()
    Call CheckInstall
    Call DeleteFromCellMenu
    Call AddToCellMenu
    
End Sub


Private Sub Workbook_AddinUninstall()
     Call DeleteFromCellMenu
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call DeleteFromCellMenu
End Sub

Private Sub Workbook_Open()
    Call CheckInstall
    Call DeleteFromCellMenu
    Call AddToCellMenu
End Sub
Private Sub CheckInstall()
    Dim oXL As Object, oAddin As Object
    Url = Me.Path & "\"
    normalUrl = Application.UserLibraryPath ' Environ("AppData") & "\Microsoft\AddIns"
    AddinTitle = Mid(Me.Name, 1, Len(Me.Name) - 5)

    If Url <> normalUrl Then
        If MsgBox("Надстройка Импорт из 1С не установлена. Установить?", vbYesNo) = vbYes Then
            Me.SaveCopyAs normalUrl & Me.Name
            Set oXL = Application ' CreateObject("Excel.Application")
            oXL.Workbooks.Add
            Set oAddin = oXL.AddIns.Add(normalUrl & Me.Name, True)
            oAddin.Installed = True

            oXL.Quit
            Set oXL = Nothing
            MsgBox "Надстройка установлена!"
        End If
    End If
End Sub
