Attribute VB_Name = "contextMenu1CImport"
Const documentFaceId As Integer = 777
Const removeFaceId As Integer = 214

Sub AddToCellMenu()
    Dim contextMenu As CommandBar

    '�������� ����������� ���� "����� �������"
    Set contextMenu = Application.CommandBars("List Range Popup")

    On Error GoTo ErrorHandle
        ' ��������� ������ ����, ���� ��� �����������
        
        If contextMenu.FindControl(Tag:="ImportFrom1C") Is Nothing Then
            With contextMenu.Controls.Add(Type:=msoControlButton, before:=1)
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "ImportFrom1CFile"
                .FaceId = documentFaceId
                .Caption = "������ �� 1�"
                .Tag = "ImportFrom1C"
            End With
        End If
        
        If contextMenu.FindControl(Tag:="removeMarked") Is Nothing Then
            With contextMenu.Controls.Add(Type:=msoControlButton, before:=2)
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "removeMarked"
                .FaceId = removeFaceId
                .Caption = "������� ���������� �� ��������"
                .Tag = "removeMarked"
            End With
            contextMenu.Controls(2).BeginGroup = True
        End If
        
    Exit Sub
ErrorHandle:
    Debug.Print Err.description
End Sub

Sub DeleteFromCellMenu()
   
    On Error GoTo ErrorHandle
    
        With Application.CommandBars("List Range Popup")
            .FindControl(Tag:="ImportFrom1C").Delete
            .FindControl(Tag:="removeMarked").Delete
        End With
    
    Exit Sub
    
ErrorHandle:
    Debug.Print Err.Source & " " & Err.description

End Sub

