VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parser1C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' �����: ������� �.�.
' mail: shibanovfv@yandex.ru

Public parsedFiles As Collection


' ���������� ��������� ��������� ����� �� ���������

Private firstRow As String
Private formatVersion As String
Private generalInfo As Object
Private selectionInfo As Object
Private accountSections As Collection
Private docSections As Collection

' ���������� ����������� ��� ������ ������ �����

Private fileSystem As Object
Private file_name As String
Private textStream As Object


' ���������� ��� ������������� ��������

Private Const generalInfoFields = "�������������,���������,�����������,����������,������������,�������������"
 
' ���������� ��� ������ ������ ������

Private key As String, value As String
Private row_index As Integer
Const FileEOF As String = "EndOfStream"

' �������������� ������ � ����� ������

Private Const accTrigger_begin = "��������������"
Private Const accTrigger_finish = "�������������"

Private Const docTrigger_begin = "��������������"
Private Const docTrigger_finish = "��������������"

' ������� ������ ���������� ������ ������ ������ �����

Private cancelled As Boolean

' ��������� ������ ������ �����

Private errors As Collection

Private Sub Class_Initialize()
    
    Set parsedFiles = New Collection
    
    Call generalInfo_Init
    Call selectionInfo_Init
    
    Set accountSections = New Collection
    Set docSections = New Collection
    Set errors = New Collection
    
    cancelled = False
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    
End Sub
Public Sub Run()

    Dim files As FileDialogSelectedItems
    Set files = Get_file_name()
    If files.Count = 0 Then Exit Sub

    Dim data As parsed1CData
    
    For Each file In files
        Set data = New parsed1CData
        Call Clear
        Call validatelog("====================================")
        Call Read(file)
        With data
            .firstRow = firstRow
            .formatVersion = formatVersion
            Set .generalInfo = generalInfo
            Set .selectionInfo = selectionInfo
            Set .accountSections = accountSections
            Set .docSections = docSections
            Set .errors = errors
        End With
        parsedFiles.Add data
        
    Next file
    
End Sub
Private Sub Validate()

    validatelog "������ ���������"
    
    Call validateSection(generalInfo)
    Call validateSection(selectionInfo)
    
    Dim account As Object
    For Each account In accountSections
    
        Call validateSection(account)
        
    Next account
    
    Dim doc As Object
    For Each doc In docSections
    
        Call validateSection(doc)
        
    Next doc
    
    validatelog "����� ���������"
    
End Sub
Private Sub validateSection(ByRef section As Object)

    Dim field As DataField
    
    For i = 0 To section.Count - 1
    
        Set field = section.Items()(i)
        If field.required And Not field.filled _
            And field.key <> "�������������" And field.key <> "�����������" Then
            
            validatelog "���� " & field.key & " �� ���������", field
            
        End If
        
        If field.key = "�������������" Or field.key = "�����������" Then
        
            If field.key = "�������������" And Not field.filled And Not section.Item("�����������").filled Then
            
                validatelog "���� " & field.key & " �� ���������", field
                
            End If
            
            If field.key = "�����������" And Not field.filled And Not section.Item("�������������").filled Then
            
                validatelog "���� " & field.key & " �� ���������", field
                
            End If
            
        End If
        
    Next i
    
End Sub
Private Static Sub generalInfo_Init()

    Set generalInfo = CreateObject("Scripting.Dictionary")
    
    Dim fields
    fields = Array( _
                    Array("�������������", True, "^1.(01|02|03)$", "����� ������ ������� ������"), _
                    Array("���������", True, "^(DOS|Windows)$", "��������� �����"), _
                    Array("�����������", False, "", "���������-�����������"), _
                    Array("����������", True, "", "���������-����������"), _
                    Array("������������", False, "", "���� ������������ �����"), _
                    Array("�������������", False, "", "����� ������������ �����") _
                    )
                    
    Dim dfield As DataField
    For Each field In fields
    
        Set dfield = CreateDataField(field(0), field(1), field(2), field(3))
        dfield.section = "�������������"
        
        generalInfo.Add field(0), dfield
        
    Next field
    
End Sub
Private Static Sub selectionInfo_Init()

    Set selectionInfo = CreateObject("Scripting.Dictionary")
    
    Dim fields
    fields = Array( _
                    Array("����������", True, "^\d{1,2}\.\d{1,2}\.\d{4}$", "���� ������ ���������"), _
                    Array("���������", True, "^\d{1,2}\.\d{1,2}\.\d{4}$", "���� ����� ���������"), _
                    Array("��������", True, "^\d{20}$", "��������� ���� �����������") _
                    )
                    
    Dim dfield As DataField
    For Each field In fields
    
        Set dfield = CreateDataField(field(0), field(1), field(2), field(3))
        dfield.section = "�������������"
        
        selectionInfo.Add field(0), dfield
        
    Next field

End Sub
Private Sub NextLine()
    
    If Not textStream.AtEndOfStream Then
        Call parseKeyValue(textStream.ReadLine)
        row_index = row_index + 1
        
    Else
        key = "EndOfStream"
    End If
    
End Sub
Private Sub Read(ByVal fileName As String)

    file_name = fileName
    If file_name = "" Then
        MsgBox "������. ��� ����� �� �������", vbCritical
        Exit Sub
    End If
    
    Set textStream = fileSystem.OpenTextFile(file_name, 1, False, -2)
    

    Call ReadHead
    If cancelled Then Exit Sub
    
    Call ReadGeneralInfo
    If cancelled Then Exit Sub
    
    Call ReadSelectionInfo
    If cancelled Then Exit Sub

    Do While Not textStream.AtEndOfStream And Not cancelled
        Call ReadAccountSection
        Call ReadDocumentSection
        If key = "����������" Then
            Call NextLine
        ElseIf key <> accTrigger_begin And key <> docTrigger_begin And Not textStream.AtEndOfStream Then
            log "����������� �������������, ���� ������������� � ����������� ����� key=" & key & " value=" & value
            Call NextLine
        End If
    Loop
    
    Call Normalize
    Call Validate
End Sub

Private Sub ReadHead()
    
    Call NextLine
    
    If key = FileEOF Then
        log "���� ������"
        cancelled = True
    End If
    
    If key <> "1CClientBankExchange" Then
        log "�� ���������� ������ �������� ����� 1CClientBankExchange"
        If MsgBox("�� ��������� ������� ����� 1CClientBankExchange. ���������� ������ ����� � ����� ������?", vbYesNo, "������ ������ ����� " & file_name) = vbNo Then
            cancelled = True
        End If
    End If
   
End Sub


Private Sub ReadGeneralInfo()

    Call NextLine
    If key = FileEOF Then
        log "����������� ����� �����"
        Exit Sub
    End If
    
    Do While generalInfo.exists(key)
    
        generalInfo.Item(key).value = value
        generalInfo.Item(key).row_index = row_index
        
        If (Not generalInfo.Item(key).testValue) Then
            log "������: ""����� ��������"" �������������: " & key & " ����� �������� ��������", generalInfo.Item(key)
        End If
        
        Call NextLine
        
    Loop
    
    fileVersion = generalInfo.Item("�������������").value
    
End Sub
Private Sub ReadSelectionInfo()
    
    Do While selectionInfo.exists(key)
    
        selectionInfo.Item(key).value = value
        selectionInfo.Item(key).row_index = row_index
        
        If (Not selectionInfo.Item(key).testValue) Then
            log "������: ""������� ������"" �������������: " & key & " ����� �������� ��������", selectionInfo.Item(key)
        End If
        
        Call NextLine
        
    Loop
    
End Sub
Private Function CreateAccountSection() As Object

    Set CreateAccountSection = CreateObject("Scripting.Dictionary")
    
    Dim fields
    fields = Array( _
                    Array("����������", True, "^\d{1,2}\.\d{1,2}\.\d{4}$", "���� ������ ���������"), _
                    Array("���������", True, "^\d{1,2}\.\d{1,2}\.\d{4}$", "���� ����� ���������"), _
                    Array("��������", True, "^\d{20}$", "��������� ���� �����������"), _
                    Array("����������������", True, "^\d+(\.\d{1,2})?$", ""), _
                    Array("��������������", True, "^\d+(\.\d{1,2})?$", ""), _
                    Array("������������", True, "^\d+(\.\d{1,2})?$", ""), _
                    Array("���������������", True, "^\d+(\.\d{1,2})?$", "") _
                    )
                    
    Dim dfield As DataField
    For Each field In fields
    
        Set dfield = CreateDataField(field(0), field(1), field(2), field(3))
        dfield.section = "�����������������"
        
        CreateAccountSection.Add field(0), dfield
        
    Next field
    
End Function
Private Function CreateDocSection() As Object

    Set CreateDocSection = CreateObject("Scripting.Dictionary")
    
    Dim fields As New Collection
    fields.Add Array("��������������", True, "", "������� ������ ������, �������� ��� ���������")
    fields.Add Array("�����", True, "", "����� ���������")
    fields.Add Array("����", True, "^\d{1,2}\.\d{1,2}\.\d{4}$", "���� ���������")
    fields.Add Array("�����", True, "^\d+(\.\d{1,2})?$", "����� �������")
    fields.Add Array("�������������", False, "^\d{1,2}\.\d{1,2}\.\d{4}$", "���� ������������ ���������")
    fields.Add Array("��������������", False, "^\d{1,2}\:\d{1,2}\:\d{1,2}$", "����� ������������ ���������")
    fields.Add Array("�������������������", False, "", "���������� ���������")
    fields.Add Array("��������������", True, "^\d{20}$", "��������� ���� �����������")
    fields.Add Array("�����������", True, "^\d{1,2}\.\d{1,2}\.\d{4}$", "���� �������� ������� � �/�")
    fields.Add Array("����������", False, "", "��� � ������������ �����������")
    fields.Add Array("�������������", True, "^\d{10,12}$", "��� �����������")
    fields.Add Array("����������1", False, "", "������������ �����������")
    fields.Add Array("����������2", False, "", "��������� ���� �����������")
    fields.Add Array("����������3", False, "", "���� �����������")
    fields.Add Array("����������4", False, "", "����� ����� �����������")
    fields.Add Array("������������������", False, "^\d{20}$", "������� ����� �����������")
    fields.Add Array("��������������1", False, "", "�� ����� �����������")
    fields.Add Array("��������������2", False, "", "����� �� ����� �����������")
    fields.Add Array("�������������", False, "^\d{9}$", "��� �� ����� �����������")
    fields.Add Array("�����������������", False, "^\d{20}$", "������� �� ����� �����������")
    fields.Add Array("��������������", True, "^\d{20}$", "��������� ���� ����������")
    fields.Add Array("�������������", True, "^\d{1,2}\.\d{1,2}\.\d{4}$", "���� ����������� ������� �� �/�")
    fields.Add Array("����������", False, "", "���������� ����� �������")
    fields.Add Array("�������������", True, "^\d{10,12}$", "��� ����������")
    fields.Add Array("����������1", False, "", "������������ ����������")
    fields.Add Array("����������2", False, "", "��������� ���� ����������")
    fields.Add Array("����������3", False, "", "���� ����������")
    fields.Add Array("����������4", False, "", "����� ����� ����������")
    fields.Add Array("������������������", False, "^\d{20}$", "������� ����� ����������")
    fields.Add Array("��������������1", False, "", "�� ����� ����������")
    fields.Add Array("��������������2", False, "", "����� �� ����� ����������")
    fields.Add Array("�������������", False, "^\d{9}$", "��� �� ����� ����������")
    fields.Add Array("�����������������", False, "^\d{20}$", "������� �� ����� ����������")
    fields.Add Array("����������", False, "", "��� �������")
    fields.Add Array("�������������", False, "", "���������� ������� �������")
    fields.Add Array("���������", False, "^\d{1,2}$", "�������� �������� ����������� ���������, �������� �������������� ���")
    fields.Add Array("���", False, "^.{0,25}$", "���������� ������������� �������")
    fields.Add Array("�����������������", False, "", "���������� ������� ����� �������")
    fields.Add Array("����������������� 1", False, "", "���������� ������� �������� �� ������")
    fields.Add Array("����������������� 2", False, "", "���������� ������� �������� �� ������")
    fields.Add Array("����������������� 3", False, "", "���������� ������� �������� �� ������")
    fields.Add Array("����������������� 4", False, "", "���������� ������� �������� �� ������")
    fields.Add Array("����������������� 5", False, "", "���������� ������� �������� �� ������")
    fields.Add Array("����������������� 6", False, "", "���������� ������� �������� �� ������")
    fields.Add Array("�����������������", False, "^\d{2}$", "������ ����������� ���������� ���������")
    fields.Add Array("�������������", False, "^(\d{9}|0)$", "��� �����������")
    fields.Add Array("�������������", False, "^(\d{9}|0)$", "��� ����������")
    fields.Add Array("�������������", False, "^(\d{20}|0)$", "���������� ���� ��������� �������������")
    fields.Add Array("�����", False, "^(\d{11}|0)$", "��� ����� ����������, �� ������� �������������� �������� �������� �� ������ ������, ����� � ����� �������")
    fields.Add Array("�������������������", False, "", "^(\d{2}|0)$")
    fields.Add Array("�����������������", False, "^(\d{2}|0)$", "���������� ���������� ������� / ��� ����������� ������")
    fields.Add Array("����������������", False, "", "���������� ������ ���������")
    fields.Add Array("��������������", False, "^\d{1,2}\.\d{1,2}\.\d{4}|0$", "���������� ���� ���������")
    fields.Add Array("�����������", False, "^\d{0,2}$", "����������� �������")
    fields.Add Array("�����������", False, "^\d+$", "���� �������, ���������� ����")
    fields.Add Array("��������������", False, "", "��� �����������")
    fields.Add Array("�����������", False, "^\d{1,2}\.\d{1,2}\.\d{4}$", "���� ������� (�����������)")
    fields.Add Array("������������1", False, "", "������� ������, ���.1")
    fields.Add Array("������������2", False, "", "������� ������, ���.2")
    fields.Add Array("������������3", False, "", "������� ������, ���.3")
    fields.Add Array("��������������", False, "", "������ �� ��������������")
    fields.Add Array("�������������", False, "", "�������������� �������")
    fields.Add Array("�������������������", False, "", "����� ����� ����������")
    fields.Add Array("��������������", False, "^\d{1,2}\.\d{1,2}\.\d{4}$", "���� ������� ����������")

    Dim dfield As DataField
    For Each field In fields
    
        Set dfield = CreateDataField(field(0), field(1), field(2), field(3))
        dfield.section = "�����������������"
        
        CreateDocSection.Add field(0), dfield
        
    Next field
    
End Function
Private Sub ReadAccountSection()

    If key = accTrigger_begin Then
    
        Call NextLine
        If key = FileEOF Then
            log "����������� ����� ����� ��������������"
            Exit Sub
        End If
        
        Dim accSection As Object
        Set accSection = CreateAccountSection()
        If accSection.exists(getkey()) Then

            Do While Not textStream.AtEndOfStream And key <> accTrigger_finish And key <> docTrigger_begin And key <> "����������"
                If accSection.exists(key) Then
                    Dim accField As DataField
                    Set accField = accSection.Item(key)
                    accField.value = value
                    accField.row_index = row_index
                    
                    If (Not accField.testValue) Then
                        log " �������������: " & key & " ����� �������� ��������", accField
                    End If
                    Set accField = Nothing
                End If
                Call NextLine
                
            Loop
            
            accountSections.Add accSection
            
            If key <> accTrigger_finish Then
                log "�� ��������� ������� ����� ������ ��������������"
            ElseIf key = accTrigger_finish Then
                NextLine
            End If
            
        Else
        
            log "�� ���������� ������ ������ ��������������"
            Set accSection = Nothing
            
        End If
        
    End If
    
    
End Sub
Private Function getValue() As String
    getValue = value
End Function
Private Function getkey() As String
    getkey = key
End Function

Private Sub ReadDocumentSection()

    If key = docTrigger_begin Then
        
        Dim docSection As Object
        Set docSection = CreateDocSection()
        
        Do While Not textStream.AtEndOfStream And key <> docTrigger_finish And key <> "����������"
           If docSection.exists(key) Then
                Dim docField As DataField
                Set docField = docSection.Item(key)
                docField.value = value
                docField.row_index = row_index
                If (Not docField.testValue) Then
                    log " �������������: " & key & " ����� �������� ��������", docField
                End If
                Set docField = Nothing
            End If
            Call NextLine
            
        Loop
        
        docSections.Add docSection
            
        If key <> docTrigger_finish Then
            log "�� ��������� ������� ����� ������ ��������������"
        ElseIf key = docTrigger_finish Then
            NextLine
        End If
        
    End If
End Sub
Private Static Function CreateDataField(ByVal key As String, ByVal required As Boolean, ByVal pattern As String, ByVal description As String) As DataField

    Set CreateDataField = New DataField
    With CreateDataField
        .key = key
        .required = required
        .valuePattern = pattern
        .description = description
    End With
    
End Function
Private Sub parseKeyValue(ByVal row As String)

    Dim equalSymbolIndex As Integer
    equalSymbolIndex = InStr(1, row, "=", vbTextCompare)
    
    If equalSymbolIndex = 0 Then
        key = row
        value = ""
    Else
        key = Mid(row, 1, equalSymbolIndex - 1)
        value = Mid(row, equalSymbolIndex + 1)
    End If
    
End Sub
Private Sub log(ByVal message, Optional ByRef field As DataField)

    Dim Error_string As String
    
    If field Is Nothing Then
        Error_string = file_name & " ������: " & row_index & " [���������] " & message
    Else
        Error_string = file_name & " ������: " & field.row_index & " [���������] " & field.section & " " & message
    End If
    
    errors.Add Error_string

    
End Sub
Private Sub validatelog(ByVal message, Optional ByRef field As DataField)
    Dim Error_string As String
    If field Is Nothing Then
        Error_string = file_name & " [���������] " & message
    Else
        Error_string = file_name & " " & field.row_index & " [���������] " & field.section & " " & message
    End If
    errors.Add Error_string

    
End Sub
Private Sub Normalize()
    
    Dim doc As Object
    validatelog "������������ ������"
    For Each doc In docSections
    
        Call NormalizePaymentDetails(doc)
        Call NormalizeBankAccounts(doc)
        Call NormalizeCompanyName(doc)
        
    Next doc
    validatelog "����� ������������ ������"
End Sub
Private Sub NormalizePaymentDetails(ByRef doc As Object)

    Dim paymentDetails As DataField
    Set paymentDetails = doc.Item("�����������������")
    
    If Not paymentDetails.filled Or paymentDetails.value = "" Then
        log "����������������� ������. ������� ���������� ������������� ����.", paymentDetails
        paymentDetails.value = paymentDetail.value & doc.Item("�����������������1").value
        paymentDetails.value = paymentDetail.value & doc.Item("�����������������2").value
        paymentDetails.value = paymentDetail.value & doc.Item("�����������������3").value
        paymentDetails.value = paymentDetail.value & doc.Item("�����������������4").value
        paymentDetails.value = paymentDetail.value & doc.Item("�����������������5").value
        paymentDetails.value = paymentDetail.value & doc.Item("�����������������6").value
        If paymentDetail.value <> "" Then
            log "����������������� ������� ������� �� ���������� ����������� �����", paymentDetails
        Else
            log "����������������� �� ������� ����������", paymentDetails
        End If
    End If

End Sub

Private Sub NormalizeBankAccounts(ByRef doc As Object)

    If Not doc.Item("��������������").filled Then
        If Not doc.Item("����������2").filled Then
            doc.Item("��������������").value = doc.Item("������������������").value
            log "���������� ���� �������������� �� �������� ��������� ���� ������������������", doc.Item("��������������")
        Else
            doc.Item("��������������").value = doc.Item("����������2").value
            log "���������� ���� �������������� �� �������� ��������� ���� ����������2", doc.Item("��������������")
        End If
    End If
    If Not doc.Item("��������������").filled Then
        If Not doc.Item("����������2").filled Then
            doc.Item("��������������").value = doc.Item("������������������").value
            log "���������� ���� �������������� �� �������� ��������� ���� ������������������", doc.Item("��������������")
        Else
            doc.Item("��������������").value = doc.Item("����������2").value
             log "���������� ���� �������������� �� �������� ��������� ���� ����������2", doc.Item("��������������")
        End If
    End If
    
End Sub
Private Sub NormalizeCompanyName(ByRef doc As Object)

    Dim matches As Object
    Dim regExp As Object
    Set regExp = CreateObject("VBScript.RegExp")
    
    If Not doc.Item("����������1").filled Then
       
        With regExp
            .Global = True
            .IgnoreCase = True
            .MultiLine = False
            .pattern = "^(\d{10,12})( )(.+)$"
        End With
        
        If regExp.test(doc.Item("����������").value) Then
            
            Set matches = regExp.Execute(doc.Item("����������").value)
            doc.Item("����������1").value = Trim(matches(0).submatches(2))
                log "���������� ���� ����������1 �� ���� ����������", doc.Item("����������")
            If Not doc.Item("�������������").filled Then
                doc.Item("�������������").value = Trim(matches(0).submatches(0))
                log "���������� ���� ������������� �� ���� ����������", doc.Item("����������")
            End If
            
        End If
        
    End If
    
    If Not doc.Item("����������1").filled Then

        With regExp
            .Global = True
            .IgnoreCase = True
            .MultiLine = False
            .pattern = "^(\d{10,12})( )(.+)$"
        End With
        
        If regExp.test(doc.Item("����������").value) Then
            Set matches = regExp.Execute(doc.Item("����������").value)
            doc.Item("����������1").value = Trim(matches(0).submatches(2))
            log "���������� ���� ����������1 �� ���� ����������", doc.Item("����������")
            If Not doc.Item("�������������").filled Then
                doc.Item("�������������").value = Trim(matches(0).submatches(0))
                log "���������� ���� ������������� �� ���� ����������", doc.Item("����������")
            End If
        End If
        
    End If
    
End Sub
Private Sub Clear()
    

    firstRow = ""
    formatVersion = ""
    Set generalInfo = Nothing
    Call generalInfo_Init
    Set selectionInfo = Nothing
    Call selectionInfo_Init
    
    Set accountSections = Nothing
    Set accountSections = New Collection
    
    Set docSections = Nothing
    Set docSections = New Collection
    
    file_name = ""
    If Not textStream Is Nothing Then textStream.Close
    
    key = ""
    value = ""
    row_index = 0
    
    cancelled = False
    
    Set errors = Nothing
    Set errors = New Collection
    
End Sub
Private Sub Class_Terminate()

    If Not textStream Is Nothing Then textStream.Close
    Set textStream = Nothing
    Set fileSystem = Nothing
    
End Sub
Private Function Get_file_name() As FileDialogSelectedItems

    Dim file_dialog As FileDialog
    Set file_dialog = Application.FileDialog(msoFileDialogFilePicker)
   
    With file_dialog
        .AllowMultiSelect = True
        .Filters.Add "���������� �������", "*.1c,*.txt", 1
       
    End With
    file_dialog.Show
    Set Get_file_name = file_dialog.SelectedItems

End Function
