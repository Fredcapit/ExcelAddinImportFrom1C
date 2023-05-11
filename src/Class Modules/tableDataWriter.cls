VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tableDataWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Автор: Шибанов Ф.В.
' mail: shibanovfv@yandex.ru
Private firstRowIsEmpty As Boolean
Public Function getTableOfCell(ByVal Target As Range) As ListObject
    
    Dim sht As Worksheet
    
    Set sht = Target.Parent
    
    For Each listObj In sht.ListObjects
        
        If Not Intersect(Target, listObj.Range) Is Nothing Then
            Set getTableOfCell = listObj
            Exit For
        End If
            
    Next listObj
    
End Function

Public Function VerifyTableHeaders(ByRef listObj As ListObject) As Boolean

    Const fieldsRequired As String = "ID,РасчСчет,Статус,ДатаСтатуса,СекцияДокумент,Дата,Номер,Сумма,ДатаПоступило,ДатаСписано,ПолучательРасчСчет,ПлательщикРасчСчет"
    Dim fieldsRequired_dict As Object
    Set fieldsRequired_dict = CreateObject("Scripting.Dictionary")
    

    
    Dim fieldRequired_name As Variant
    For Each fieldRequired_name In Split(fieldsRequired, ",")
        fieldsRequired_dict.Add fieldRequired_name, ""
    Next fieldRequired_name
    
    
    Dim header As Range
    For Each header In listObj.HeaderRowRange
        
        If fieldsRequired_dict.exists(header.value) Then fieldsRequired_dict.Remove header.value
        
    Next header
    
    If listObj.DataBodyRange Is Nothing Then
        listObj.ListRows.Add
        firstRowIsEmpty = True
    Else
        firstRowIsEmpty = False
    End If
    
    If fieldsRequired_dict.Count = 0 Then
        VerifyTableHeaders = True
    Else
        Dim missing_names As String
        Dim i As Integer
        For i = 0 To fieldsRequired_dict.Count - 1
            missing_names = missing_names & Chr(10) & Chr(13) & (i + 1) & ") " & fieldsRequired_dict.keys()(i)
        Next i
        If MsgBox("В таблице отсутствуют столбцы:" & Chr(10) & Chr(13) & missing_names & Chr(10) & Chr(13) & "Для продолжения их необходимо добавить", vbOKCancel) = vbOK Then
            For i = 0 To fieldsRequired_dict.Count - 1
                
                listObj.ListColumns.Add().Name = fieldsRequired_dict.keys()(i)
                
            Next i
            If fieldsRequired_dict.exists("ID") Then
                listObj.ListColumns("ID").DataBodyRange.FormulaR1C1 = "=CONCAT([@РасчСЧет],[@СекцияДокумент],[@Номер],[@Дата],[@ПлательщикРасчСчет])"
            End If
            If fieldsRequired_dict.exists("РасчСчет") Then
                listObj.ListColumns("РасчСчет").DataBodyRange.FormulaR1C1 = "=IF([@ДатаСписано]<>"""",[@ПлательщикРасчСчет],[@ПолучательРасчСчет])"
            End If
            
            VerifyTableHeaders = True
            
        End If
    End If

     listObj.ListColumns("Дата").DataBodyRange.NumberFormatLocal = "ДД.ММ.ГГГГ"
     listObj.ListColumns("ДатаСписано").DataBodyRange.NumberFormatLocal = "ДД.ММ.ГГГГ"
     listObj.ListColumns("ДатаПоступило").DataBodyRange.NumberFormatLocal = "ДД.ММ.ГГГГ"
     listObj.ListColumns("ДатаСтатуса").DataBodyRange.NumberFormatLocal = "ДД.ММ.ГГГГ чч:мм:сс"
     listObj.ListColumns("РасчСчет").DataBodyRange.NumberFormat = "@"
     listObj.ListColumns("ПлательщикРасчСчет").DataBodyRange.NumberFormat = "@"
     listObj.ListColumns("ПолучательРасчСчет").DataBodyRange.NumberFormat = "@"
     
End Function


Public Sub addNewRecords(ByRef docs As Object, ByRef list_object As ListObject)
    Dim lo As ListRow
    Dim lo_copy As ListRow
    
    Dim cell_range As Range
    
    Dim dateReceived As Range
    Dim datePayed As Range
    Dim doc As Object
    
    For Each doc In docs
        If firstRowIsEmpty Then
         Set lo = list_object.ListRows(1)
         firstRowIsEmpty = False
        Else
            Set lo = list_object.ListRows.Add()
        End If
        
        For Each field In doc.keys
            If doc.Item(field).filled And Not list_object.HeaderRowRange.Find(field, , xlValues, xlWhole) Is Nothing Then
                
                Set cell_range = lo.Range(1, list_object.ListColumns(field).Index)
                Select Case cell_range.NumberFormat
                    Case "General"
                        cell_range.Value2 = doc.Item(field).value
                    Case "@"
                        cell_range.Value2 = doc.Item(field).value
                    Case Else
                        
                        If IsDate(doc.Item(field).value) Then
                            cell_range.value = CDate(doc.Item(field).value)
                        Else
                            cell_range.value = doc.Item(field).value
                        End If
                End Select

            End If
        Next field
        lo.Range(1, list_object.ListColumns("Статус").Index).Value2 = "добавлено"
        lo.Range(1, list_object.ListColumns("ДатаСтатуса").Index).Value2 = Now()
        
        'В случае если поля ДатаСписано и ДатаПоступило указаны одновременно в документе, то разделяем строку на две, где одна имеет только ДатаПоступило, а
        ' другая только ДатаСписано с одинаковыми другими полями
        Set dateReceived = lo.Range(1, list_object.ListColumns("ДатаПоступило").Index)
        Set datePayed = lo.Range(1, list_object.ListColumns("ДатаСписано").Index)
        
        If dateReceived.Value2 <> "" And datePayed.Value2 <> "" Then
            Set lo_copy = list_object.ListRows.Add()
        
            For Each field In doc.keys
                If doc.Item(field).filled And Not list_object.HeaderRowRange.Find(field, , xlValues, xlWhole) Is Nothing Then
                    
                    Set cell_range = lo_copy.Range(1, list_object.ListColumns(field).Index)
                    Select Case cell_range.NumberFormat
                        Case "General"
                            cell_range.Value2 = doc.Item(field).value
                        Case "@"
                            cell_range.Value2 = doc.Item(field).value
                        Case Else
                            
                            If IsDate(doc.Item(field).value) Then
                                cell_range.value = CDate(doc.Item(field).value)
                            Else
                                cell_range.value = doc.Item(field).value
                            End If
                    End Select
                    lo_copy.Range(1, list_object.ListColumns("Статус").Index).Value2 = "добавлено"
                    lo_copy.Range(1, list_object.ListColumns("ДатаСтатуса").Index).Value2 = Now()
                End If
            Next field
            lo_copy.Range(1, list_object.ListColumns("Статус").Index).Value2 = "добавлено"
            lo_copy.Range(1, list_object.ListColumns("ДатаСтатуса").Index).Value2 = Now()
            
            dateReceived.Value2 = ""
            lo_copy.Range(1, list_object.ListColumns("ДатаСписано").Index).Value2 = ""
        End If
    
    Next doc
End Sub
Public Sub markForDeletion(ByVal account As Object, ByRef list_object As ListObject)
    'Чтобы выбрать данные из таблицы по условию быстро лучше всего использовать объект ADODB, который позволит сформировать sql запрос
    
    Dim conn As Object
    Dim rs As Object
    
    Dim connStr As String
    Dim sqlStr As String
    Dim bookFullName As String
    
    If InStr(1, ActiveWorkbook.FullName, "https://") Then
        bookFullName = getLocalFileName(ActiveWorkbook.FullName)
    Else
        bookFullName = ActiveWorkbook.FullName
    End If
    
    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & bookFullName
    connStr = connStr & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"
    Set conn = CreateObject("ADODB.Connection")
    conn.Open connStr
    
    Dim sht_name As String
    Dim sht As Worksheet
    Set sht = list_object.Parent
    sht_name = sht.Name
    
    sqlStr = "SELECT DISTINCT ID FROM [" & sht_name & "$" & list_object.Range.Address(0, 0) & "] " & _
            "WHERE РасчСчет='" & account.Item("РасчСчет").value & "' " & _
            " and (ДатаСписано between #" & Format(CDate(account("ДатаНачала").value), "YYYY-MM-dd") & "# " & _
            "       and #" & Format(CDate(account("ДатаКонца").value), "YYYY-MM-dd") & "# " & _
            " or ДатаПоступило between #" & Format(CDate(account("ДатаНачала").value), "YYYY-MM-dd") & "# " & _
            " and #" & Format(CDate(account("ДатаКонца").value), "YYYY-MM-dd") & "# )"
    
    Set rs = CreateObject("ADODB.Recordset")
    
    Dim records As Variant
    
    rs.Open sqlStr, conn
    If Not rs.EOF Then
        records = rs.GetRows()
    End If
    rs.Close
    conn.Close
    
    If Not IsEmpty(records) Then
        For i = 0 To UBound(records, 2)
            Call removeRecord(records(0, i), list_object)
        Next i
    End If
End Sub
Private Sub removeRecord(ByVal ID As String, ByRef list_object As ListObject)

    Dim found_row As Range
    Set found_row = list_object.ListColumns("ID").Range.Find(ID, , xlValues, xlWhole)
    Dim first_found As Range
    Set first_found = found_row
    
    Dim record_state As Range
    Dim record_state_date As Range
    Dim record As Range
    
    If Not found_row Is Nothing Then
        Set record = list_object.ListRows(found_row.row - list_object.HeaderRowRange.row).Range
        Set record_state = record(1, list_object.ListColumns("Статус").Index)
        Set record_state_date = record(1, list_object.ListColumns("ДатаСтатуса").Index)
        
        If record_state.Value2 <> "Удалено" Then
            record_state.Value2 = "Удалено"
            record_state_date.Value2 = Now()
        End If
        'Продолжаем поиск
        Set found_row = list_object.ListColumns("ID").Range.FindNext(found_row)
        Do While Not found_row Is Nothing And found_row.Address <> first_found.Address
            Set record = list_object.ListRows(found_row.row - list_object.HeaderRowRange.row).Range
            Set record_state = record(1, list_object.ListColumns("Статус").Index)
            Set record_state_date = record(1, list_object.ListColumns("ДатаСтатуса").Index)
            
            If record_state.Value2 <> "Удалено" Then
                record_state.Value2 = "Удалено"
                record_state_date.Value2 = Now()
            End If
            Set found_row = list_object.ListColumns("ID").Range.FindNext(found_row)
        Loop
        
    End If
    
    
End Sub
Public Sub removeMarked(ByRef list_object As ListObject)
    Dim found_row As Range
    If list_object.HeaderRowRange.Find("Статус", , xlValues, xlWhole) Is Nothing Then
        Exit Sub
    End If
    
    Set found_row = list_object.ListColumns("Статус").Range.Find("Удалено", , xlValues, xlWhole)
    Dim first_found As Range
    Set first_found = found_row
    
    Dim record_state As Range
    Dim record_state_date As Range
    Dim record As Range
    
    If Not found_row Is Nothing Then
        
        list_object.ListRows(found_row.row - list_object.HeaderRowRange.row).Delete
        
        'Продолжаем поиск
        Set found_row = list_object.ListColumns("Статус").Range.Find("Удалено", , xlValues, xlWhole)
        Do While Not found_row Is Nothing
            list_object.ListRows(found_row.row - list_object.HeaderRowRange.row).Delete
            Set found_row = list_object.ListColumns("Статус").Range.Find("Удалено", , xlValues, xlWhole)
        Loop
        
    End If
End Sub
Private Function getLocalFileName(ByVal FullName As String) As String

    Dim regExp As Object
    Set regExp = CreateObject("VBScript.RegExp")
    
    Dim oneDrivePattern As String
    Dim corpOneDrivePattern As String
    
    oneDrivePattern = "^(https:\/\/d.docs.live.net\/\w+\/)(.*\.(xlsx|xlsb|xlsm))(.*)?$"
    corpOneDrivePattern = "^(https:\/\/.*\.sharepoint.com\/personal\/\w+\/Documents\/)(.*\.(xlsx|xlsm|xlsb))(.*)$"
    
     With regExp
        .Global = True
        .IgnoreCase = True
        .MultiLine = False
    End With
    
    regExp.pattern = oneDrivePattern
    Dim matches As Object
    If regExp.test(FullName) Then
        Set matches = regExp.Execute(FullName)
        getLocalFileName = Environ("OneDrive") & "\" & matches(0).submatches(1)
    End If
    
    regExp.pattern = corpOneDrivePattern
    If regExp.test(FullName) Then
        Set matches = regExp.Execute(FullName)
        getLocalFileName = Environ("OneDriveCommercial") & "\" & matches(0).submatches(1)
        
    End If
    
End Function

