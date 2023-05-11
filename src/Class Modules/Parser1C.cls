VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parser1C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Автор: Шибанов Ф.В.
' mail: shibanovfv@yandex.ru

Public parsedFiles As Collection


' Переменные структуры читаемого файла по стандарту

Private firstRow As String
Private formatVersion As String
Private generalInfo As Object
Private selectionInfo As Object
Private accountSections As Collection
Private docSections As Collection

' Переменные необходимые для чтения самого файла

Private fileSystem As Object
Private file_name As String
Private textStream As Object


' Переменная для инициализации объектов

Private Const generalInfoFields = "ВерсияФормата,Кодировка,Отправитель,Получатель,ДатаСоздания,ВремяСоздания"
 
' Переменные для чтения каждой строки

Private key As String, value As String
Private row_index As Integer
Const FileEOF As String = "EndOfStream"

' Идентификаторы начала и конца секций

Private Const accTrigger_begin = "СекцияРасчСчет"
Private Const accTrigger_finish = "КонецРасчСчет"

Private Const docTrigger_begin = "СекцияДокумент"
Private Const docTrigger_finish = "КонецДокумента"

' Признак отмены выполнения полной логики чтения файла

Private cancelled As Boolean

' Коллекция ошибок чтения файла

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

    validatelog "Начало валидации"
    
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
    
    validatelog "Конец валидации"
    
End Sub
Private Sub validateSection(ByRef section As Object)

    Dim field As DataField
    
    For i = 0 To section.Count - 1
    
        Set field = section.Items()(i)
        If field.required And Not field.filled _
            And field.key <> "ДатаПоступило" And field.key <> "ДатаСписано" Then
            
            validatelog "Поле " & field.key & " не заполнено", field
            
        End If
        
        If field.key = "ДатаПоступило" Or field.key = "ДатаСписано" Then
        
            If field.key = "ДатаПоступило" And Not field.filled And Not section.Item("ДатаСписано").filled Then
            
                validatelog "Поле " & field.key & " не заполнено", field
                
            End If
            
            If field.key = "ДатаСписано" And Not field.filled And Not section.Item("ДатаПоступило").filled Then
            
                validatelog "Поле " & field.key & " не заполнено", field
                
            End If
            
        End If
        
    Next i
    
End Sub
Private Static Sub generalInfo_Init()

    Set generalInfo = CreateObject("Scripting.Dictionary")
    
    Dim fields
    fields = Array( _
                    Array("ВерсияФормата", True, "^1.(01|02|03)$", "Номер версии формата обмена"), _
                    Array("Кодировка", True, "^(DOS|Windows)$", "Кодировка файла"), _
                    Array("Отправитель", False, "", "Программа-отправитель"), _
                    Array("Получатель", True, "", "Программа-получатель"), _
                    Array("ДатаСоздания", False, "", "Дата формирования файла"), _
                    Array("ВремяСоздания", False, "", "Время формирования файла") _
                    )
                    
    Dim dfield As DataField
    For Each field In fields
    
        Set dfield = CreateDataField(field(0), field(1), field(2), field(3))
        dfield.section = "УсловияОтбора"
        
        generalInfo.Add field(0), dfield
        
    Next field
    
End Sub
Private Static Sub selectionInfo_Init()

    Set selectionInfo = CreateObject("Scripting.Dictionary")
    
    Dim fields
    fields = Array( _
                    Array("ДатаНачала", True, "^\d{1,2}\.\d{1,2}\.\d{4}$", "Дата начала интервала"), _
                    Array("ДатаКонца", True, "^\d{1,2}\.\d{1,2}\.\d{4}$", "Дата конца интервала"), _
                    Array("РасчСчет", True, "^\d{20}$", "Расчетный счет организации") _
                    )
                    
    Dim dfield As DataField
    For Each field In fields
    
        Set dfield = CreateDataField(field(0), field(1), field(2), field(3))
        dfield.section = "УсловияОтбора"
        
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
        MsgBox "Ошибка. Имя файла не указано", vbCritical
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
        If key = "КонецФайла" Then
            Call NextLine
        ElseIf key <> accTrigger_begin And key <> docTrigger_begin And Not textStream.AtEndOfStream Then
            log "Неизвестный идентификатор, либо идентификатор в неожиданном месте key=" & key & " value=" & value
            Call NextLine
        End If
    Loop
    
    Call Normalize
    Call Validate
End Sub

Private Sub ReadHead()
    
    Call NextLine
    
    If key = FileEOF Then
        log "Файл пустой"
        cancelled = True
    End If
    
    If key <> "1CClientBankExchange" Then
        log "Не обнаружена строка признака файла 1CClientBankExchange"
        If MsgBox("Не обнаружен признак файла 1CClientBankExchange. Продолжить чтение файла в любом случае?", vbYesNo, "Ошибка чтения файла " & file_name) = vbNo Then
            cancelled = True
        End If
    End If
   
End Sub


Private Sub ReadGeneralInfo()

    Call NextLine
    If key = FileEOF Then
        log "Неожиданный конец файла"
        Exit Sub
    End If
    
    Do While generalInfo.exists(key)
    
        generalInfo.Item(key).value = value
        generalInfo.Item(key).row_index = row_index
        
        If (Not generalInfo.Item(key).testValue) Then
            log "Секция: ""Общие сведения"" Идентификатор: " & key & " имеет неверное значение", generalInfo.Item(key)
        End If
        
        Call NextLine
        
    Loop
    
    fileVersion = generalInfo.Item("ВерсияФормата").value
    
End Sub
Private Sub ReadSelectionInfo()
    
    Do While selectionInfo.exists(key)
    
        selectionInfo.Item(key).value = value
        selectionInfo.Item(key).row_index = row_index
        
        If (Not selectionInfo.Item(key).testValue) Then
            log "Секция: ""Условия отбора"" Идентификатор: " & key & " имеет неверное значение", selectionInfo.Item(key)
        End If
        
        Call NextLine
        
    Loop
    
End Sub
Private Function CreateAccountSection() As Object

    Set CreateAccountSection = CreateObject("Scripting.Dictionary")
    
    Dim fields
    fields = Array( _
                    Array("ДатаНачала", True, "^\d{1,2}\.\d{1,2}\.\d{4}$", "Дата начала интервала"), _
                    Array("ДатаКонца", True, "^\d{1,2}\.\d{1,2}\.\d{4}$", "Дата конца интервала"), _
                    Array("РасчСчет", True, "^\d{20}$", "Расчетный счет организации"), _
                    Array("НачальныйОстаток", True, "^\d+(\.\d{1,2})?$", ""), _
                    Array("ВсегоПоступило", True, "^\d+(\.\d{1,2})?$", ""), _
                    Array("ВсегоСписано", True, "^\d+(\.\d{1,2})?$", ""), _
                    Array("КонечныйОстаток", True, "^\d+(\.\d{1,2})?$", "") _
                    )
                    
    Dim dfield As DataField
    For Each field In fields
    
        Set dfield = CreateDataField(field(0), field(1), field(2), field(3))
        dfield.section = "ПлатежныйДокумент"
        
        CreateAccountSection.Add field(0), dfield
        
    Next field
    
End Function
Private Function CreateDocSection() As Object

    Set CreateDocSection = CreateObject("Scripting.Dictionary")
    
    Dim fields As New Collection
    fields.Add Array("СекцияДокумент", True, "", "Признак начала секции, содержит вид документа")
    fields.Add Array("Номер", True, "", "Номер документа")
    fields.Add Array("Дата", True, "^\d{1,2}\.\d{1,2}\.\d{4}$", "Дата документа")
    fields.Add Array("Сумма", True, "^\d+(\.\d{1,2})?$", "Сумма платежа")
    fields.Add Array("КвитанцияДата", False, "^\d{1,2}\.\d{1,2}\.\d{4}$", "Дата формирования квитанции")
    fields.Add Array("КвитанцияВремя", False, "^\d{1,2}\:\d{1,2}\:\d{1,2}$", "Время формирования квитанции")
    fields.Add Array("КвитанцияСодержание", False, "", "Содержание квитанции")
    fields.Add Array("ПлательщикСчет", True, "^\d{20}$", "Расчетный счет плательщика")
    fields.Add Array("ДатаСписано", True, "^\d{1,2}\.\d{1,2}\.\d{4}$", "Дата списания средств с р/с")
    fields.Add Array("Плательщик", False, "", "ИНН и наименование плательщика")
    fields.Add Array("ПлательщикИНН", True, "^\d{10,12}$", "ИНН плательщика")
    fields.Add Array("Плательщик1", False, "", "Наименование плательщика")
    fields.Add Array("Плательщик2", False, "", "Расчетный счет плательщика")
    fields.Add Array("Плательщик3", False, "", "Банк плательщика")
    fields.Add Array("Плательщик4", False, "", "Город банка плательщика")
    fields.Add Array("ПлательщикРасчСчет", False, "^\d{20}$", "Корсчет банка плательщика")
    fields.Add Array("ПлательщикБанк1", False, "", "РЦ банка плательщика")
    fields.Add Array("ПлательщикБанк2", False, "", "Город РЦ банка плательщика")
    fields.Add Array("ПлательщикБИК", False, "^\d{9}$", "БИК РЦ банка плательщика")
    fields.Add Array("ПлательщикКорсчет", False, "^\d{20}$", "Корсчет РЦ банка плательщика")
    fields.Add Array("ПолучательСчет", True, "^\d{20}$", "Расчетный счет получателя")
    fields.Add Array("ДатаПоступило", True, "^\d{1,2}\.\d{1,2}\.\d{4}$", "дата поступления средств на р/с")
    fields.Add Array("Получатель", False, "", "Получатель одной строкой")
    fields.Add Array("ПолучательИНН", True, "^\d{10,12}$", "ИНН получателя")
    fields.Add Array("Получатель1", False, "", "Наименование получателя")
    fields.Add Array("Получатель2", False, "", "Расчетный счет получателя")
    fields.Add Array("Получатель3", False, "", "Банк получателя")
    fields.Add Array("Получатель4", False, "", "Город банка получателя")
    fields.Add Array("ПолучательРасчСчет", False, "^\d{20}$", "Корсчет банка получателя")
    fields.Add Array("ПолучательБанк1", False, "", "РЦ банка получателя")
    fields.Add Array("ПолучательБанк2", False, "", "Город РЦ банка получателя")
    fields.Add Array("ПолучательБИК", False, "^\d{9}$", "БИК РЦ банка получателя")
    fields.Add Array("ПолучательКорсчет", False, "^\d{20}$", "Корсчет РЦ банка получателя")
    fields.Add Array("ВидПлатежа", False, "", "Вид платежа")
    fields.Add Array("КодНазПлатежа", False, "", "Назначение платежа кодовое")
    fields.Add Array("ВидОплаты", False, "^\d{1,2}$", "Условное цифровое обозначение документа, согласно установленного ЦБР")
    fields.Add Array("Код", False, "^.{0,25}$", "Уникальный идентификатор платежа")
    fields.Add Array("НазначениеПлатежа", False, "", "Назначение платежа одной строкой")
    fields.Add Array("НазначениеПлатежа 1", False, "", "Назначение платежа разбитое на строки")
    fields.Add Array("НазначениеПлатежа 2", False, "", "Назначение платежа разбитое на строки")
    fields.Add Array("НазначениеПлатежа 3", False, "", "Назначение платежа разбитое на строки")
    fields.Add Array("НазначениеПлатежа 4", False, "", "Назначение платежа разбитое на строки")
    fields.Add Array("НазначениеПлатежа 5", False, "", "Назначение платежа разбитое на строки")
    fields.Add Array("НазначениеПлатежа 6", False, "", "Назначение платежа разбитое на строки")
    fields.Add Array("СтатусСоставителя", False, "^\d{2}$", "Статус составителя расчетного документа")
    fields.Add Array("ПлательщикКПП", False, "^(\d{9}|0)$", "КПП плательщика")
    fields.Add Array("ПолучательКПП", False, "^(\d{9}|0)$", "КПП получателя")
    fields.Add Array("ПоказательКБК", False, "^(\d{20}|0)$", "Показатель кода бюджетной классификации")
    fields.Add Array("ОКАТО", False, "^(\d{11}|0)$", "Код ОКТМО территории, на которой мобилизируются денежные средства от уплаты налога, сбора и иного платежа")
    fields.Add Array("ПоказательОснования", False, "", "^(\d{2}|0)$")
    fields.Add Array("ПоказательПериода", False, "^(\d{2}|0)$", "Показатель налогового периода / Код таможенного органа")
    fields.Add Array("ПоказательНомера", False, "", "Показатель номера документа")
    fields.Add Array("ПоказательДаты", False, "^\d{1,2}\.\d{1,2}\.\d{4}|0$", "Показатель даты документа")
    fields.Add Array("Очередность", False, "^\d{0,2}$", "Очередность платежа")
    fields.Add Array("СрокАкцепта", False, "^\d+$", "Срок акцепта, количество дней")
    fields.Add Array("ВидАккредитива", False, "", "Вид аккредитива")
    fields.Add Array("СрокПлатежа", False, "^\d{1,2}\.\d{1,2}\.\d{4}$", "Срок платежа (аккредитива)")
    fields.Add Array("УсловиОплаты1", False, "", "Условие оплаты, стр.1")
    fields.Add Array("УсловиОплаты2", False, "", "Условие оплаты, стр.2")
    fields.Add Array("УсловиОплаты3", False, "", "Условие оплаты, стр.3")
    fields.Add Array("ПлатежПоПредст", False, "", "Платеж по предоставлению")
    fields.Add Array("ДополнУсловия", False, "", "Дополнительные условия")
    fields.Add Array("НомерСчетПоставщика", False, "", "Номер счета поставщика")
    fields.Add Array("ДатаОтсылкиДок", False, "^\d{1,2}\.\d{1,2}\.\d{4}$", "Дата отсылки документов")

    Dim dfield As DataField
    For Each field In fields
    
        Set dfield = CreateDataField(field(0), field(1), field(2), field(3))
        dfield.section = "ПлатежныйДокумент"
        
        CreateDocSection.Add field(0), dfield
        
    Next field
    
End Function
Private Sub ReadAccountSection()

    If key = accTrigger_begin Then
    
        Call NextLine
        If key = FileEOF Then
            log "Неожиданный конец файла СекцияРасчСчет"
            Exit Sub
        End If
        
        Dim accSection As Object
        Set accSection = CreateAccountSection()
        If accSection.exists(getkey()) Then

            Do While Not textStream.AtEndOfStream And key <> accTrigger_finish And key <> docTrigger_begin And key <> "КонецФайла"
                If accSection.exists(key) Then
                    Dim accField As DataField
                    Set accField = accSection.Item(key)
                    accField.value = value
                    accField.row_index = row_index
                    
                    If (Not accField.testValue) Then
                        log " Идентификатор: " & key & " имеет неверное значение", accField
                    End If
                    Set accField = Nothing
                End If
                Call NextLine
                
            Loop
            
            accountSections.Add accSection
            
            If key <> accTrigger_finish Then
                log "Не обнаружен признак конца секции СекцияРасчСчет"
            ElseIf key = accTrigger_finish Then
                NextLine
            End If
            
        Else
        
            log "Не обнаружены данные секции СекцияРасчСчет"
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
        
        Do While Not textStream.AtEndOfStream And key <> docTrigger_finish And key <> "КонецФайла"
           If docSection.exists(key) Then
                Dim docField As DataField
                Set docField = docSection.Item(key)
                docField.value = value
                docField.row_index = row_index
                If (Not docField.testValue) Then
                    log " Идентификатор: " & key & " имеет неверное значение", docField
                End If
                Set docField = Nothing
            End If
            Call NextLine
            
        Loop
        
        docSections.Add docSection
            
        If key <> docTrigger_finish Then
            log "Не обнаружен признак конца секции СекцияДокумент"
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
        Error_string = file_name & " строка: " & row_index & " [Сообщение] " & message
    Else
        Error_string = file_name & " строка: " & field.row_index & " [Сообщение] " & field.section & " " & message
    End If
    
    errors.Add Error_string

    
End Sub
Private Sub validatelog(ByVal message, Optional ByRef field As DataField)
    Dim Error_string As String
    If field Is Nothing Then
        Error_string = file_name & " [Сообщение] " & message
    Else
        Error_string = file_name & " " & field.row_index & " [Сообщение] " & field.section & " " & message
    End If
    errors.Add Error_string

    
End Sub
Private Sub Normalize()
    
    Dim doc As Object
    validatelog "Нормализация данных"
    For Each doc In docSections
    
        Call NormalizePaymentDetails(doc)
        Call NormalizeBankAccounts(doc)
        Call NormalizeCompanyName(doc)
        
    Next doc
    validatelog "Конец нормализации данных"
End Sub
Private Sub NormalizePaymentDetails(ByRef doc As Object)

    Dim paymentDetails As DataField
    Set paymentDetails = doc.Item("НазначениеПлатежа")
    
    If Not paymentDetails.filled Or paymentDetails.value = "" Then
        log "НазначениеПлатежа пустое. Попытка объединить многострочное поле.", paymentDetails
        paymentDetails.value = paymentDetail.value & doc.Item("НазначениеПлатежа1").value
        paymentDetails.value = paymentDetail.value & doc.Item("НазначениеПлатежа2").value
        paymentDetails.value = paymentDetail.value & doc.Item("НазначениеПлатежа3").value
        paymentDetails.value = paymentDetail.value & doc.Item("НазначениеПлатежа4").value
        paymentDetails.value = paymentDetail.value & doc.Item("НазначениеПлатежа5").value
        paymentDetails.value = paymentDetail.value & doc.Item("НазначениеПлатежа6").value
        If paymentDetail.value <> "" Then
            log "НазначениеПлатежа успешно собрано из нескольких нумерованых полей", paymentDetails
        Else
            log "НазначениеПлатежа не удалось обнаружить", paymentDetails
        End If
    End If

End Sub

Private Sub NormalizeBankAccounts(ByRef doc As Object)

    If Not doc.Item("ПлательщикСчет").filled Then
        If Not doc.Item("Плательщик2").filled Then
            doc.Item("ПлательщикСчет").value = doc.Item("ПлательщикРасчСчет").value
            log "Заполнение поля ПлательщикСчет на основани значенияв поле ПлательщикРасчСчет", doc.Item("ПлательщикСчет")
        Else
            doc.Item("ПлательщикСчет").value = doc.Item("Плательщик2").value
            log "Заполнение поля ПлательщикСчет на основани значенияв поле Плательщик2", doc.Item("ПлательщикСчет")
        End If
    End If
    If Not doc.Item("ПолучательСчет").filled Then
        If Not doc.Item("Получатель2").filled Then
            doc.Item("ПолучательСчет").value = doc.Item("ПолучательРасчСчет").value
            log "Заполнение поля ПолучательСчет на основани значенияв поле ПолучательРасчСчет", doc.Item("ПолучательСчет")
        Else
            doc.Item("ПолучательСчет").value = doc.Item("Получатель2").value
             log "Заполнение поля ПолучательСчет на основани значенияв поле Получатель2", doc.Item("ПолучательСчет")
        End If
    End If
    
End Sub
Private Sub NormalizeCompanyName(ByRef doc As Object)

    Dim matches As Object
    Dim regExp As Object
    Set regExp = CreateObject("VBScript.RegExp")
    
    If Not doc.Item("Плательщик1").filled Then
       
        With regExp
            .Global = True
            .IgnoreCase = True
            .MultiLine = False
            .pattern = "^(\d{10,12})( )(.+)$"
        End With
        
        If regExp.test(doc.Item("Плательщик").value) Then
            
            Set matches = regExp.Execute(doc.Item("Плательщик").value)
            doc.Item("Плательщик1").value = Trim(matches(0).submatches(2))
                log "Извлечение поля Плательщик1 из поля Плательщик", doc.Item("Плательщик")
            If Not doc.Item("ПлательщикИНН").filled Then
                doc.Item("ПлательщикИНН").value = Trim(matches(0).submatches(0))
                log "Извлечение поля ПлательщикИНН из поля Плательщик", doc.Item("Плательщик")
            End If
            
        End If
        
    End If
    
    If Not doc.Item("Получатель1").filled Then

        With regExp
            .Global = True
            .IgnoreCase = True
            .MultiLine = False
            .pattern = "^(\d{10,12})( )(.+)$"
        End With
        
        If regExp.test(doc.Item("Получатель").value) Then
            Set matches = regExp.Execute(doc.Item("Получатель").value)
            doc.Item("Получатель1").value = Trim(matches(0).submatches(2))
            log "Извлечение поля Получатель1 из поля Получатель", doc.Item("Получатель")
            If Not doc.Item("ПолучательИНН").filled Then
                doc.Item("ПолучательИНН").value = Trim(matches(0).submatches(0))
                log "Извлечение поля ПолучательИНН из поля Плательщик", doc.Item("Получатель")
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
        .Filters.Add "Банковские выписки", "*.1c,*.txt", 1
       
    End With
    file_dialog.Show
    Set Get_file_name = file_dialog.SelectedItems

End Function
