VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Автор: Шибанов Ф.В.
' mail: shibanovfv@yandex.ru

Public key As String
Private p_value As String
Public valuePattern As String
Public required As Boolean
Public description As String
Public section As String
Public row_index As Integer

Private p_filled As Boolean
Public Property Get filled()
    filled = p_filled
End Property

Public Property Let value(val As String)
    p_value = val
    p_filled = True
End Property
Public Property Get value() As String
    value = p_value
End Property


Public Static Function testValue()
    Dim regExp As Object
    Set regExp = CreateObject("VBScript.RegExp")
    With regExp
        .Global = True
        .IgnoreCase = True
        .MultiLine = False
        .pattern = valuePattern
    End With
    testValue = regExp.test(value)
End Function

Private Sub Class_Initialize()
    p_filled = False
End Sub
