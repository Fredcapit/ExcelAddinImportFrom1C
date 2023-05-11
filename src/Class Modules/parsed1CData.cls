VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "parsed1CData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Автор: Шибанов Ф.В.
' mail: shibanovfv@yandex.ru

    Public firstRow As String
    Public formatVersion As String
    Public generalInfo As Object
    Public selectionInfo As Object
    Public accountSections As Collection
    Public docSections As Collection
    Public cancelled As Boolean
    Public errors As Collection

