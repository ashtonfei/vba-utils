VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim utils As New clsUtils

Public data As New Collection
Public hasError As Boolean
Public errorMessage As String

Private Function readRowData(ByVal key As String, ByRef values As Variant, ByVal rowIndex As Long, ByVal headerRowIndex As Long)
    If utils.collectionHasKey(data, key) Then Exit Function
    Dim item As New Collection
    Dim i As Long
    For i = 1 To UBound(values, 2)
        Dim itemKey As String
        itemKey = Trim(values(headerRowIndex, i))
        If IsEmpty(itemKey) = False Then
            If utils.collectionHasKey(item, itemKey) = False Then
                item.Add values(rowIndex, i), itemKey
            End If
        End If
    Next i
    data.Add item, key
End Function
Private Function readData(ByRef sheet As clsSheetValiator, ByVal keyColumnName As String)
    Dim values As Variant
    values = sheet.sheet.Range("A" & sheet.headerRowIndex).CurrentRegion.value
    
    Dim headerRowIndex As Long
    headerRowIndex = sheet.headerRowIndex
    
    Dim keyIndex As Long
    keyIndex = sheet.headerIndexes(keyColumnName)
    
    Dim i As Long
    For i = headerRowIndex + 1 To UBound(values, 1)
        Dim key As String
        key = Trim(values(i, keyIndex))
        If IsEmpty(key) = False Then
            Call readRowData(key, values, i, headerRowIndex)
        End If
    Next i
End Function
Public Function init(ByVal sheetName As String, Optional ByVal keyColumnName As String = "Material", Optional ByVal headerRowIndex As Long = 1)
    Dim sheet As New clsSheetValiator
    Call sheet.init(ActiveWorkbook, sheetName, Array(keyColumnName), headerRowIndex)
    If sheet.hasError Then
        hasError = True
        errorMessage = sheet.errorMessage
        Exit Function
    End If
    Call readData(sheet, keyColumnName)
End Function

Function getValue(ByVal key As String, ByVal columnName As String) As Variant
    If utils.collectionHasKey(data, key) = False Then
        getValue = Null
        Exit Function
    End If
    If utils.collectionHasKey(data(key), columnName) = False Then
        getValue = Null
        Exit Function
    End If
    getValue = data(key)(columnName)
End Function

