VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSheetValiator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public sheet As Worksheet
Public hasError As Boolean
Public errorMessage As String
Public headerIndexes As New Collection

Private M_WB As Workbook
Private M_NAME As String
Private M_HEADER_ROW_INDEX As Long
Private M_REQUIRED_HEADERS As Variant

Private utils As New clsUtils
Property Get wb() As Workbook
    Set wb = M_WB
End Property
Property Set wb(ByRef value As Workbook)
    Set M_WB = value
End Property

Property Get name() As String
    name = M_NAME
End Property
Property Let name(ByVal value As String)
    M_NAME = value
End Property

Property Get headerRowIndex() As Long
    headerRowIndex = M_HEADER_ROW_INDEX
End Property
Property Let headerRowIndex(ByVal value As Long)
    M_HEADER_ROW_INDEX = value
End Property

Property Get requiredHeaders() As Variant
    requiredHeaders = M_REQUIRED_HEADERS
End Property
Property Let requiredHeaders(ByVal value As Variant)
    M_REQUIRED_HEADERS = value
End Property

Function init(ByRef wb As Workbook, ByVal name As String, Optional ByRef requiredHeaders As Variant, Optional ByVal headerRowIndex As Long = 1)
    Set M_WB = wb
    M_NAME = name
    M_HEADER_ROW_INDEX = headerRowIndex
    M_REQUIRED_HEADERS = requiredHeaders
    Call validate
End Function

Function getHeaders()
    If hasError Then Exit Function
    Call updateHeaders
    Dim values As Variant
    values = wb.Sheets(name).usedRange.value
    
    If IsEmpty(values) Then
        hasError = True
        errorMessage = Join(Array("No headers found at row ", headerRowIndex, " in sheet '", name, "'."), "")
        Exit Function
    End If
    If headerRowIndex > UBound(values, 1) Then
        hasError = True
        errorMessage = Join(Array("No headers found at row ", headerRowIndex, " in sheet '", name, "'."), "")
        Exit Function
    End If
    Dim i As Long
    For i = 1 To UBound(values, 2)
        Dim header As String
        header = Trim(values(headerRowIndex, i))
        If Not utils.collectionHasKey(headerIndexes, header) Then
            headerIndexes.Add i, header
        End If
    Next i
End Function
Private Function updateHeaders()
    If hasError Then Exit Function
    Dim headerRange As Range
    Dim lastColumn As Long
    lastColumn = utils.getLastColumn(sheet, M_HEADER_ROW_INDEX)
    Set headerRange = sheet.Range(sheet.Cells(M_HEADER_ROW_INDEX, 1), sheet.Cells(M_HEADER_ROW_INDEX, lastColumn))
    Dim headers As Variant
    headers = headerRange.value
    Dim i As Long
    For i = 1 To lastColumn
        headers(1, i) = utils.capitalizeHeader(headers(1, i))
    Next i
    headerRange.value = headers
End Function

Private Function checkRequiredHeaders()
        If hasError Then Exit Function
        Call getHeaders
        Dim header As Variant
        For Each header In requiredHeaders
            If header <> "" Then
                If Not utils.collectionHasKey(headerIndexes, header) Then
                    hasError = True
                    errorMessage = Join(Array("Header '", header, "' was not found in sheet '", name, "',"), "")
                    Exit Function
                End If
            End If
        Next header
End Function

Private Function validate()
    hasError = Not utils.workbookHasSheet(wb, name)
    If hasError Then
        errorMessage = Join(Array("Sheet '", name, " ' was not found in workbook ' ", wb.name, "'."), "")
        Exit Function
    End If
    Set sheet = wb.Sheets(name)
    Call checkRequiredHeaders
End Function
