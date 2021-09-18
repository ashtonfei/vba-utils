VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function disableExcelDefaultSettings()
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With
End Function

Function enableExcelDefaultSettings()
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
    End With
End Function

Function alert(ByVal msg As String, Optional ByVal msg_box_style As VbMsgBoxStyle = vbInformation, Optional ByVal title As String = "Message")
    MsgBox msg, msg_box_style, title
End Function

Function confirm(ByVal msg As String, Optional ByVal title As String = "Confirmation") As VbMsgBoxResult
    confirm = MsgBox(msg, vbYesNo, title)
End Function

Function workbookHasSheet(ByRef Workbook As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Workbook.Sheets(sheetName)
    workbookHasSheet = (Err.number = 0)
    Err.Clear
End Function

Function activeWorkbookHasSheets(ByRef arrayRequiredSheetNames As Variant) As Boolean
    ' read sheet names into a sheet name collection
    Dim sheetnames As New Collection
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        sheetnames.Add ws.name, ws.name
    Next ws

    ' check each name in array required sheet names
    Dim name As Variant
    For Each name In arrayRequiredSheetNames
        If Not collectionHasKey(sheetnames, name) Then
            activeWorkbookHasSheets = False
            Set sheetnames = Nothing
            Exit Function
        End If
    Next name
    Set sheetnames = Nothing
    activeWorkbookHasSheets = True
End Function

Function arrayHasValue(ByVal arrayItems As Variant, ByVal value As Variant) As Boolean
    Dim item As Variant
    For Each item In arrayItems
        If item = value Then
            arrayHasValue = True
            Exit Function
        End If
    Next item
    arrayHasValue = False
End Function

Function arraySort(ByRef arrayItems As Variant, Optional ByVal desc As Boolean = False)
    Dim n As Long
    n = UBound(arrayItems)
    If n < 1 Then Exit Function
    Dim i As Long
    Do
        Dim swapped As Boolean
        swapped = False
        For i = 1 To n
            Dim previousItem, thisItem As Variant
            previousItem = arrayItems(i - 1)
            thisItem = arrayItems(i)
            If (previousItem > thisItem And desc = False) _
                Or (previousItem < thisItem And desc = True) Then
                arrayItems(i) = previousItem
                arrayItems(i - 1) = thisItem
                swapped = True
            End If
        Next i
    Loop Until Not swapped
End Function

Function collectionHasValue(ByRef coll As Collection, ByVal value As Variant) As Boolean
    Dim item As Variant
    For Each item In coll
        If item = value Then
            collectionHasValue = True
            Exit Function
        End If
    Next item
    collectionHasValue = False
End Function

Function collectionHasKey(ByRef coll As Collection, ByVal key As String) As Boolean
    Dim item As Variant
    On Error Resume Next
    item = IsObject(coll(key))
    collectionHasKey = (Err.number = 0)
    Err.Clear
End Function

Function collectionSwap(ByRef coll As Collection, ByVal position As Long)
    Dim tempValue As Variant
    tempValue = coll(position)
    Call coll.Remove(position)
    Call coll.Add(tempValue, , position - 1)
End Function

Function collectionSort(ByRef coll As Collection, Optional ByRef desc As Boolean = False)
    Dim n As Long
    n = coll.Count
    If n < 2 Then Exit Function
    Dim i As Long
    Do
        Dim swapped As Boolean
        swapped = False
        For i = 2 To n
            Dim previousItem, thisItem As Variant
            previousItem = coll(i - 1)
            thisItem = coll(i)
            If (previousItem > thisItem And desc = False) _
                Or (previousItem < thisItem And desc = True) Then
                Call collectionSwap(coll, i)
                swapped = True
            End If
        Next i
    Loop Until Not swapped
End Function


Function getSheetByName(ByRef Workbook As Workbook, ByVal sheetName As String) As Worksheet
    If workbookHasSheet(Workbook, sheetName) Then
        Set getSheetByName = Workbook.Sheets(sheetName)
    Else
        Set getSheetByName = Workbook.Sheets.Add
        getSheetByName.name = sheetName
    End If
End Function

Function valuesToWorksheet(ByRef wb As Workbook, ByVal sheetName As String, ByRef values As Variant, Optional ByVal clearFormat As Boolean = False) As Worksheet
    Dim ws As Worksheet
    Set ws = getSheetByName(wb, sheetName)
    With ws
        If clearFormat Then
            .Cells.Clear
        Else
            .Cells.ClearContents
        End If
        .Range(ws.Cells(1, 1), ws.Cells(UBound(values, 1), UBound(values, 2))).value = values
        .Activate
        .rows.AutoFit
        .columns.AutoFit
    End With
    Set valuesToWorksheet = ws
End Function

Function advacedFilter( _
    ByRef sourceRange As Range, _
    ByRef criterias As Variant, _
    ByRef copyToSheet As Worksheet, _
    Optional ByRef headers As Variant, _
    Optional ByVal removeFormat As Boolean = True, _
    Optional ByVal unique As Boolean = False) As Variant

    Dim criteriaRange As Range
    Dim copyToRange As Range

    Dim targetRow As Long
    targetRow = UBound(criterias, 1) + 2

    With copyToSheet
        Call removeFilters(copyToSheet)
        If removeFormat Then
            .Cells.Clear
        Else
            .Cells.ClearContents
        End If
        .Range(.Cells(1, 1), .Cells(UBound(criterias, 1), UBound(criterias, 2))).value = criterias
        Set criteriaRange = .Range("A1").CurrentRegion

        If Not IsMissing(headers) Then
            .Range(.Cells(targetRow, 1), .Cells(targetRow, UBound(headers) + 1)).value = createHeaderArray(headers)
            Set copyToRange = .Range("A" & targetRow).CurrentRegion
        Else
            Set copyToRange = .Range("A" & targetRow)
        End If
        sourceRange.AdvancedFilter xlFilterCopy, criteriaRange, copyToRange, unique
        .Range("1:" & (targetRow - 1)).Delete
        .rows.AutoFit
        .columns.AutoFit
        If removeFormat Then
            .Cells.ClearFormats
        End If
        advacedFilter = .Range("A1").CurrentRegion.value
    End With
End Function
Function removeFilters(ByRef ws As Worksheet)
    If ws.AutoFilterMode Then ws.AutoFilter.ShowAllData
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        Debug.Print lo.ShowAutoFilter
        If lo.ShowAutoFilter Then
            lo.AutoFilter.ShowAllData
            lo.ShowAutoFilter = False
        End If
    Next lo
    ws.rows.Hidden = False
    ws.columns.Hidden = False
End Function

Function createJSON(ByRef keys As Variant, Optional ByRef values As Variant) As Collection
    If IsMissing(values) Then values = keys
    Dim coll As New Collection
    Dim min As Long
    min = UBound(keys)
    If UBound(values) < min Then min = UBound(values)
    Dim i As Long
    For i = 0 To min
        coll.Add values(i), keys(i)
    Next i
    Set createJSON = coll
End Function

Function getUniqueValues(ByRef values As Variant) As Collection
    Dim coll As New Collection
    Dim value As Variant
    For Each value In values
        Dim key As String
        key = CStr(value)
        If Not collectionHasKey(coll, key) And key <> "" Then
            coll.Add value, key
        End If
    Next value
    Set getUniqueValues = coll
End Function

Function getUniqueValuesByColumn(ByRef values As Variant, ByVal columnIndex As Long) As Collection
    Dim coll As New Collection
    Dim i As Long
    For i = LBound(values, 1) To UBound(values, 1)
        Dim key As String
        key = CStr(values(i, columnIndex))
        If Not collectionHasKey(coll, key) And key <> "" Then
            coll.Add values(i, columnIndex), key
        End If
    Next i
    Set getUniqueValuesByColumn = coll
End Function

Function setCollectionItem(ByRef coll As Collection, ByVal key As String, ByVal value As Variant)
    If collectionHasKey(coll, key) Then
        Call coll.Remove(key)
    End If
    coll.Add value, key
End Function

Function applyDuplicateFormat(ByRef rng As Range, Optional ByVal fontColor As Long = vbWhite, Optional ByVal bgColor As Long = vbRed)
    With rng.FormatConditions
        .Delete
        .AddUniqueValues
        .item(1).DupeUnique = xlDuplicate
        .item(1).StopIfTrue = False
        .item(1).Font.Color = vbWhite
        .item(1).Interior.Color = bgColor
    End With
End Function

Function capitalize(ByVal text As String) As String
    text = Trim(text)
    capitalize = UCase(Left(text, 1)) & LCase(Mid(text, 2))
End Function

Function capitalizeHeader(ByVal header As String, Optional sep As String = " ") As String
    header = Trim(header)
    Dim words As Variant
    words = Split(header, sep)
    Dim i As Integer
    For i = 0 To UBound(words)
        words(i) = capitalize(words(i))
    Next i
    capitalizeHeader = Join(words, sep)
End Function

Function getLastRow(ByRef ws As Worksheet, ByVal columnIndex As Long) As Long
    getLastRow = ws.Cells(ws.rows.Count, columnIndex).End(xlUp).Row
End Function

Function getLastColumn(ByRef ws As Worksheet, ByVal rowIndex As Long) As Long
    getLastColumn = ws.Cells(rowIndex, ws.columns.Count).End(xlToLeft).Column
End Function

Function createHeaderArray(ByRef headers) As Variant
    Dim values As Variant
    ReDim values(1 To 1, 1 To UBound(headers) + 1)
    Dim i As Long
    For i = 0 To UBound(headers)
        values(1, i + 1) = headers(i)
    Next i
    createHeaderArray = values
End Function

Function setHeaders(ByRef ws As Worksheet, ByRef headers As Variant, Optional ByVal headerRowIndex As Long = 1, Optional ByVal headerStartColumnIndex As Long = 1)
    Dim values As Variant
    values = createHeaderArray(headers)
    With ws
        .Range(.Cells(headerRowIndex, headerStartColumnIndex), .Cells(headerRowIndex, UBound(values, 2) + headerStartColumnIndex - 1)).value = values
    End With
End Function

Function getMonday(ByVal dateValue As Variant) As Date
    getMonday = DateAdd("d", 1 - Weekday(dateValue, vbMonday), dateValue)
End Function
