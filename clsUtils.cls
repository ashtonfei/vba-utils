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
'utilities functions
Function message(ByVal msg As String, ByVal msg_box_style As VbMsgBoxStyle, Optional ByVal title As String = "Message")
    MsgBox msg, msg_box_style, title
End Function

Function confirm(ByVal msg As String, Optional ByVal title As String = "Confirmation") As VbMsgBoxResult
    confirm = MsgBox(msg, vbYesNo, title)
End Function

Function get_last_row(ByRef ws As Worksheet, ByVal col_index As Long) As Long
    get_last_row = ws.Cells(ws.rows.count, col_index).End(xlUp).Row
End Function

Function get_last_col(ByRef ws As Worksheet, ByVal row_index As Long) As Long
    get_last_col = ws.Cells(row_index, ws.Columns.count).End(xlToLeft).Column
End Function

Function get_sheet_by_name(ByVal name As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        If UCase(Trim(ws.name)) = UCase(Trim(name)) Then
            Set get_sheet_by_name = ws
            Exit Function
        End If
    Next ws
    Set get_sheet_by_name = ActiveWorkbook.Sheets.Add
    get_sheet_by_name.name = name
End Function

Function sum_array(ByRef data As Variant, Optional ByVal count As Long = 0) As Long
    Dim i As Variant
    Dim c As Long
    For Each i In data
        If c < count Or count = 0 Then
            If IsNumeric(i) Then sum_array = sum_array + i
            c = c + 1
        Else
            Exit For
        End If
    Next i
End Function

Function copy_to(ByRef from_ws As Worksheet, ByRef to_ws As Worksheet)
    Dim values As Variant
    values = from_ws.UsedRange.value
    With to_ws
        .Cells.Clear
        .Range(.Cells(1, 1), .Cells(UBound(values, 1), UBound(values, 2))).value = values
        .Activate
    End With
End Function

Function includes(ByVal values As String, ByVal value As String) As Boolean
    includes = False
    Dim item As Variant
    For Each item In Split(values, ",")
        If UCase(Trim(item)) = UCase(Trim(value)) Then
            includes = True
            Exit Function
        End If
    Next item
End Function

Function contains(ByVal text As String, ByVal to_be_found_text As String) As Boolean
    contains = InStr(1, text, to_be_found_text, vbTextCompare) > 0
End Function

