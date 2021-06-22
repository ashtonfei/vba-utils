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
Function message(ByVal msg As String, Optional ByVal msg_box_style As VbMsgBoxStyle = vbInformation, Optional ByVal title As String = "Message")
    MsgBox msg, msg_box_style, title
End Function

Function confirm(ByVal msg As String, Optional ByVal title As String = "Confirmation") As VbMsgBoxResult
    confirm = MsgBox(msg, vbYesNo, title)
End Function

Function update_statusbar(ByVal message As String)
    Application.StatusBar = message
End Function

Function get_last_row(ByRef ws As Worksheet, ByVal col_index As Long) As Long
    get_last_row = ws.Cells(ws.rows.count, col_index).End(xlUp).Row
End Function

Function get_last_column(ByRef ws As Worksheet, ByVal row_index As Long) As Long
    get_last_column = ws.Cells(row_index, ws.Columns.count).End(xlToLeft).Column
End Function

Function get_used_range(ByRef ws As Worksheet, Optional ByVal row_index As Long = 1, Optional ByVal column_index As Long = 1) As Range
    Dim last_row As Long
    last_row = get_last_row(ws, column_index)
    Dim last_column As Long
    last_column = get_last_column(ws, row_index)
    Set get_used_range = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_column))
End Function

Function get_sheet_by_name(ByVal name As String, Optional ByVal create_new As Boolean = True, Optional ByVal case_sensitive As Boolean = False) As Worksheet
    name = Trim(name)
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        Dim ws_name As String
        ws_name = ws.name
        If Not case_sensitive Then
           name = UCase(name)
           ws_name = UCase(ws_name)
        End If
        If ws_name = name Then
            Set get_sheet_by_name = ws
            Exit Function
        End If
    Next ws
    If create_new Then
        Set get_sheet_by_name = ActiveWorkbook.Sheets.Add
        get_sheet_by_name.name = name
    End If
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

Function includes(ByVal values As Variant, ByVal value As String, Optional ByVal case_sensitive As Boolean = False) As Boolean
    includes = False
    If VarType(values) = vbString Then values = Split(values, ",")
    value = Trim(value)
    Dim item As Variant
    For Each item In values
        item = Trim(item)
        If case_sensitive Then
            item = UCase(Trim(item))
            value = UCase(value)
        End If
        If item = value Then
            includes = True
            Exit Function
        End If
    Next item
End Function

Function contains(ByVal text As String, ByVal to_be_found_text As String) As Boolean
    contains = InStr(1, text, to_be_found_text, vbTextCompare) > 0
End Function

Sub disable_excel_default_settings()
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With
End Sub

Sub enable_excel_default_settings()
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Function has_key(ByRef coll As Collection, ByVal key As String) As Boolean
    Dim item As Variant
    On Error Resume Next
    item = IsObject(coll(key))
    has_key = (Err.Number = 0)
    Err.Clear
End Function

Function output_to_sheet(ByRef ws As Worksheet, ByVal values As Variant)
    With ws
        .Cells.ClearContents
        .Range(.Cells(1, 1), .Cells(UBound(values, 1) - LBound(values, 1) + 1, UBound(values, 2) - LBound(values, 2) + 1)).value = values
        .Activate
    End With
End Function

Function get_header_columns(ByRef values As Variant, Optional ByVal HEADER_ROW_INDEX As Long = 1, Optional ByVal upper_case As Boolean = True) As Collection
    Dim coll As New Collection
    Dim i As Long
    Dim header As String
    For i = LBound(values, 2) To UBound(values, 2)
        header = Trim(values(HEADER_ROW_INDEX, i))
        If upper_case Then header = UCase(header)
        If Not has_key(coll, header) Then coll.Add i, header
    Next i
    Set get_header_columns = coll
End Function

Function send_email(ByVal subject As String, _
    ByVal mail_to As String, _
    ByVal html_body As String, _
    Optional ByVal cc As String, _
    Optional ByVal bcc As String, _
    Optional ByVal attachments As String)
    
    Dim app As Object
    Dim mail As Object

    Set app = CreateObject("Outlook.Application")
    Set mail = app.CreateItem(0)
    
    With mail
        .To = mail_to
        .cc = cc
        .bcc = bcc
        .subject = subject
        .HTMLBody = html_body
        
        Dim file As Variant
        If attachments <> "" Then
            For Each file In Split(attachments, ",")
                .attachments.Add file
            Next file
        End If
        
    End With
    mail.send
End Function

Function remove_all_filters(ByRef ws As Worksheet)
    If ws.AutoFilterMode Then ws.AutoFilter.ShowAllData
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        Debug.Print lo.ShowAutoFilter
        If lo.ShowAutoFilter Then
            lo.AutoFilter.ShowAllData
            lo.ShowAutoFilter = False
        End If
    Next lo
End Function

Function create_week_key(ByVal d_date As Variant, Optional ByVal suffix As String = "") As String
    If VarType(d_date) = vbString Then d_date = DateValue(d_date)
    
    Dim i_week_number As Integer: i_week_number = CInt(Format(d_date, "WW", vbSunday, vbFirstFullWeek))
    Dim i_year As Integer: i_year = Year(d_date)
    If Month(d_date) = 1 And i_week_number > 10 Then
        i_year = i_year - 1
    End If
    create_week_key = "W" + CStr(i_week_number) + "Y" + Right(CStr(i_year), 2) + suffix
End Function

Function create_month_key(ByVal d_date As Variant, Optional ByVal suffix As String = "") As String
    If VarType(d_date) = vbString Then d_date = DateValue(d_date)
    Dim s_month As String: s_month = Format(d_date, "MM")
    Dim i_year As Integer: i_year = Year(d_date)
    create_month_key = "M" + s_month + "Y" + Right(CStr(i_year), 2) + suffix
End Function

Function join_collection(ByRef coll As Collection, Optional ByVal sep As String = ",") As String
    If coll.count = 0 Then Exit Function
    Dim arr As Variant
    ReDim arr(1 To coll.count)
    Dim item As Variant
    Dim i As Long: i = 1
    For Each item In coll
        arr(i) = item
        i = i + 1
    Next item
    join_collection = Join(arr, sep)
End Function

Function is_number(ByVal value As Variant) As Boolean
    is_number = False
    If VarType(value) = vbDouble Then
        is_number = True
    ElseIf VarType(value) = vbInteger Then
        is_number = True
    ElseIf VarType(value) = vbDecimal Then
        is_number = True
    ElseIf VarType(value) = vbLong Then
        is_number = True
    ElseIf VarType(value) = vbCurrency Then
        is_number = True
    End If
End Function


