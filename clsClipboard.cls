VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function read_clipboard() As String
    Dim clipboard As New MSForms.DataObject
    clipboard.GetFromClipboard
    read_clipboard = Trim(clipboard.GetText)
End Function

Function read_clipboard_to_array(Optional delimiter As String = vbTab) As Variant
    Dim text As String
    text = read_clipboard
    Dim rows As Variant
    rows = Split(text, vbNewLine)
    
    Dim cols As Variant
    cols = Split(rows(0), delimiter)
    
    If UBound(cols) = -1 Then Exit Function
    
    Dim values As Variant
    ReDim values(0 To UBound(rows), 0 To UBound(cols))
    
    Dim r As Long
    For r = LBound(rows) To UBound(rows)
        Dim c As Long
        cols = Split(rows(r), delimiter)
        For c = LBound(cols) To UBound(cols)
            values(r, c) = cols(c)
        Next c
    Next r
    read_clipboard_to_array = values
End Function
