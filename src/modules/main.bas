Attribute VB_Name = "main"
Option Explicit

Public Sub DEMO_JSON()
    ' Let's say we have a sheet named "Users" and have headers "ID", "Name", "Age", "Email" at row 1
    Dim users As New clsJson
    Call users.init(sheetName:="Users", keyColumnName:="ID", headerRowIndex:=1)
    ' get user data
    Debug.Print "Name:" & Space(4) & users.getValue(key:="1", columnName:="Name")
    Debug.Print "Email:" & Space(4) & users.getValue("1", "Email")
    Debug.Print "Age:" & Space(4) & users.data("1")("Age")
End Sub


