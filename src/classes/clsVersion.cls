VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim utils As New clsUtils
Public author As String
Private changes As New Collection

Property Get version() As String
    version = changes(changes.Count)(0)
End Property

Private Sub Class_Initialize()
    author = "Ashton Fei"
    changes.Add Array("1.0.0", "Initial release for 'SMC Merge Application'", "9/14/2021")
End Sub

Function show(ByVal title As String)
    Dim msg As String
    msg = "Changes:" & vbNewLine
    Dim change As Variant
    For Each change In changes
        msg = msg & vbNewLine & change(0) & ": " & change(1) & " [" & change(2) & "]"
    Next change
    
    msg = msg & vbNewLine & vbNewLine & "Created by - " & author
    Call utils.alert(msg, vbInformation, title)
End Function
