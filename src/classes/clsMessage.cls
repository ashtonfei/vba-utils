VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public success As String
Public error As String
Public confirm As String
Public info As String
Public version As String

Public Sub init(ByVal appName As String)
    success = appName & " [SUCCESS]"
    error = appName & " [ERROR]"
    confirm = appName & " [CONFIRM]"
    info = appName & " [INFO]"
    version = appName & " [VERSION]"
End Sub

