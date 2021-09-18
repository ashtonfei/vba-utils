VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const FOLDER_NAME_MODULES = "modules"
Const FOLDER_NAME_CLASS_MODUELS = "classes"
Const FOLDER_NAME_FORMS = "forms"

Private wb As Workbook
Private fso As FileSystemObject

Public Function import_(Optional ByVal folderName As String = "src")
    Dim path As String
    path = fso.BuildPath(path, folderName)
    Dim files As files
    Dim folders As folders
End Function

Public Function export_(Optional ByVal folderName As String = "src")
    
End Function

Private Sub Class_Initialize()
    Set wb = ThisWorkbook
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub
