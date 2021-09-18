VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileSystemObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public fso As FileSystemObject

Public Function join_(ByVal path As String, ByVal items As Variant) As String
    If VarType(items) = vbString Then
        join_ = fso.BuildPath(path, items)
        Exit Function
    End If
    Dim item As Variant
    For Each item In items
        join_ = join_ & fso.BuildPath(path, item)
    Next item
End Function

Public Function getFiles_(ByVal folderPath As String, Optional ByVal fileExtensions As Variant = "*") As Collection
    If fso.FolderExists(folderPath) = False Then Exit Function
    Dim files_ As New Collection
    Dim folder_ As folder
    Set folder_ = fso.GetFolder(folderPath)
    Dim file_ As file
    For Each file_ In folder_.files
        If file_.name <> "" Then files_.Add file_
    Next file_
    Set getFiles_ = files_
    Set files_ = Nothing
End Function

Private Sub Class_Initialize()
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub
