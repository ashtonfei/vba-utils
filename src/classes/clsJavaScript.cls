VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJavaScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function getArrayLength(ByVal items As Variant) As Long
    Const func = "function func(items){return items.length;}"
    Dim script As New ScriptControl
    With script
        .Language = "JScript"
        .AddCode func
        getArrayLength = .run("func", items)
    End With
End Function


Public Function arrayToCollection(ByVal items As Variant) As Collection
    Dim coll As New Collection
    Dim item As Variant
    For Each item In items
        coll.Add item
    Next item
    Set arrayToCollection = coll
End Function


