VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHttp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const RunAsync As Boolean = True
Const ProcessComplete As Integer = 4

Const JS_ARRAY_LENGTH As String = "function getArrayLength(items){return items.length;}"

Public Function fetch(ByVal URL As String, ByVal METHOD As String) As String
    Dim request As MSXML2.XMLHTTP60
    Set request = New MSXML2.XMLHTTP60
    Dim response As String
    With request
        .Open METHOD, URL, RunAsync
        .setRequestHeader "Cotent-Type", "application/json"
        .send
        
        Do While .readyState <> ProcessComplete
            DoEvents
        Loop
        response = .responseText
    End With
    fetch = response
End Function

Public Function parseJSON(ByVal value As String) As Object
    Dim script As MSScriptControl.ScriptControl
    Set script = New MSScriptControl.ScriptControl
    script.Language = "JScript"
    
    Dim items As Object
    Set items = script.Eval(value)
    
    Set parseJSON = items
End Function

