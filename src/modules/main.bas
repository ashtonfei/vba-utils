Attribute VB_Name = "main"
Option Explicit
Const ADDIN_NAME = "Addin Name"

Sub run()
    Dim app As New AppSmcMerge
    Dim inputs As Variant
    inputs = Array("Input1", "Input2")
    Const OUTPUT = "Report"
    Call app.run(inputs, OUTPUT)
End Sub

Sub show_version()
    Dim version As New clsVersion
    Call version.show(ADDIN_NAME)
End Sub

