Attribute VB_Name = "modHelper"
Option Explicit
Const srcPath = "./src"
Const modulesPath = "./src/modules"
Const classesPath = "./src/classes"
Const formsPath = "./src/forms"
Const modHelperName = "modHelper"

' export modules to this workbook
Sub exportModules()
    Call ChDir(ThisWorkbook.path)
    Dim project As VBIDE.VBProject
    Set project = ThisWorkbook.VBProject
    Dim component As VBIDE.VBComponent
    For Each component In project.vbComponents
        If component.Type = vbext_ct_StdModule Then
            If component.name = modHelperName Then
                Call component.Export(srcPath & "/" & component.name & ".bas")
            Else
                Call component.Export(modulesPath & "/" & component.name & ".bas")
            End If
        ElseIf component.Type = vbext_ct_ClassModule Then
            Call component.Export(classesPath & "/" & component.name & ".cls")
        ElseIf component.Type = vbext_ct_MSForm Then
            Call component.Export(formsPath & "/" & component.name & ".frm")
        End If
    Next component
End Sub
' import modules from this workbook
Sub importModules()
    Call ChDir(ThisWorkbook.path)
    Call removeModules
    Dim modules As Collection
    Set modules = getFilePathes(modulesPath, "*.bas")
    Call import(modules)
    
    Dim classes As Collection
    Set classes = getFilePathes(classesPath, "*.cls")
    Call import(classes)
    
    Dim forms As Collection
    Set forms = getFilePathes(classesPath, "*.fr")
    Call import(forms)
End Sub
Private Function import(ByRef pathes As Collection)
    Dim vbComponents As VBIDE.vbComponents
    Set vbComponents = ThisWorkbook.VBProject.vbComponents
    Dim path As Variant
    For Each path In pathes
        Call vbComponents.import(path)
    Next path
End Function

Private Function removeModules()
    Dim components As VBIDE.vbComponents
    Dim component As VBIDE.VBComponent
    Set components = ThisWorkbook.VBProject.vbComponents
    For Each component In components
        If component.Type = vbext_ct_StdModule _
            Or component.Type = vbext_ct_ClassModule _
            Or component.Type = vbext_ct_MSForm Then
            If component.name <> modHelperName Then Call components.Remove(component)
        End If
    Next component
End Function
Private Function getFilePathes(Optional ByVal path As String = "", Optional ByVal pathname As String = "*", Optional ByVal exclude As String) As Collection
    Dim currentPath As String
    currentPath = CurDir
    If path <> "" Then
        Call ChDir(path)
    End If

    Dim filePathes As New Collection
    Dim filename As String
    filename = Dir(pathname)
    While filename <> ""
        filePathes.Add CurDir & "\" & filename
        filename = Dir()
    Wend
    Call ChDir(currentPath)
    Set getFilePathes = filePathes
End Function


