# VBA Utilities

A repo for my daily use VBA utilities. You can download all scripts from [vba-utils.xlam](https://github.com/ashtonfei/vba-utils/blob/main/vba-utils.xlam), or download the modules from the [./src](https://github.com/ashtonfei/vba-utils/tree/main/src) folder.

### Helper Module [./src/modHelper.bas](https://github.com/ashtonfei/vba-utils/blob/main/src/modHelper.bas)

```vba
Call importModules
```

Import VBA components from ./src/modules, ./src/forms, ./src/classes into VBA Project.

```vba
Call exportModules
```

Export VBA components from VBA project to ./src/modules, ./src/forms, ./src/classes.

### VBA Standard Modules [./src/modules/\*.bas](https://github.com/ashtonfei/vba-utils/tree/main/src/modules)

VBA modules

### VBA Class Modules [./src/classes/\*.cls](https://github.com/ashtonfei/vba-utils/tree/main/src/classes)

VBA Class Modules

[clsJson](https://github.com/ashtonfei/vba-utils/tree/main/src/classes/clsJson.cls) A JSON like class module, read data from an Excel sheet and store to the JSON like object.

```vba
Public Sub DEMO_JSON()
    ' Let's say we have a sheet named "Users" and have headers "ID", "Name", "Age", "Email" at row 1
    Dim users As New clsJson
    Call users.init(sheetName:="Users", keyColumnName:="ID", headerRowIndex:=1)
    ' get user data
    Debug.Print "Name:" & Space(4) & users.getValue(key:="1", columnName:="Name")
    Debug.Print "Email:" & Space(4) & users.getValue("1", "Email")
    Debug.Print "Age:" & Space(4) & users.data("1")("Age")
End Sub
```

### VBA Form Modules [./src/forms/\*.frm](https://github.com/ashtonfei/vba-utils/tree/main/src/forms)

VBA Form Modules
