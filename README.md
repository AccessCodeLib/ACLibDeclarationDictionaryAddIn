# ACLib Declaration Dictionary
This Add-In lists all VBA declarations (variables, function names, constants, etc.) and ensures consistent letter case.

The idea was born from a discussion ([msaccess-vcs-add-in: issue 599](https://github.com/joyfullservice/msaccess-vcs-addin/issues/599)) about the behavior of the VBA editor, which adapts each existing declaration to the last written capitalization of the same word. This leads to many unnecessary changes in commits when using a version control system. 

![ACLibDeclarationDictionary](https://github.com/user-attachments/assets/0ef05ef7-72aa-4786-a790-74e679df7f24)


# ACLibDeclarationDictionaryCore
This Add-In lists all VBA declarations (variables, function names, constants, etc.) and ensures consistent letter case - Core components

## API

### RunVcsCheck([OpenDialogToFixLettercase], [DeclDictFilePath]) As Variant
Compare lettercase from CurrentVbProject with saved (table/file) dictionary items
 
#### Parameters
- OpenDialogToFixLettercase - (Boolean) - Open dialog to fix lettercase
- DeclDictFilePath - (String) - use alternative file path 

#### Returns
- Boolean (True) ... if DiffCount = 0
- String ... if DiffCount > 0 => "Failed: <lettercase info>"
 
#### Example
```
Dim Result As Variant
Result = Application.Run(Environ("appdata") & "\Microsoft\AddIns\ACLibDeclarationDictCore.RunVcsCheck", True)
If Result = True Then
    Debug.Print "No problems with letter case"
Else
    Debug.Print Result
End If
```

### RunVcsCheckDialog()
Equal to RunVcsCheck(True)
