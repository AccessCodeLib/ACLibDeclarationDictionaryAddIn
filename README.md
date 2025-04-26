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
