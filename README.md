# ACLib Declaration Dictionary

This Add-In lists all VBA declarations (variables, function names, constants, etc.) and ensures consistent letter case.

The idea was born from a discussion ([msaccess-vcs-add-in: issue 599](https://github.com/joyfullservice/msaccess-vcs-addin/issues/599)) about the behavior of the VBA editor, which adapts each existing declaration to the last written capitalization of the same word. This leads to many unnecessary changes in commits when using a version control system. 

![ACLibDeclarationDictionary](https://github.com/user-attachments/assets/0ef05ef7-72aa-4786-a790-74e679df7f24)


# ACLibDeclarationDictionaryCore

Core components: https://github.com/AccessCodeLib/ACLibDeclarationDictionaryCore

## Branches

| Branch | Purpose |
|---|---|
| `main` | Release line of this add-in |
| `add-in-dev` | Development of the add-in |
| `core` | Mirror of the `main` branch of ACLibDeclarationDictionaryCore |

Core changes are developed in the Core repository and taken over 1:1 into this repository; add-in specific code is added on top.
Core updates are merged, not rebased: `core` is updated from Core `main`, then merged into `add-in-dev` (preferably via a `feature/core-<version>` branch and a pull request).

## API

### RunVcsCheck([OpenDialogToFixLettercase], [DeclDictFilePath], [IncludeUsedMembers]) As Variant
Compare lettercase from CurrentVbProject with saved file dictionary items
 
#### Parameters
- OpenDialogToFixLettercase - (Boolean) - Open dialog to fix lettercase
- DeclDictFilePath - (String) - use alternative dictionary file path
- IncludeUsedMembers - (Boolean) - also add used members to the dictionary (default: False)

#### Returns
- Boolean (True) ... if DiffCount = 0
- String ... if DiffCount > 0 => "Failed: <lettercase info>"
- String ... if dict file not exists => "Info: No dictionary data found. A new dictionary has been created."
 
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
Equal to RunVcsCheck(True, , True)
