Attribute VB_Name = "_LoadAddIn"
'---------------------------------------------------------------------------------------
' Module: _LoadAddIn
'---------------------------------------------------------------------------------------
'
' API examples
'
'-------------------------------------------------------------------------
Option Compare Database
Option Explicit

Public Sub LoadAddIn_RunVcsCheck()

'API: RunVcsCheck(Optional ByVal OpenDialogToFixLettercase As Boolean = False)

   Dim AddInCallPath As String
   AddInCallPath = CurrentProject.Path & "\ACLibDeclarationDict.RunVcsCheck"

   Dim Result As Variant
   Result = Application.Run(AddInCallPath, True)
   If Result = True Then
      Debug.Print "No problems with letter case"
   Else
      Debug.Print Result
   End If

End Sub
