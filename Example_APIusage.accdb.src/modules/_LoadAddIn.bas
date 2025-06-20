Attribute VB_Name = "_LoadAddIn"
Option Compare Database
Option Explicit

Public Sub LoadAddIn()

'API: RunVcsCheck(Optional ByVal OpenDialogToFixLettercase As Boolean = False)

   Dim AddInCallPath As String
   AddInCallPath = CurrentProject.Path & "\ACLibDeclarationDictCore.RunVcsCheck"

   Dim Result As Variant
   Result = Application.Run(AddInCallPath, True)
   If Result = True Then
      Debug.Print "No problems with letter case"
   Else
      Debug.Print Result
   End If

End Sub
