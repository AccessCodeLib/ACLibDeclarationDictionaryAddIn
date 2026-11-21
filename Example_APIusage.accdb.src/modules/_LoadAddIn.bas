Attribute VB_Name = "_LoadAddIn"
Option Compare Database
Option Explicit

Public Sub LoadAddIn()

'API: Public Function RunVcsCheck(Optional ByVal OpenDialogToFixLettercase As Boolean = False, _
'                                 Optional ByVal DeclDictFilePath As String = vbNullString, _
'                                 Optional ByVal IncludeUsedMembers As Boolean = False) As Variant

   Dim AddInCallPath As String
   AddInCallPath = CurrentProject.Path & "\ACLibDeclarationDictCore.RunVcsCheck"

   Dim Result As Variant
   Result = Application.Run(AddInCallPath, True, vbNullString, True)
   If Result = True Then
      Debug.Print "No problems with letter case"
   Else
      Debug.Print Result
   End If

End Sub
