Attribute VB_Name = "_AddInAPI"
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Function: API
'---------------------------------------------------------------------------------------
'
'  Open API Information Form
'
'---------------------------------------------------------------------------------------
Public Function API()
   DoCmd.OpenForm "InfoFormAPI"
End Function


'---------------------------------------------------------------------------------------
' Function: RunVcsCheckDialog
'---------------------------------------------------------------------------------------
'
'  Equal to RunVcsCheck(True)
'
'---------------------------------------------------------------------------------------
Public Function RunVcsCheckDialog() As Variant
   RunVcsCheckDialog = RunVcsCheck(True)
End Function


'---------------------------------------------------------------------------------------
' Function: RunVcsCheck
'---------------------------------------------------------------------------------------
'
' Compare lettercase from CurrentVbProject with saved (table/file) dictionary items
'
' Parameters:
'     OpenDialogToFixLettercase - (Boolean) - Open dialog to fix lettercase
'
' Returns:
'      Boolean (True) ... if DiffCount = 0
'      String         ... if DiffCount > 0 => "Failed: <lettercase info>"
'      String         ... if dict file not exists => "Info: No dictionary data found. A new dictionary has been created."
'
'---------------------------------------------------------------------------------------
Public Function RunVcsCheck(Optional ByVal OpenDialogToFixLettercase As Boolean = False, _
                            Optional ByVal DeclDictFilePath As String = vbNullString) As Variant

    Dim CheckMsg As String
    Dim DiffCnt As Long
    Dim UseTable As Boolean
    Dim StoreDictData As Boolean
    Dim IntialCnt As Long

    Dim DeclDict As DeclarationDict
    Set DeclDict = New DeclarationDict

    If Len(DeclDictFilePath) = 0 Then
        DeclDictFilePath = CurrentProject.Path & "\" & CurrentProject.Name & ".DeclarationDict.txt"
    End If

    If Not DeclDict.LoadFromFile(DeclDictFilePath) Then
       ImportVBProject CurrentVbProject, DeclDict
       ' ... log info: first export
       DeclDict.ExportToFile DeclDictFilePath
       RunVcsCheck = "Info: No dictionary data found. A new dictionary has been created."
       Exit Function
    End If

    IntialCnt = DeclDict.Count
    ImportVBProject CurrentVbProject, DeclDict

    DiffCnt = DeclDict.DiffCount
    If DiffCnt = 0 Then
        If DeclDict.Count <> IntialCnt Then
            StoreDictData = True
        End If
    End If

    If OpenDialogToFixLettercase Then
        If DiffCnt > 0 Then
            SetDeclarationDictTransferReference DeclDict
            DoCmd.OpenForm "DeclarationDictApiDialog", , , , , acDialog
            DiffCnt = DeclDict.DiffCount
            If DiffCnt = 0 Then
                StoreDictData = True
            End If
        End If
    End If

    If StoreDictData Then
        DeclDict.ExportToFile DeclDictFilePath
    End If

    If DiffCnt > 0 Then
        CheckMsg = DeclDict.DiffCount & " word" & IIf(DeclDict.DiffCount > 1, "s", vbNullString) & " with different letter case"
        RunVcsCheck = "Failed: " & CheckMsg
    Else
        RunVcsCheck = True
    End If

End Function

Private Sub ImportVBProject(ByVal VbProjectToImport As VBIDE.VBProject, ByVal DeclDict As DeclarationDict)

   With New CodemoduleDeclarationReader
      .ImportVBProject VbProjectToImport, DeclDict
   End With

End Sub
