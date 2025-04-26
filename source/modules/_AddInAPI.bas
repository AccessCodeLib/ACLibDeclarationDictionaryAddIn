Attribute VB_Name = "_AddInAPI"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Function: StartAddIn
'---------------------------------------------------------------------------------------
'
' Default start of add-in ... open add-in form
'
'---------------------------------------------------------------------------------------
Public Function StartAddIn()
   StartApplication
End Function


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

    With New DeclarationDict

        If Len(DeclDictFilePath) = 0 Then
            DeclDictFilePath = CurrentProject.Path & "\" & CurrentProject.Name & ".DeclarationDict.txt"
        End If

        If Not .LoadFromFile(DeclDictFilePath) Then
           .ImportVBProject CurrentVbProject
           ' ... log info: first export
           .ExportToFile DeclDictFilePath
           RunVcsCheck = "Info: No dictionary data found. A new dictionary has been created."
           Exit Function
        End If

        IntialCnt = .Count
        .ImportVBProject CurrentVbProject

        DiffCnt = .DiffCount
        If DiffCnt = 0 Then
            If .Count <> IntialCnt Then
                StoreDictData = True
            End If
        End If

        If OpenDialogToFixLettercase Then
            If DiffCnt > 0 Then
                SetDeclarationDictTransferReference .Self
                DoCmd.OpenForm "DeclarationDictApiDialog", , , , , acDialog
                DiffCnt = .DiffCount
                If DiffCnt = 0 Then
                    StoreDictData = True
                End If
            End If
        End If

        If StoreDictData Then
            .ExportToFile DeclDictFilePath
        End If

        If DiffCnt > 0 Then
            CheckMsg = .DiffCount & " word" & IIf(.DiffCount > 1, "s", vbNullString) & " with different letter case"
            RunVcsCheck = "Failed: " & CheckMsg
        Else
            RunVcsCheck = True
        End If

   End With

End Function
