Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents m_DeclarationDict As DeclarationDict
Attribute m_DeclarationDict.VB_VarHelpID = -1
Private m_InsertRecordset As DAO.Recordset
Private m_RunAsAPI As Boolean

Friend Property Get ActiveDeclarationDict() As DeclarationDict
   If m_DeclarationDict Is Nothing Then
      Set m_DeclarationDict = New DeclarationDict
   End If
   Set ActiveDeclarationDict = m_DeclarationDict
End Property

Friend Property Set ActiveDeclarationDict(ByRef DeclarationDictRef As DeclarationDict)
   Set m_DeclarationDict = DeclarationDictRef
End Property

Private Sub cmdAPI_Click()
   DoCmd.OpenForm "InfoFormAPI"
End Sub

Private Sub cmdLoadFromFile_Click()
   LoadFromFile
End Sub

Private Sub cmdLoadFromTable_Click()
   LoadFromTable
End Sub

Private Sub cmdRemoveFilter_Click()
   Me.filtWord.Value = Null
   RequeryDictData
End Sub

Private Sub cmdSaveToFile_Click()
   SaveToFile
End Sub

Private Sub cmdSaveToTable_Click()
   SaveToTable
End Sub

Private Sub filtDiff_AfterUpdate()
   If m_RunAsAPI Then
      If Me.filtDiff.Value = 0 Then ' full list
         If TableWordCount < m_DeclarationDict.Count Then
            FillWordTableFromDict False
         End If
      End If
   End If
   RequeryDictData
End Sub

Private Function TableWordCount() As Long
   With CodeDb.OpenRecordset("select Count(*) from tabWords")
      TableWordCount = Nz(.Fields(0), 0)
      .Close
   End With
End Function

Private Sub filtWord_AfterUpdate()
   RequeryDictData
End Sub

Private Sub Form_Load()

   With CurrentApplication
      Me.lblVersionInfo.Caption = .ApplicationFullName & " " & ChrW(&H2022) & " Version " & CurrentApplication.Version
   End With

   CodeDb.Execute "delete from tabWords"

   If Me.OpenArgs = "RunAsApiDialog" Then
      SetApiDialogMode
   End If

End Sub

Private Sub SetApiDialogMode()
   m_RunAsAPI = True
   Set m_DeclarationDict = GetDeclarationDictTransferReference
   FillWordTableFromDict True
End Sub

Private Sub Form_Timer()
   Me.TimerInterval = 0
   Me.lblTableRecInfo.Caption = vbNullString
   Me.lblVersionInfo.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   CodeDb.Execute "delete from tabWords"
End Sub

Private Sub cmdChangeLetterCase_Click()
    ChangeLetterCase
End Sub

Private Sub cmdUpdateDict_Click()
   ActiveDeclarationDict.ImportVBProject CurrentVbProject
   If Not (m_InsertRecordset Is Nothing) Then
      m_InsertRecordset.Close
      Set m_InsertRecordset = Nothing
   End If
   RequeryDictData
End Sub

Private Sub lbDictData_AfterUpdate()
   FillWordVariationsList
End Sub

Private Sub lbVariations_AfterUpdate()
   SetVariationCommitMode
End Sub

Private Sub m_DeclarationDict_WordChanged(ByVal WordKey As String)
   UpdateChangedWord WordKey
End Sub

Private Sub UpdateChangedWord(ByVal WordKey As String)

   Static db As DAO.Database
   If db Is Nothing Then
      Set db = CodeDb
   End If

   Dim UpdateSql As String
   UpdateSql = "Update tabWords set Diff=" & SqlTools.BooleanToSqlText(m_DeclarationDict.IsWordWithChangedLetterCase(WordKey), "True")
   If m_DeclarationDict.VariationsDict.Item(WordKey).Count > 1 Then
      UpdateSql = UpdateSql & ", Variations =" & SqlTools.TextToSqlText(Join(m_DeclarationDict.GetWordVariations(WordKey, True), "|"))
   End If

   UpdateSql = UpdateSql & " where Word=" & SqlTools.TextToSqlText(WordKey)

   db.Execute UpdateSql, dbFailOnError

End Sub

Private Sub m_DeclarationDict_WordInserted(ByVal WordKey As String)

   If m_InsertRecordset Is Nothing Then
      Set m_InsertRecordset = CodeDb.OpenRecordset("tabWords", dbOpenDynaset, dbAppendOnly)
   End If

   m_InsertRecordset.AddNew
   m_InsertRecordset.Fields(0).Value = WordKey
   If m_DeclarationDict.VariationsDict.Item(WordKey).Count > 1 Then
      m_InsertRecordset.Fields(1).Value = Join(m_DeclarationDict.GetWordVariations(WordKey, True), "|")
   End If
   m_InsertRecordset.Fields(2).Value = m_DeclarationDict.IsWordWithChangedLetterCase(WordKey)
   m_InsertRecordset.Update

End Sub

Private Sub ogViewMode_AfterUpdate()
   RequeryDictData
End Sub

Private Sub FillWordTableFromDict(Optional ByVal DiffWordsOnly As Boolean = False)

   CodeDb.Execute "delete from tabWords"

   Dim DictIndex As Long
   Dim DictKey As String
   Dim WordsDict As Dictionary

   If DiffWordsOnly Then
      Set WordsDict = m_DeclarationDict.ToDict(False)
   Else
      Set WordsDict = m_DeclarationDict.WordsDict
   End If

   If m_InsertRecordset Is Nothing Then
      Set m_InsertRecordset = CodeDb.OpenRecordset("tabWords", dbOpenDynaset, dbAppendOnly)
   End If

   For DictIndex = 0 To WordsDict.Count - 1
      DictKey = WordsDict.Keys(DictIndex)
      AddRecord DictKey
   Next

   RequeryDictData

End Sub

Private Sub AddRecord(ByVal WordKey As String)

   m_InsertRecordset.AddNew
   m_InsertRecordset.Fields(0).Value = WordKey
   If m_DeclarationDict.VariationsDict.Item(WordKey).Count > 1 Then
      m_InsertRecordset.Fields(1).Value = Join(m_DeclarationDict.GetWordVariations(WordKey, True), "|")
   End If
   m_InsertRecordset.Fields(2).Value = m_DeclarationDict.IsWordWithChangedLetterCase(WordKey)
   m_InsertRecordset.Update

End Sub

Private Sub RequeryDictData()

   ResetVariations

   Dim SelectSql As String
   Dim CriteriaString As String
   Dim rs As DAO.Recordset

   SelectSql = "SELECT W.Word, W.Variations FROM tabWords AS W"

   With New FilterStringBuilder
      .ConfigSqlFormat "\#yyyy-mm-dd\#", "True", "*"
      .Add "Diff", SQL_Boolean, SQL_Equal, CBool(Me.filtDiff.Value), , False
      .Add "Word", SQL_Text, SQL_Like, Me.filtWord.Value
      CriteriaString = .ToString()
   End With
   If Len(CriteriaString) > 0 Then
      SelectSql = SelectSql & " WHERE " & CriteriaString
   End If

   SelectSql = SelectSql & " ORDER BY W.Word;"

   Me.txtCurrentLetterCase.Value = Null
   Me.lbDictData.Value = Null
   With Me.lbDictData
      .Value = Null
      If .RowSource <> SelectSql Then
         .RowSource = SelectSql
      Else
         .Requery
      End If
   End With

   RefreshDictInfo

End Sub

Private Sub RefreshDictInfo()
   With ActiveDeclarationDict
      Me.txtDictInfo.Value = .DiffCount & " / " & .Count
   End With
End Sub

Private Sub ResetVariations()
   Me.lbVariations.RowSource = vbNullString
   SetVariationCommitMode
End Sub

Private Sub FillWordVariationsList()

   Dim lb As ListBox
   Set lb = Me.lbVariations
   lb.Value = Null
   lb.RowSource = vbNullString

   Dim Items() As String
   Items = ActiveDeclarationDict.GetWordVariations(Me.txtWord.Value)

   Dim i As Long
   For i = LBound(Items) To UBound(Items)
      lb.AddItem Items(i)
   Next

   Me.txtCurrentLetterCase.Value = Items(UBound(Items))

   SetVariationCommitMode

End Sub

Private Sub SetVariationCommitMode()

   Dim EnableChange As Boolean

   EnableChange = IsDifferentLetterCase(Nz(Me.txtCurrentLetterCase.Value, vbNullString), Nz(Me.txtSelectedLetterCase.Value, vbNullString))
   If Not EnableChange Then
      EnableChange = IsDifferentLetterCase(Nz(Me.txtWord.Value, vbNullString), Nz(Me.txtSelectedLetterCase.Value, vbNullString))
   End If

   Me.cmdChangeLetterCase.Enabled = EnableChange

End Sub

Private Property Get IsDifferentLetterCase(ByVal String1 As String, ByVal String2 As String) As Boolean

   If Len(String1) = 0 Or Len(String2) = 0 Then
      IsDifferentLetterCase = False
   ElseIf StrComp(String1, String2, vbTextCompare) <> 0 Then
      IsDifferentLetterCase = False
   ElseIf StrComp(String1, String2, vbBinaryCompare) = 0 Then
      IsDifferentLetterCase = False
   Else
      IsDifferentLetterCase = True
   End If

End Property

Private Sub ChangeLetterCase()

    Dim ChangeFrom As String
    Dim ChangeTo As String

    ChangeFrom = Me.txtCurrentLetterCase.Value
    ChangeTo = Me.txtSelectedLetterCase.Value

    If Not IsDifferentLetterCase(ChangeFrom, ChangeTo) Then
      If IsDifferentLetterCase(Me.txtWord.Value, ChangeTo) Then
         SetCurrentLetterCaseAsDefault ChangeTo
      End If
      Exit Sub
    End If

    With New CodeModulGenerator
        .CreateCodemodule vbext_ct_StdModule
        .InsertDeclarationLine "Private " & ChangeTo
        .RemoveCodemodule
    End With

    ActiveDeclarationDict.AddWord ChangeTo

    FillWordVariationsList

    Me.SetFocus

    If m_RunAsAPI Then
      RequeryDictData
    End If

End Sub

Private Sub SetCurrentLetterCaseAsDefault(ByVal ChangeTo As String)

   ActiveDeclarationDict.ChangeKeyLetterCase ChangeTo

   Dim UpdateSql As String
   Dim VariationsString As String
   VariationsString = Join(m_DeclarationDict.GetWordVariations(ChangeTo, True), "|")
   With SqlTools
      UpdateSql = "update tabWords set Word = " & .TextToSqlText(ChangeTo) & ", Variations=" & .TextToSqlText(VariationsString) & " where Word = " & .TextToSqlText(ChangeTo)
   End With
   CodeDb.Execute UpdateSql, dbFailOnError

   With Me.lbDictData
      .Requery
      .Value = ChangeTo
   End With

   RefreshDictInfo

End Sub

Private Sub SaveToTable()

   With ActiveDeclarationDict
      .SaveToTable DefaultDeclDictTableName
      ShowTableRecInfo .Count & " records were inserted into " & DefaultDeclDictTableName
   End With

End Sub

Private Sub LoadFromTable()

   Dim ErrMsg As String

   If Not ActiveDeclarationDict.LoadFromTable(DefaultDeclDictTableName, ErrMsg) Then
      ShowTableRecInfo ErrMsg
      Exit Sub
   End If

   ShowTableRecInfo ActiveDeclarationDict.Count & " records were imported from " & DefaultDeclDictTableName
   RequeryDictData

End Sub

Private Sub ShowTableRecInfo(ByVal Info As String)
   Me.lblTableRecInfo.Caption = Info
   Me.lblVersionInfo.Visible = False
   Me.TimerInterval = 4000
End Sub

Private Sub SaveToFile()
   With ActiveDeclarationDict
      .ExportToFile CurrentProject.Path & "\" & CurrentProject.Name & ".DeclarationDict.txt"
      ShowTableRecInfo .Count & " records were exported to " & CurrentProject.Name & ".DeclarationDict.txt"
   End With
End Sub

Private Sub LoadFromFile()
   With ActiveDeclarationDict
      .LoadFromFile CurrentProject.Path & "\" & CurrentProject.Name & ".DeclarationDict.txt"
      ShowTableRecInfo .Count & " records were imported from " & CurrentProject.Name & ".DeclarationDict.txt"
   End With
   RequeryDictData
End Sub
