Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents m_DeclarationDict As DeclarationDict
Attribute m_DeclarationDict.VB_VarHelpID = -1
Private m_WordListRecordset As ADODB.Recordset

Friend Property Get ActiveDeclarationDict() As DeclarationDict
   If m_DeclarationDict Is Nothing Then
      Set m_DeclarationDict = New DeclarationDict
   End If
   Set ActiveDeclarationDict = m_DeclarationDict
End Property

Friend Property Set ActiveDeclarationDict(ByRef DeclarationDictRef As DeclarationDict)
   Set m_DeclarationDict = DeclarationDictRef
End Property

Private Sub Form_Load()

   With CurrentApplication
      Me.Caption = .ApplicationName & " " & ChrW(&H2022) & " Differences"
      Me.lblVersionInfo.Caption = .ApplicationFullName & " " & ChrW(&H2022) & " Version " & CurrentApplication.Version
   End With

   SetApiDialogMode

End Sub

Private Sub SetApiDialogMode()
   Set m_DeclarationDict = GetDeclarationDictTransferReference
   FillWordTableFromDict
End Sub

Private Sub cmdChangeLetterCase_Click()
    ChangeLetterCase
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

    If Not m_DeclarationDict.IsWordWithChangedLetterCase(WordKey) Then
       DeleteWordRecord WordKey
    End If

End Sub

Private Sub DeleteWordRecord(ByVal WordKey As String)

    With m_WordListRecordset
        .MoveFirst
        .Find "Word='" & Replace(WordKey, "'", "''") & "'", 0, adSearchForward, 0
        .Delete adAffectCurrent
   End With

End Sub

Private Sub FillWordTableFromDict()

    Dim DictIndex As Long
    Dim DictKey As String
    Dim WordsDict As Dictionary

    If m_DeclarationDict Is Nothing Then
        Exit Sub
    End If

    Set WordsDict = m_DeclarationDict.ToDict(False)

    InitWordRecordset

    For DictIndex = 0 To WordsDict.Count - 1
       DictKey = WordsDict.Keys(DictIndex)
       AddRecord DictKey
    Next

    m_WordListRecordset.Sort = "Word ASC"
    m_WordListRecordset.MoveFirst
    Set Me.lbDictData.Recordset = m_WordListRecordset

End Sub

Private Sub InitWordRecordset()
   Set m_WordListRecordset = New ADODB.Recordset
   m_WordListRecordset.CursorLocation = adUseClient
   m_WordListRecordset.CursorType = adOpenKeyset
   m_WordListRecordset.LockType = adLockOptimistic
   m_WordListRecordset.Fields.Append "Word", adVarChar, 255, adFldUpdatable
   m_WordListRecordset.Fields.Append "Variations", adVarChar, 255, adFldUpdatable
   m_WordListRecordset.Open
End Sub

Private Sub AddRecord(ByVal WordKey As String)

   m_WordListRecordset.AddNew
   m_WordListRecordset.Fields(0).Value = WordKey
   If m_DeclarationDict.VariationsDict.Item(WordKey).Count > 1 Then
      m_WordListRecordset.Fields(1).Value = Join(m_DeclarationDict.GetWordVariations(WordKey, True), "|")
   End If
   m_WordListRecordset.Update

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
            ActiveDeclarationDict.ChangeKeyLetterCase ChangeTo
            DeleteWordRecord ChangeTo
            RefreshDiffData
        End If
        Exit Sub
    End If

    ActiveDeclarationDict.FixLetterCase ChangeTo

    With Me.lbDictData
        .Value = Null
        Set .Recordset = m_WordListRecordset
    End With

    Me.SetFocus
    RefreshDiffData

End Sub

Private Sub RefreshDiffData()

    If m_DeclarationDict.DiffCount = 0 Then
        DoCmd.Close
    Else
        With Me.lbDictData
            Set .Recordset = m_WordListRecordset
            .Value = Me.lbDictData.Column(0, 0)
        End With
        FillWordVariationsList
    End If

End Sub
