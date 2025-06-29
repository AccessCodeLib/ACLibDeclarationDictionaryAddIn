VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeclarationDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

#Const DebugPrintEnabled = 0

Private m_Words As Scripting.Dictionary
Private m_WordVariations As Scripting.Dictionary
Private m_DiffCnt As Long

Public Event WordInserted(ByVal WordKey As String)
Public Event WordChanged(ByVal WordKey As String)

Private Sub Class_Initialize()
   Set m_Words = GetNewDict(TextCompare)
   Set m_WordVariations = GetNewDict(TextCompare)
   m_DiffCnt = -1
End Sub

Friend Property Get Self() As DeclarationDict
   Set Self = Me
End Property

Private Function GetNewDict(ByVal CompareMode As Scripting.CompareMethod) As Scripting.Dictionary
   Dim Dict As Scripting.Dictionary
   Set Dict = New Scripting.Dictionary
   Dict.CompareMode = CompareMode
   Set GetNewDict = Dict
End Function

Friend Property Get WordsDict() As Scripting.Dictionary
   Set WordsDict = m_Words
End Property

Friend Property Get VariationsDict() As Scripting.Dictionary
   Set VariationsDict = m_WordVariations
End Property

Public Property Get Count() As Long
   Count = m_Words.Count
End Property

Public Property Get DiffCount() As Long
   If m_DiffCnt = -1 Then
      m_DiffCnt = GetDiffCnt
   End If
   DiffCount = m_DiffCnt
End Property

Private Function GetDiffCnt() As Long

   Dim Word As String
   Dim i As Long
   Dim DifCntr As Long

   For i = 0 To m_Words.Count - 1
      Word = m_Words.Keys(i)
      If IsChangedItem(Word, m_WordVariations.Item(Word)) Then
         DifCntr = DifCntr + 1
      End If
   Next

   GetDiffCnt = DifCntr

End Function

Public Function IsWordWithChangedLetterCase(ByVal Word As String) As Boolean

   IsWordWithChangedLetterCase = IsChangedItem(Word, m_WordVariations.Item(Word))

End Function

Private Function IsChangedItem(ByVal Word As String, ByVal SubDict As Scripting.Dictionary) As Boolean

   Dim LastKey As String

   If StrComp(Word, m_Words.Item(Word), vbBinaryCompare) <> 0 Then
      Word = m_Words.Item(Word)
   End If

   With SubDict
      If .Count > 1 Then
         LastKey = .Keys(.Count - 1)
         If StrComp(LastKey, Word, vbBinaryCompare) <> 0 Then
            IsChangedItem = True
         End If
      End If
   End With

End Function

Public Sub AddWord(ByVal Word As String)

   Word = Trim(Word)

   If Left(Word, 1) = "[" Then
      If Right(Word, 1) = "]" Then
         Word = Trim(Mid(Word, 2, Len(Word) - 2))
      End If
   End If

   If Len(Word) = 0 Then
      Exit Sub
   End If

   m_DiffCnt = -1

   Dim SubDict As Scripting.Dictionary
   Dim LastKey As String
   Dim RaiseInsert As Boolean
   Dim RaiseChanged As Boolean

   If Not m_Words.Exists(Word) Then
      m_Words.Add Word, Word
      m_WordVariations.Add Word, GetNewDict(BinaryCompare)
      m_WordVariations.Item(Word).Add Word, Word
      RaiseEvent WordInserted(Word)
      Exit Sub
   End If

   Set SubDict = m_WordVariations.Item(Word)
   With SubDict
      If Not .Exists(Word) Then
         .Add Word, Word
         RaiseChanged = True
      ElseIf .Count > 1 Then
         LastKey = .Keys(.Count - 1)
         If StrComp(LastKey, Word, vbBinaryCompare) <> 0 Then
            .Remove Word
            .Add Word, Word
         End If
         RaiseChanged = True
      End If
   End With

   If RaiseChanged Then
      RaiseEvent WordChanged(Word)
   End If

End Sub

Public Function ToString(Optional ByVal ShowAll As Boolean = False) As String

   Dim WordKey As Variant
   Dim WordIndex As Long

   Dim VariationsDict As Scripting.Dictionary

   Dim OutputString As String

   For WordIndex = 0 To m_Words.Count - 1

      WordKey = m_Words.Keys(WordIndex)
      Set VariationsDict = m_Words.Item(WordKey)

      If VariationsDict.Count > (1 - Abs(ShowAll)) Then
         OutputString = OutputString & vbNewLine & WordKey & ":" & GetWordVariationsOutputString(WordKey)
      End If

   Next

   ToString = Mid(OutputString, Len(vbNewLine) + 1)

End Function

Public Function ToDict(Optional ByVal ShowAll As Boolean = False) As Scripting.Dictionary

   Dim WordKey As Variant
   Dim WordIndex As Long
   Dim OutputWord As Boolean
   Dim VariationsString As String

   Dim VariationsDict As Scripting.Dictionary

   Dim OutputDict As Scripting.Dictionary
   Set OutputDict = New Scripting.Dictionary

   OutputWord = ShowAll

   For WordIndex = 0 To m_Words.Count - 1

      WordKey = m_Words.Keys(WordIndex)
      Set VariationsDict = m_WordVariations.Item(WordKey)

      If Not ShowAll Then
         OutputWord = IsChangedItem(WordKey, VariationsDict)
      End If
      If OutputWord Then
         If VariationsDict.Count > 1 Then
            VariationsString = GetWordVariationsOutputString(WordKey)
         Else
            VariationsString = vbNullString
         End If
         OutputDict.Add WordKey, VariationsString
      End If

   Next

   Set ToDict = OutputDict

End Function

Private Function GetWordVariationsOutputString(ByVal Word As String, Optional ByVal IgnoreOriginalWord As Boolean = True) As String

   GetWordVariationsOutputString = Join(GetWordVariations(Word, IgnoreOriginalWord), "|")

End Function

Public Function GetWordVariations(ByVal Word As String, Optional ByVal IgnoreOriginalWord As Boolean = False) As String()

   Dim Variations() As String
   Dim VariantWord As String
   Dim AppendVariant As Boolean
   Dim i As Long, k As Long
   Dim VariationsDict As Scripting.Dictionary

   If StrComp(Word, m_Words.Item(Word), vbBinaryCompare) <> 0 Then
      Word = m_Words.Item(Word)
   End If

   Set VariationsDict = m_WordVariations.Item(Word)
   ReDim Variations(VariationsDict.Count - 1 - Abs(IgnoreOriginalWord))

   AppendVariant = Not IgnoreOriginalWord

   For i = 0 To VariationsDict.Count - 1
      VariantWord = VariationsDict.Keys(i)
      If IgnoreOriginalWord Then
         AppendVariant = StrComp(Word, VariantWord, vbBinaryCompare) <> 0
      End If
      If AppendVariant Then
         Variations(k) = VariantWord
         k = k + 1
      End If
   Next

   GetWordVariations = Variations

End Function

Public Sub ChangeKeyLetterCase(ByVal WordWithNewLetterCase As String)

   With m_Words
      .Key(WordWithNewLetterCase) = "~ " & WordWithNewLetterCase & " ~"
      .Key("~ " & WordWithNewLetterCase & " ~") = WordWithNewLetterCase
      .Item(WordWithNewLetterCase) = WordWithNewLetterCase
   End With

   m_DiffCnt = -1

End Sub

Public Sub FixLetterCase(ByVal WordWithNewLetterCase As String)

    With New CodeModulGenerator
        .CreateCodemodule vbext_ct_StdModule
        If Left(WordWithNewLetterCase, 1) = "_" Then
           WordWithNewLetterCase = "[" & WordWithNewLetterCase & "]"
        End If
        .InsertDeclarationLine "Private " & WordWithNewLetterCase
        .RemoveCodemodule
    End With

    AddWord WordWithNewLetterCase

End Sub

Public Function LoadFromTable(ByVal TableName As String, Optional ByRef ErrorMsg As String) As Boolean

   Dim db As DAO.Database
   Dim rs As DAO.Recordset
   Dim i As Long

   Set db = CurrentDb

   If Not TableDefExists(TableName, db) Then
      ErrorMsg = "Table " & TableName & " not exists"
      LoadFromTable = False
      Exit Function
   End If

   Set rs = db.OpenRecordset("select DeclWord from " & TableName, dbOpenSnapshot)
   Do While Not rs.EOF
      AddWord rs.Fields(0)
      rs.MoveNext
   Loop
   rs.Close
   Set db = Nothing

   LoadFromTable = True

End Function

Public Sub SaveToTable(ByVal TableName As String)

   Dim db As DAO.Database
   Set db = CurrentDb

   If Not TableDefExists(TableName, db) Then
      db.Execute "create table " & TableName & " (DeclWord varchar(255) Not Null Primary Key)", dbFailOnError
   End If

   Dim rs As DAO.Recordset
   Dim i As Long

   db.Execute "delete from " & TableName
   Set rs = db.OpenRecordset(TableName, dbOpenDynaset, dbAppendOnly)
   For i = 0 To m_Words.Count - 1
      rs.AddNew
      rs.Fields(0).Value = m_Words.Keys(i)
      rs.Update
   Next

End Sub

Public Function TableDefExists(ByVal TableDefName As String, _
                      Optional ByVal DbRef As DAO.Database = Nothing) As Boolean

   Dim rst As DAO.Recordset
   Dim FilterString As String
   Dim ObjectTypeFilterString As String

   If DbRef Is Nothing Then
      Set DbRef = CodeDb
   End If

   FilterString = "where Name = '" & Replace(TableDefName, "'", "''") & "'" & _
                  " and Type IN (1, 4, 6)"

   Set rst = DbRef.OpenRecordset("select Name from MSysObjects " & FilterString, dbOpenForwardOnly, dbReadOnly)
   TableDefExists = Not rst.EOF
   rst.Close

End Function


Public Function LoadFromFile(ByVal FullFileName As String, Optional ByRef ErrorMsg As String) As Boolean

   If Not FileExists(FullFileName) Then
      LoadFromFile = False
      ErrorMsg = "File " & FullFileName & " not exists"
      Exit Function
   End If

   With New ADODB.Stream
      .Charset = "utf-8"
      .Open
      .LoadFromFile FullFileName
      Do Until .EOS
          AddWord .ReadText(StreamReadEnum.adReadLine)
      Loop
      .Close
   End With

   LoadFromFile = True

End Function

Public Sub ExportToFile(ByVal FullFileName As String, Optional ByVal ExportAsSortedList As Boolean = True)

   Dim WordArray() As Variant
   Dim i As Long

   WordArray = m_Words.Keys

   If ExportAsSortedList Then
      QuickSort WordArray, LBound(WordArray), UBound(WordArray)
   End If

   With New ADODB.Stream
      .Type = 2
      .Charset = "utf-8"
      .Open
      For i = LBound(WordArray) To UBound(WordArray)
         .WriteText WordArray(i), adWriteLine
      Next
      .SaveToFile FullFileName, adSaveCreateOverWrite
      .Close
   End With

End Sub


#If DebugPrintEnabled Then

Private Sub DebugPrint(ByVal Text2Output As String, _
              Optional ByVal PrintHeaderLine As Boolean = False, _
              Optional ByVal HeaderLineCaption As String = vbNullString)

   Const HeaderLen As Long = 40
   Dim StringReplCount As Long
   If PrintHeaderLine Then
      StringReplCount = (HeaderLen - Len(HeaderLineCaption)) / 2
      If StringReplCount < 2 Then
         StringReplCount = 2
      End If
      Debug.Print String(StringReplCount, "-") & HeaderLineCaption & String(StringReplCount, "-")
   End If

   Debug.Print Text2Output

End Sub

#End If
