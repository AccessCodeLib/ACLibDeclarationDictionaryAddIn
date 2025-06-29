VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodemoduleDeclarationReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private m_DeclDict As DeclarationDict

Public Sub ImportVBProject(ByVal VBProject2Import As VBProject, ByVal DeclDict As DeclarationDict)

   Dim vbc As VBComponent

   For Each vbc In VBProject2Import.VBComponents
      ImportVBComponent vbc, DeclDict
   Next

End Sub

Public Sub ImportVBComponent(ByVal VBComponent2Import As VBComponent, ByVal DeclDict As DeclarationDict)
   ImportCodeModule VBComponent2Import.CodeModule, DeclDict
End Sub

Public Sub ImportCodeModule(ByVal CodeModule2Import As CodeModule, ByVal DeclDict As DeclarationDict)
   If CodeModule2Import.CountOfLines > 0 Then
      ImportCode CodeModule2Import.Lines(1, CodeModule2Import.CountOfLines), DeclDict
   End If
End Sub

Public Sub ImportCode(ByVal Code As String, ByVal DeclDict As DeclarationDict)

   Dim RegEx As RegExp
   Set RegEx = NewRegExp

   Set m_DeclDict = DeclDict

   Code = PrepareCode(Code, RegEx)

   Const ProcIndex As Long = 0
   Const EnumTypeIndex As Long = 1

   Dim Patterns(2) As String

   Patterns(ProcIndex) = "(?:\r|\n|^)\s*(?:Sub|Function|Property Get|Property Let|Property Set|Event)\s+([^\r\n]*)"
   Patterns(EnumTypeIndex) = "(?:\r|\n|^)\s*(?:Enum|Type)([\s\S]*?)(?:End\s+(?:Enum|Type))"
   Patterns(2) = "(?:\r|\n|^)\s*(?:Dim|ReDim|Private|Friend|Public|Const|Static|Global|Implements)\s+([^\r\n]*)"

   Dim i As Long
   For i = 0 To UBound(Patterns)
      RegEx.Pattern = Patterns(i)
      AddFromCode Code, RegEx, i = ProcIndex, i = EnumTypeIndex
   Next

End Sub

Private Function NewRegExp() As RegExp

   Dim RegEx As RegExp

   Set RegEx = New RegExp
   RegEx.IgnoreCase = True
   RegEx.Global = True

   Set NewRegExp = RegEx

End Function

Private Function PrepareCode(ByVal Code As String, ByVal RegEx As RegExp) As String

   Code = Replace(Code, " _" & vbNewLine, " ")

#If DebugPrintEnabled Then
   DebugPrint Code, True, "PrepareCode - after Replace ' _' & vbNewLine"
#End If

   With RegEx

      ' clear all strings
      .Pattern = """[^""\r\n]*"""
      Code = .Replace(Code, "")

      ' remove comments
      '.Pattern = "'(.*)[\r\n]"
      .Pattern = "'(.*)([\r\n]|$)"
      Code = .Replace(Code, "$2")

#If DebugPrintEnabled Then
   DebugPrint Code, True, "PrepareCode- after remove comments"
#End If

      ' treat line labels as dim (but not line numbers)
      .Pattern = "([\r\n]|^)([^0-9\r\n]\S*):(\s|[\r\n]|$)"
      Code = .Replace(Code, "$1Dim $2:$3")

      ' dim a as String: a = 5  => insert line break
      .Pattern = "(\:\s)"
      Code = .Replace(Code, vbNewLine)

      ' remove Withevents
      .Pattern = "(Public|Private|Friend|Global)\sWithEvents\s"
      Code = .Replace(Code, "$1 ")

      ' API declaration => convert to normal procedure declaration
      .Pattern = "(?:Declare PtrSafe)\s(Function|Sub)\s"
      Code = .Replace(Code, "Declare $1 ")

      RegEx.Pattern = "(?:Declare)\s(Function|Sub)\s([^ ]*)[^(]*\("
      Code = .Replace(Code, "$1 $2(")

      ' remove Static before Function, Sub, Property, ..
      .Pattern = "(?:Static)+\s(Function|Sub|Property)\s"
      Code = .Replace(Code, "$1 ")

      ' remove Public, Private, Friend before Function, Sub, Property, ..
      .Pattern = "(?:Public|Private|Friend|Global)\s(Function|Sub|Property|Event|Enum|Type|Const)\s"
      Code = .Replace(Code, "$1 ")

   End With

#If DebugPrintEnabled Then
   DebugPrint Code, True, "PrepareCode - completed"
#End If

   PrepareCode = Code

End Function

Private Sub AddFromCode(ByVal Code As String, ByVal RegEx As RegExp, ByVal IsProcedure As Boolean, ByVal IsEnumTypeBlock As Boolean)

   Dim Match As Match
   Dim i As Long

   For Each Match In RegEx.Execute(Code)
      For i = 0 To Match.SubMatches.Count - 1
         AddWordFromDeclaration Match.SubMatches(i), IsProcedure, IsEnumTypeBlock
      Next
   Next

End Sub

Private Sub AddWordFromDeclaration(ByRef Declarations As String, ByVal IsProcedure As Boolean, ByVal IsEnumTypeBlock As Boolean)

   Dim Word As String
   Dim i As Long
   Dim Pos As Long
   Dim PosX As Long

   Dim DeclArray() As String

   If IsEnumTypeBlock Then
      Declarations = Replace(Declarations, vbCr, ",")
      Declarations = Replace(Declarations, vbLf, vbNullString)
   End If

   Declarations = Trim(Declarations)

   If IsProcedure Then
     ' Debug.Print Declarations
      Declarations = Replace(Declarations, "()", vbNullString)
      Declarations = Replace(Declarations, "Optional ", vbNullString)
      Declarations = Replace(Declarations, "ByRef ", vbNullString)
      Declarations = Replace(Declarations, "ByVal ", vbNullString)
      Declarations = Replace(Declarations, "ParamArray ", vbNullString)

      Pos = InStr(1, Declarations, "(")
      If Pos > 0 Then
         Mid(Declarations, Pos, 1) = ","
      End If
      Declarations = Replace(Declarations, ")", vbNullString)
   End If

   Do While InStr(1, Declarations, "  ") > 0
      Declarations = Replace(Declarations, "  ", " ")
   Loop

   If Not (IsProcedure Or IsEnumTypeBlock) Then
      Do While Declarations Like "*(*,*)*"
         ' prevent multi-dimensional Dim from transforming into new declarations (might be numeric)
         Pos = InStr(1, Declarations, "(")
         PosX = InStr(Pos, Declarations, ")")
         Declarations = Left(Declarations, Pos - 1) & " " & Mid(Declarations, PosX + 1)
      Loop
   End If

   DeclArray = Split(Trim(Declarations), ",")

   For i = LBound(DeclArray) To UBound(DeclArray)
      Word = Trim(DeclArray(i))
      Pos = CutPos(Word)
      If Pos > 1 Then
         Word = Trim(Left(Word, Pos - 1))
      End If
      If Len(Word) > 0 Then
         m_DeclDict.AddWord Word
      End If
   Next

End Sub

Private Function CutPos(ByVal Expression As String) As Long

   Dim Pos As Long
   Dim PosX As Long

   Const CutChars As String = " ()"

   Dim i As Long

   For i = 1 To Len(CutChars)
      PosX = InStr(1, Expression, Mid(CutChars, i, 1))
      If PosX > 0 Then
         If Pos = 0 Or PosX < Pos Then
            Pos = PosX
         End If
      End If
   Next

   CutPos = Pos

End Function
