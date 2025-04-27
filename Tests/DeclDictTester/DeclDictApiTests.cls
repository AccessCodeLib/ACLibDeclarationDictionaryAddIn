VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeclDictApiTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'AccUnit:TestClass

Public Sub RunVcsCheck_WithoutDialogChangedLettercase_CheckReturnMessage()

   Dim AddInCallPath As String
   Dim DictFilePath As String
   Dim Result As Variant
   
   AddInCallPath = CurrentProject.Path & "\ACLibDeclarationDictCore.RunVcsCheck"
   DictFilePath = CurrentProject.Path & "\DeclDictApiTests_DeclDict.txt"
   
   ' ensure run first export => delete file if exists
   If FileExists(DictFilePath) Then
      Kill DictFilePath
   End If
   
   ' 1st call => run first export
   Result = Application.Run(AddInCallPath, False, DictFilePath)
   Assert.That Result, Iz.EqualTo("Info: no export data exists, run first export")
   
   ' 2nd call => no changes => Return value = True
   Result = Application.Run(AddInCallPath, False, DictFilePath)
   Assert.That Result, Iz.EqualTo(True)
   
   ' 3rd call: change letter case => Return "Failed:..."
   ChangeLetterCaseInClassToTest
   Result = Application.Run(AddInCallPath, False, DictFilePath)
   ChangeLetterCaseOfClassToTestMethod1 "Method1"
   
   Assert.That Result, Iz.EqualTo("Failed: 1 word with different letter case")
   
End Sub

Private Function ChangeLetterCaseInClassToTest() As String

   Const ProcName As String = "Method1"
   Const ProcDeclarationText As String = "Public Function Method1(ByVal X As Long) As Long"

   Dim vbp As VBProject
   Dim vbc As VBComponent
   Dim cm As CodeModule
   Dim ProcStartLine As Long
   Dim ProcLineText As String
   Dim ChangeTo As String
   
   
   Set vbp = CodeVBProject
   Set vbc = vbp.VBComponents("ClassToTest")
   Set cm = vbc.CodeModule
   
   ProcStartLine = cm.ProcBodyLine(ProcName, vbext_pk_Proc)
   ProcLineText = cm.Lines(ProcStartLine, 1)
   
   If InStr(1, ProcLineText, ProcName, vbBinaryCompare) > 0 Then
      ChangeTo = LCase(Left(ProcName, 1)) & Mid(ProcName, 2)
   Else
      ChangeTo = ProcName
   End If
   
   ChangeLetterCaseOfClassToTestMethod1 ChangeTo
   
End Function

Private Sub ChangeLetterCaseOfClassToTestMethod1(ByVal ChangeTo As String)

   Dim vbp As VBProject
   Dim vbc As VBComponent
   Dim cm As CodeModule
   Dim CodeStartLine As Long

   Set vbp = CodeVBProject
   Set vbc = vbp.VBComponents.Add(vbext_ct_StdModule)
   Set cm = vbc.CodeModule
   CodeStartLine = cm.CountOfDeclarationLines + 1
   cm.InsertLines CodeStartLine, "Private " & ChangeTo
   vbp.VBComponents.Remove vbc

End Sub

Private Property Get CodeVBProject() As VBIDE.VBProject

   Dim Proj As VBProject
   Dim strCodeDbName As String
   Dim objCodeVbProject As Object

   Set objCodeVbProject = VBE.ActiveVBProject
   strCodeDbName = FileTools.UncPath(CodeDb.Name)
   If objCodeVbProject.FileName <> strCodeDbName Then
      Set objCodeVbProject = Nothing
      For Each Proj In VBE.VBProjects
         If Proj.FileName = strCodeDbName Then
            Set objCodeVbProject = Proj
            Exit For
         End If
      Next
   End If

   Set CodeVBProject = objCodeVbProject

End Property
