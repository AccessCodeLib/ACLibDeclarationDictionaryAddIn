VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeclarationDictTestCodemodule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'AccUnit:TestRelated

'Implements CodeModulGenerator

Dim AccUnitX As Long
Private m_AccUnitInfo As String
Public Field As String

Private Const Const1 As String = "abc"

Private Enum TestEnum: TestEnum_P1 = 2: End Enum
Private Type TestType1
   FldA As Long
   FldB As String
   FldC As Boolean
End Type

Private Enum TestEnum2
   TestEnum2_P1 = 2
   TestEnum2_P2 = 3
End Enum

Private Type TestType2
   Fld2A As Long
   Fld2B As String
   FldC As Date
End Type

Private WithEvents m_TextBox As TextBox
Attribute m_TextBox.VB_VarHelpID = -1

Public Event RaiseSomething(ByVal EventParam1 As Variant)

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Private Declare PtrSafe Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long) As Long

Public Event RaiseSomething2(ByVal EventParam1 As Variant, ByVal EventParam2 As Variant)

Private Sub Class_Initialize()
' Dim Class_Initialize_Xyz as String ... ignored!
End Sub

Public Function AccUnitTestFunct(ByVal FuncParam1 As Variant, FuncParam2() As String) As Variant

   Dim FuncVar1 As Variant, FuncVar2
   Dim FuncVar3() ' _
   Dim CommentX as String, CommentY As Long

   Dim FuncVar4 As Long: FuncVar4 = 5

   Dim Dim1 As Long: Dim Dim2

       Dim Counter2 ' _
   Public X asString, Y As Long

End Function

Public Sub ProcWith3DeclLines( _
               ByVal D3P1 As Variant, _
               ByRef D3P2() As String)
'
End Sub

' Declaration of a property procedure in one line:
Friend Property Get Name1() As String: Name1 = "TestName": End Property

Friend Property Let Name2(ByVal NewValue As String)
'
End Property

Friend Property Set PropertySet(ByVal ObjRef As Object)
'
End Property

Private Sub TestMe() ' _
Private Sub ThisIsOnlyAComment(

End Sub

Public Sub VariableParams(ParamArray Args() As Variant)
'
End Sub

Private Static Sub MyStaticSub(Optional ByVal Reset As Boolean = False)
    Static Counter2 As Integer
End Sub


Private Function CodeModulGenerator_CreateCodemodule(ByVal ComponentType As VBIDE.vbext_ComponentType, Optional ByVal Name As String = vbNullString) As VBIDE.VBComponent
'
End Function

Private Sub CodeModulGenerator_InsertDeclarationLine(ByVal Code As String)
'
End Sub

Private Sub CodeModulGenerator_RemoveCodemodule()
'
End Sub
