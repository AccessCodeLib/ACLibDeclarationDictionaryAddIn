VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Interface: ITest
'---------------------------------------------------------------------------------------
'
' Interface for Implement test in DeclarationDictTestCodemodule
'
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

'AccUnit:TestRelated

Public Sub TestMe()
'
End Sub

Public Function TestProc(ByVal TestParam1 As Long)
'
End Function
