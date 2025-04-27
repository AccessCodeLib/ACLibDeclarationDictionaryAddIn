VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassToTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'---------------
' Class to test
'---------------

Public Property Get Info() As String

   Info = "abc"

    Dim Abc As String

End Property

Public Function Method1(ByVal X As Long) As Long

    X = X + 1
    Method1 = X

End Function

Public Function Method2a() As Long

    Dim X As Long

    X = 1 + 1

    If X = 2 Then
        X = 9
    Else
        X = 10
    End If

    Method2a = X

Exit Function

    X = 5 'This line will never be reached.

End Function

Public Function Method3() As Long

 Dim X As Long


End Function
