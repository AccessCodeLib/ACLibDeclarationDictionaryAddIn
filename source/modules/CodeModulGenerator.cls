VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeModulGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private m_VBComponent As VBComponent

Public Function CreateCodemodule(ByVal ComponentType As vbext_ComponentType, Optional ByVal Name As String = vbNullString) As VBComponent
    Set m_VBComponent = CurrentVbProject.VBComponents.Add(ComponentType)
    If Len(Name) > 0 Then
        m_VBComponent.Name = Name
    End If
    Set CreateCodemodule = m_VBComponent
End Function

Public Sub InsertDeclarationLine(ByVal Code As String)

    Dim cm As CodeModule
    Dim StartLine As Long

    Set cm = m_VBComponent.CodeModule
    StartLine = cm.CountOfDeclarationLines + 1

    cm.InsertLines StartLine, Code

End Sub

Public Sub RemoveCodemodule()

    CurrentVbProject.VBComponents.Remove m_VBComponent
    Set m_VBComponent = Nothing

End Sub
