Attribute VB_Name = "modApiDialogSupport"
Option Compare Text
Option Explicit
Option Private Module

Public Const DeclDictAddInName As String = "ACLib Declaration Dictionary"
Public Const DeclDictVersion As String = "0.5.0"

Private m_DeclarationDict As DeclarationDict

Public Function GetDeclarationDictTransferReference() As DeclarationDict
   Set GetDeclarationDictTransferReference = m_DeclarationDict
   Set m_DeclarationDict = Nothing
End Function

Public Sub SetDeclarationDictTransferReference(ByVal DicRef As DeclarationDict)
   Set m_DeclarationDict = DicRef
End Sub
