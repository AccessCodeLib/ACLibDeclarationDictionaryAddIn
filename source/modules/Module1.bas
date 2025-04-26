Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

Private Sub testn()

    Dim props As Object
    Dim prop As Object

    Set props = CurrentProject.Properties
    For Each prop In props
        If prop.Name = "VCS Build Path" Then
            props.Remove prop.Name
        End If

    Next

End Sub
