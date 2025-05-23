﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddInInstaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_ConfigData As AddInConfiguration

Public Function InstallAddIn(ByVal AddInConfigData As AddInConfiguration, _
                    Optional ByVal CompileAddIn As Boolean = False, _
                    Optional ByRef CompletedMsg As String) As Boolean

   Dim AddInFileInstalled As Boolean

   Set m_ConfigData = AddInConfigData

   If CompileAddIn Then
      AddInFileInstalled = CreateAccde(GetSourceFileFullName, GetDestFileFullName)
      If AddInFileInstalled Then
         CompletedMsg = "Add-In was compiled and saved in '" + GetAddInLocation + "'."
      Else
         CompletedMsg = "Error! Compiled file was not created."
      End If
   Else
      DeleteAddInFiles
      AddInFileInstalled = TryFileCopy(GetSourceFileFullName, GetDestFileFullName)
      If AddInFileInstalled Then
         CompletedMsg = "Add-In was saved in '" + GetAddInLocation + "'."
      Else
         CompletedMsg = "Error! File was not copied."
      End If
   End If

   If AddInFileInstalled = True Then
      RegisterAddIn GetDestFileFullName()
   End If

   InstallAddIn = AddInFileInstalled

End Function

Public Property Get AddInName() As String
   AddInName = m_ConfigData.AddInRegPathName
End Property

Public Property Get AddInFileName() As String
   AddInFileName = m_ConfigData.FileName
End Property

Public Property Get MsgBoxTitle() As String
   MsgBoxTitle = "Install " & AddInName
End Property

Function GetSourceFileFullName()
   GetSourceFileFullName = CurrentDb.Name
End Function

Function GetDestFileFullName()
   GetDestFileFullName = GetAddInLocation & AddInFileName
End Function

Function GetAddInLocation()
   GetAddInLocation = GetAppDataLocation & "Microsoft\AddIns\"
End Function

Function GetAppDataLocation()
   GetAppDataLocation = Environ("APPDATA") & "\"
End Function

Function DeleteAddInFiles()

   Dim DestFile As String
   DestFile = GetDestFileFullName()
   DeleteFile DestFile

End Function

Function DeleteFile(File2Delete)
   If FileTools.FileExists(File2Delete) Then
      Kill File2Delete
   End If
End Function

Private Function TryFileCopy(ByVal SourceFilePath As String, ByVal DestFilePath As String) As Boolean

On Error Resume Next

   With CreateObject("Scripting.FileSystemObject")
      If .FileExists(DestFilePath) Then
         .DeleteFile DestFilePath, True
      End If
      .CopyFile SourceFilePath, DestFilePath, True
   End With

   If Err.Number <> 0 Then
      Err.Clear
      TryFileCopy = False
   Else
      TryFileCopy = True
   End If

End Function

Function CreateAccde(ByVal SourceFilePath As String, ByVal DestFilePath As String) As Boolean

   Dim FileToCompile As String
   Dim AccessApp As Access.Application

   DeleteAddInFiles

   FileToCompile = DestFilePath & ".accdb"
   If Not TryFileCopy(SourceFilePath, FileToCompile) Then
      Exit Function
   End If

   Set AccessApp = CreateObject("Access.Application")
   AccessApp.SysCmd 603, (FileToCompile), (DestFilePath)

   DeleteFile FileToCompile

   CreateAccde = True

End Function


'##################################################
' Register Menu Add-In

Function RegisterAddIn(AddInFile)

   Dim AddInDb As DAO.Database
   Dim rst As DAO.Recordset
   Dim ItemValue As Variant
   Dim wsh As Object

   Set AddInDb = DBEngine.OpenDatabase(AddInFile)

   Set wsh = CreateObject("WScript.Shell")
   Set rst = AddInDb.OpenRecordset("select Subkey, ValName, Type, Value from USysRegInfo where ValName > '' Order By ValName", 8) 'dbOpenForwardOnly=8
   Do While Not rst.EOF
        ItemValue = rst.Fields("Value").Value
        If Len(ItemValue) > 0 Then
        If InStr(1, ItemValue, "|ACCDIR") > 0 Then
            ItemValue = AddInDb.Name
        End If
        End If
        RegisterMenuAddInItem wsh, rst.Fields("Subkey").Value, rst.Fields("ValName").Value, rst.Fields("Type").Value, ItemValue
        rst.MoveNext
   Loop
   rst.Close

   AddInDb.Close

End Function

Function RegisterMenuAddInItem(wsh, ByVal SubKey, ByVal ItemValName, ByVal RegType, ByVal ItemValue)
    Dim RegName
    RegName = GetRegistryPath(SubKey)
    With wsh
        If Len(ItemValName) > 0 Then
            RegName = RegName & "\" & ItemValName
        End If
        .RegWrite RegName, ItemValue, GetRegTypeString(RegType)
    End With
End Function

Function GetRegTypeString(ByVal RegType)
    Select Case RegType
        Case 1
            GetRegTypeString = "REG_SZ"
        Case 4
            GetRegTypeString = "REG_DWORD"
        Case 0
            GetRegTypeString = vbNullString
        Case Else
            Err.Raise vbObjectError, "GetRegTypeString", "RegType not supported"
    End Select
End Function

Function GetRegistryPath(SubKey)
    GetRegistryPath = Replace(SubKey, "HKEY_CURRENT_ACCESS_PROFILE", HkeyCurrentAccessProfileRegistryPath())
End Function

Function HkeyCurrentAccessProfileRegistryPath()
    HkeyCurrentAccessProfileRegistryPath = "HKCU\SOFTWARE\Microsoft\Office\" & Access.Application.Version & "\Access"
End Function
