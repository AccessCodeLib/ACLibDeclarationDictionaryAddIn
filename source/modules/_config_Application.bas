Attribute VB_Name = "_config_Application"
'---------------------------------------------------------------------------------------
' Modul: _config_Application
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'<codelib>
'  <license>_codelib/license.bas</license>
'  <use>%AppFolder%/source/defGlobal_AccUnitLoader.bas</use>
'  <use>base/modApplication.bas</use
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/modErrorHandler.bas</use>
'  <use>_codelib/addins/shared/AccUnitConfiguration.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

'Version number
Private Const APPLICATION_VERSION As String = "0.4.2.250426"

Private Const APPLICATION_NAME As String = "ACLib Declaration Dictionary"
Private Const APPLICATION_FULLNAME As String = "Access-CodeLib - Declaration Dictionary"
Private Const APPLICATION_TITLE As String = APPLICATION_NAME

Private Const APPLICATION_STARTFORMNAME As String = "DeclarationDictForm"

Private m_Extensions As Object 'ApplicationHandler_ExtensionCollection

Public Const DefaultDeclDictTableName As String = "USysDeclDict"

'---------------------------------------------------------------------------------------
' Sub: InitConfig
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Konfigurationseinstellungen initialisieren
' </summary>
' <param name="oCurrentAppHandler">Möglichkeit einer Referenzübergabe, damit nicht CurrentApplication genutzt werden muss</param>
' <returns></returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub InitConfig(Optional ByRef CurrentAppHandlerRef As Object = Nothing)

'----------------------------------------------------------------------------
' Anwendungsinstanz einstellen
'
   If CurrentAppHandlerRef Is Nothing Then
      Set CurrentAppHandlerRef = modApplication.CurrentApplication
   End If

   With CurrentAppHandlerRef

      'Zur Sicherheit AccDb einstellen
      Set .AppDb = Application.CodeDb 'muss auf CodeDb zeigen,
                          'da diese Anwendung als Add-In verwendet wird

      'Anwendungsname
      .ApplicationName = APPLICATION_NAME
      .ApplicationFullName = APPLICATION_FULLNAME
      .ApplicationTitle = APPLICATION_TITLE

      'Version
      .Version = APPLICATION_VERSION

      ' Formular, das am Ende von CurrentApplication.Start aufgerufen wird
      .ApplicationStartFormName = APPLICATION_STARTFORMNAME

   End With

End Sub


'############################################################################
'
' Funktionen für die Anwendungswartung
' (werden nur im Anwendungsentwurf benötigt)
'
'----------------------------------------------------------------------------
' Hilfsfunktion zum Speichern von Dateien in die lokale AppFile-Tabelle
'----------------------------------------------------------------------------
'Private Sub SetAppFiles()
'
'   Dim accFileName As Variant
'
'  ' Call CurrentApplication.Extensions("AppFile").SaveAppFile("AppIcon", CodeProject.Path & "\" & APPLICATION_ICONFILE)
'   With modApplication.CurrentApplication.Extensions("AppFile")
'      For Each accFileName In AccUnitLoaderConfigProcedures.AccUnitFileNames
'         .SaveAppFile accFileName, CodeProject.Path & "\lib\" & accFileName, True
'      Next
'   End With
'
'End Sub
