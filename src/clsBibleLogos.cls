VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBibleLogos"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Option Base 1
'==============================================================================
' Logos implementation of IBible interface
' See https://github.com/renehamburger/blinx for source code, manual & license
'==============================================================================
Implements IBible

Private m_oApp As LogosApplication

'==============================================================================
'   Initialisation/Termination
'==============================================================================

Private Sub Class_Initialize()
  On Error GoTo FAILED
  bx_sFunction = "clsBibleLogos_Initialize"
  Dim Launcher As LogosLauncher

  Set Launcher = New LogosLauncher
  Launcher.LaunchApplication
  Set m_oApp = Launcher.Application
  Exit Sub

FAILED:
  Set m_oApp = Nothing
End Sub

Private Sub Class_Terminate()
  bx_sFunction = "clsBibleLogos_Terminate"
  Set m_oApp = Nothing
End Sub

'==============================================================================
'   Public Functions
'==============================================================================

Public Function IBible_IsApplicationOK() As Boolean
  bx_sFunction = "clsBibleLogos_IsApplicationOK"
  IBible_IsApplicationOK = Not m_oApp Is Nothing
End Function

Public Function IBible_IsReady() As Boolean
  bx_sFunction = "clsBibleLogos_IsReady"
  IBible_IsReady = IBible_IsApplicationOK()
End Function

Public Function IBible_BeginProcess(ByVal sVersion As String) As Boolean
  bx_sFunction = "clsBibleLogos_BeginProcess"
  IBible_BeginProcess = IBible_IsReady()
End Function

Public Sub IBible_EndProcess()
  bx_sFunction = "clsBibleLogos_EndProcess"
End Sub

Public Function IBible_CheckPassage(ByVal sRef As String, ByVal sVersionIn As String, Optional ByVal bFirstVerseOnly As Boolean = False) As Boolean
  bx_sFunction = "clsBibleLogos_CheckPassage"
  IBible_CheckPassage = IBible_GetPassage(sRef, sVersionIn, bFirstVerseOnly) <> ""
End Function

Public Function IBible_GetPassage(ByRef sRef As String, ByVal sVersionIn As String, Optional ByVal bSuppressError As Boolean = False) As String
  bx_sFunction = "clsBibleLogos_GetPassage"
  Dim oRequest As LogosCopyBibleVersesRequest
  Dim oReference As LogosDataTypeReference

  Set oRequest = m_oApp.CopyBibleVerses.CreateRequest()
  Set oReference = m_oApp.DataTypes.GetDataType("bible").ParseReference(sRef)
  oRequest.Reference = oReference
  IBible_GetPassage = m_oApp.CopyBibleVerses.GetText(oRequest)
End Function

Public Function IBible_GotoPassage(ByVal sRef As String, ByVal sVersionIn As String) As Boolean
  bx_sFunction = "clsBibleLogos_GotoVerse"
  ' TODO
End Function
