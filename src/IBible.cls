VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBible"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
'==============================================================================
' Interface IBible
' See https://github.com/renehamburger/blinx for source code, manual & license
'==============================================================================

Public Function IsApplicationOK() As Boolean
' Determines whether connection to application object (BibleWorks, Internet Explorer) has been successful
End Function

Public Function IsReady() As Boolean
' Determines whether requests can be started
End Function

Public Function BeginProcess(ByVal sVersion As String) As Boolean
' Make preparations before a batch of passage requests
End Function

Public Sub EndProcess()
' Clean up after a batch of passage requests
End Sub

Public Function CheckPassage(ByVal sRef As String, ByVal sVersion As String, Optional ByVal bFirstVerseOnly As Boolean = False) As Boolean
' Check whether passage can be found
End Function

Public Function GetPassage(ByRef sRef As String, ByVal sVersion As String, Optional ByVal bSuppressError As Boolean = False) As String
' Get text of passage
End Function

Public Function GotoPassage(ByVal sRef As String, ByVal sVersion As String) As Boolean
' Goto passage
End Function
