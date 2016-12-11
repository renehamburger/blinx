VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BX_ProcessInputForm 
   OleObjectBlob   =   "BX_ProcessInputForm.frx":0000
   Caption         =   "Blinx"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4935
   TypeInfoVer     =   13
End
Attribute VB_Name = "BX_ProcessInputForm"
Attribute VB_Base = "0{BB3D309B-F401-4B19-BDFB-4CB25D430640}{650EDC6F-E4DF-4D30-9592-17667614B57B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'<VBA_INSPECTOR_RUN />
Option Explicit
Option Base 1
'==============================================================================
' See https://github.com/renehamburger/blinx for source code, manual & license
'==============================================================================

Public sRef As String
Public bProcess As Boolean
Public bAcceptAll As Boolean
Public bAbort As Boolean
Private bFirstRun As Boolean
Public oWindow As Window

'==============================================================================
'   Public Functions
'==============================================================================

Public Function SetWindow(ByVal oWindowIn As Window)
  Set oWindow = oWindowIn
  oWindow.Activate
End Function

'==============================================================================
'   Private Functions
'==============================================================================

Private Sub UserForm_Initialize()
  sRef = ""
  bFirstRun = True
End Sub

Private Sub UserForm_Activate()
  If (bFirstRun) Then
    bFirstRun = False
    Top = 4
    '<VBA_INSPECTOR>
    ' <DEPRECATION>
    '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
    '   <ITEM>[mso]Assistant.Left</ITEM>
    '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
    ' </DEPRECATION>
    '</VBA_INSPECTOR>
    Left = ActiveWindow.Left + (ActiveWindow.Width / 2) - Width / 2
    If (Left < 0) Then Left = 10
  End If
  
  tbx_Input.Text = sRef
  tbx_Input.SelStart = 0
  tbx_Input.SelLength = tbx_Input.TextLength
  tbx_Input.SetFocus
  bAcceptAll = False
  bAbort = False
  If (bProcess) Then
    btn_AcceptAll.Enabled = True
    btn_Abort.Enabled = True
  Else
    btn_AcceptAll.Enabled = False
    btn_Abort.Enabled = False
  End If
End Sub

Private Sub btn_Accept_Click()
  sRef = tbx_Input.Text
  UserInput
End Sub

Private Sub btn_AcceptAll_Click()
  sRef = tbx_Input.Text
  bAcceptAll = True
  UserInput
End Sub

Private Sub btn_Skip_Click()
  sRef = ""
  UserInput
End Sub


Private Sub btn_Abort_Click()
  bAbort = True
  UserInput
End Sub

Private Sub UserInput()
  Me.hide
  If (Not oWindow Is Nothing And IsObject(oWindow)) Then oWindow.Activate
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If (CloseMode = vbFormControlMenu) Then
    Cancel = 1
    If (bProcess) Then
      btn_Abort_Click
    Else
      btn_Skip_Click
    End If
  End If
  On Error Resume Next 'Next line often causes an "object has been deleted error", but can't see how to avoid that
  If (Not oWindow Is Nothing And IsObject(oWindow)) Then oWindow.Activate
End Sub
