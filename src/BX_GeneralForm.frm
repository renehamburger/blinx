VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BX_GeneralForm 
   OleObjectBlob   =   "BX_GeneralForm.frx":0000
   Caption         =   "Blinx"
   ClientHeight    =   2304
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5508
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   15
End
Attribute VB_Name = "BX_GeneralForm"
Attribute VB_Base = "0{54B65651-7993-4D8E-9481-C36F317E6A68}{C951F528-C1DF-4646-AA80-0DEE798CF019}"
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
' See "ThisDocument" for license
'==============================================================================

Private m_nReturn As Integer
Private m_nButtons As Integer

Public Function MsgBox(ByVal sText As String, Optional ByVal nButtons As VbMsgBoxStyle = vbOKOnly, Optional sCaption As String = "Blinx") As Integer
  bx_sFunction = "GeneralForm_MsgBox"
  Select Case nButtons
   Case vbOKOnly, vbExclamation
      btn_OK.Caption = "OK"
      btn_OK.Visible = True
      btn_OK.Enabled = True
      btn_Cancel.Visible = False
      btn_Cancel.Enabled = False
   Case vbOKCancel
      btn_OK.Caption = "OK"
      btn_OK.Visible = True
      btn_OK.Enabled = True
      btn_Cancel.Caption = "Cancel"
      btn_Cancel.Visible = True
      btn_Cancel.Enabled = True
   Case vbYesNo
      btn_OK.Caption = "Yes"
      btn_OK.Visible = True
      btn_OK.Enabled = True
      btn_Cancel.Caption = "No"
      btn_Cancel.Visible = True
      btn_Cancel.Enabled = True
  End Select
  
  m_nButtons = nButtons
  SetText sText
  Me.Caption = sCaption
  tbx_Edit.Visible = False
  
  If (nButtons = vbExclamation) Then
    Me.BackColor = &HC0C0FF
    lbl_Text.BackColor = &HC0C0FF
  Else
    Me.BackColor = &HBBAA88
    lbl_Text.BackColor = &HBBAA88
  End If
  Me.Show
  
  MsgBox = m_nReturn
End Function

Public Function InputBox(ByVal sText As String, Optional sCaption As String = "Blinx", Optional sDefault As String = "") As String
  bx_sFunction = "GeneralForm_InputBox"
  m_nButtons = vbOKCancel
  SetText sText
  tbx_Edit.AutoSize = True
  tbx_Edit.MultiLine = False
  tbx_Edit.Text = sDefault
  tbx_Edit.AutoSize = False
  tbx_Edit.MultiLine = True
  tbx_Edit.Height = (Round(tbx_Edit.Width / 260 + 0.5)) * 15.75
  tbx_Edit.Width = 260
  tbx_Edit.Visible = True
  '<VBA_INSPECTOR>
  ' <DEPRECATION>
  '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
  '   <ITEM>[mso]Assistant.Left</ITEM>
  '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
  ' </DEPRECATION>
  '</VBA_INSPECTOR>
  tbx_Edit.Left = lbl_Text.Left
  '<VBA_INSPECTOR>
  ' <DEPRECATION>
  '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
  '   <ITEM>[mso]Assistant.Top</ITEM>
  '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
  ' </DEPRECATION>
  '</VBA_INSPECTOR>
  tbx_Edit.Top = lbl_Text.Top + lbl_Text.Height + 10
  btn_OK.Caption = "OK"
  btn_OK.Visible = True
  btn_OK.Enabled = True
  '<VBA_INSPECTOR>
  ' <DEPRECATION>
  '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
  '   <ITEM>[mso]Assistant.Top</ITEM>
  '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
  ' </DEPRECATION>
  '</VBA_INSPECTOR>
  btn_OK.Top = tbx_Edit.Top + tbx_Edit.Height + 10
  btn_Cancel.Caption = "Cancel"
  btn_Cancel.Visible = True
  btn_Cancel.Enabled = True
  '<VBA_INSPECTOR>
  ' <DEPRECATION>
  '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
  '   <ITEM>[mso]Assistant.Top</ITEM>
  '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
  ' </DEPRECATION>
  '</VBA_INSPECTOR>
  btn_Cancel.Top = tbx_Edit.Top + tbx_Edit.Height + 10
  lbl_Text.BackColor = &HBBAA88
  Me.BackColor = &HBBAA88
  '<VBA_INSPECTOR>
  ' <DEPRECATION>
  '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
  '   <ITEM>[mso]Assistant.Top</ITEM>
  '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
  ' </DEPRECATION>
  '</VBA_INSPECTOR>
  Me.Height = btn_OK.Top + btn_OK.Height + 28
  Me.Caption = sCaption
  tbx_Edit.SetFocus
  tbx_Edit.SelStart = 0
  tbx_Edit.SelLength = tbx_Edit.TextLength
  
  Me.Show
  
  If (m_nReturn = vbOK) Then
    InputBox = tbx_Edit.Text
  Else
    InputBox = ""
  End If
End Function




Public Sub ErrorBox(ByVal sText As String, Optional sCaption As String = "Blinx", Optional sDefault As String = "")
  bx_sFunction = "GeneralForm_ErrorBox"
  m_nButtons = vbExclamation
  SetText sText
  tbx_Edit.AutoSize = True
  tbx_Edit.MultiLine = False
  tbx_Edit.Text = sDefault
  tbx_Edit.AutoSize = False
  tbx_Edit.MultiLine = True
  tbx_Edit.Height = (Round(tbx_Edit.Width / 260 + 0.5)) * 15.75
  tbx_Edit.Width = 260
  tbx_Edit.Visible = True
  '<VBA_INSPECTOR>
  ' <DEPRECATION>
  '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
  '   <ITEM>[mso]Assistant.Left</ITEM>
  '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
  ' </DEPRECATION>
  '</VBA_INSPECTOR>
  tbx_Edit.Left = lbl_Text.Left
  '<VBA_INSPECTOR>
  ' <DEPRECATION>
  '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
  '   <ITEM>[mso]Assistant.Top</ITEM>
  '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
  ' </DEPRECATION>
  '</VBA_INSPECTOR>
  tbx_Edit.Top = lbl_Text.Top + lbl_Text.Height + 10
  btn_OK.Caption = "OK"
  btn_OK.Visible = True
  btn_OK.Enabled = True
  '<VBA_INSPECTOR>
  ' <DEPRECATION>
  '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
  '   <ITEM>[mso]Assistant.Top</ITEM>
  '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
  ' </DEPRECATION>
  '</VBA_INSPECTOR>
  btn_OK.Top = tbx_Edit.Top + tbx_Edit.Height + 10
  '<VBA_INSPECTOR>
  ' <DEPRECATION>
  '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
  '   <ITEM>[mso]Assistant.Left</ITEM>
  '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
  ' </DEPRECATION>
  '</VBA_INSPECTOR>
  btn_OK.Left = Me.Width - btn_OK.Width - 10
  btn_Cancel.Visible = False
  btn_Cancel.Enabled = True
  '<VBA_INSPECTOR>
  ' <DEPRECATION>
  '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
  '   <ITEM>[mso]Assistant.Top</ITEM>
  '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
  ' </DEPRECATION>
  '</VBA_INSPECTOR>
  Me.Height = btn_OK.Top + btn_OK.Height + 28
  Me.Caption = sCaption
  Me.BackColor = &HC0C0FF
  lbl_Text.BackColor = &HC0C0FF
  tbx_Edit.SetFocus
  tbx_Edit.SelStart = 0
  tbx_Edit.SelLength = tbx_Edit.TextLength
  
  Me.Show
End Sub


Public Sub SetText(ByVal sText As String)
  bx_sFunction = "GeneralForm_SetText"
  lbl_Text.Width = 260
  lbl_Text.Caption = sText
  lbl_Text.Width = 260
  '<VBA_INSPECTOR>
  ' <DEPRECATION>
  '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
  '   <ITEM>[mso]Assistant.Top</ITEM>
  '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
  ' </DEPRECATION>
  '</VBA_INSPECTOR>
  btn_Cancel.Top = lbl_Text.Top + lbl_Text.Height + 10
  '<VBA_INSPECTOR>
  ' <DEPRECATION>
  '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
  '   <ITEM>[mso]Assistant.Top</ITEM>
  '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
  ' </DEPRECATION>
  '</VBA_INSPECTOR>
  btn_OK.Top = lbl_Text.Top + lbl_Text.Height + 10
  '<VBA_INSPECTOR>
  ' <DEPRECATION>
  '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
  '   <ITEM>[mso]Assistant.Top</ITEM>
  '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
  ' </DEPRECATION>
  '</VBA_INSPECTOR>
  Me.Height = btn_OK.Top + btn_OK.Height + 28
  If (btn_Cancel.Enabled) Then
    '<VBA_INSPECTOR>
    ' <DEPRECATION>
    '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
    '   <ITEM>[mso]Assistant.Left</ITEM>
    '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
    ' </DEPRECATION>
    '</VBA_INSPECTOR>
    btn_Cancel.Left = Me.Width - btn_Cancel.Width - 10
    '<VBA_INSPECTOR>
    ' <DEPRECATION>
    '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
    '   <ITEM>[mso]Assistant.Left</ITEM>
    '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
    ' </DEPRECATION>
    '</VBA_INSPECTOR>
    btn_OK.Left = btn_Cancel.Left - btn_OK.Width - 10
  Else
    '<VBA_INSPECTOR>
    ' <DEPRECATION>
    '   <MESSAGE>Potentially contains deprecated items in the object model</MESSAGE>
    '   <ITEM>[mso]Assistant.Left</ITEM>
    '   <URL>http://go.microsoft.com/fwlink/?LinkID=215358 /URL>
    ' </DEPRECATION>
    '</VBA_INSPECTOR>
    btn_OK.Left = Width - btn_OK.Width - 10
  End If
End Sub

Private Sub btn_Cancel_Click()
  bx_sFunction = "GeneralForm_Cancel_Click"
  Select Case m_nButtons
    Case vbOKCancel
      m_nReturn = vbCancel
    Case vbYesNo
      m_nReturn = vbNo
  End Select
  Me.hide
End Sub


Private Sub btn_OK_Click()
  bx_sFunction = "GeneralForm_OK_Click"
  Select Case m_nButtons
    Case vbOKOnly Or vbExclamation
      m_nReturn = vbOK
    Case vbOKCancel
      m_nReturn = vbOK
    Case vbYesNo
      m_nReturn = vbYes
  End Select
  Me.hide
End Sub

Private Sub UserForm_Activate()
  bx_sFunction = "GeneralForm_Activate"
  bOK = False
  bCancel = False
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  bx_sFunction = "GeneralForm_QueryClose"
  If (CloseMode = vbFormControlMenu) Then
    Cancel = 1
    btn_Cancel_Click
  End If
End Sub
