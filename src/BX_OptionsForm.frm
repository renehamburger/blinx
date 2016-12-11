VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BX_OptionsForm 
   OleObjectBlob   =   "BX_OptionsForm.frx":0000
   Caption         =   "Blinx Options"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6555
   StartUpPosition =   2  'CenterScreen
   TypeInfoVer     =   149
End
Attribute VB_Name = "BX_OptionsForm"
Attribute VB_Base = "0{FF73C92F-0DBB-4E85-81A8-7329DFA2C45E}{0C69B66C-E3B8-4095-9E55-70B5C957169A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Option Base 1
'==============================================================================
' See https://github.com/renehamburger/blinx for source code, manual & license
'==============================================================================

Public Sub ReloadBookList()
  Dim nI As Integer
  Dim nJ As Integer
  Dim nSize As Integer
  
  BX_LoadVariables
  For nI = 1 To 66
    nSize = CInt(bx_asBooks(nI, 1))
    lbx_books.AddItem
    For nJ = 1 To 10
      lbx_books.List(nI - 1, nJ - 1) = bx_asBooks(nI, nJ + 1)
    Next
  Next
End Sub

Private Sub lbx_books_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  bx_oEditBookNamesForm.m_nItem = lbx_books.ListIndex
  bx_oEditBookNamesForm.Show
End Sub

Private Sub tbx_about_Change()

End Sub

Private Sub UserForm_Initialize()
  bx_sFunction = "OptionsForm_Initialize"
  
  AddRows cbx_Translation, BX_TRANSLATION
  AddRows cbx_OnlineBible, BX_ONLINE_BIBLE
  AddRows cbx_BlinkPreviewLength, BX_BLINK_PREVIEW_LENGTH
  
  tbx_about.Text = "Blinx " & BX_VERSION & " Add-In for Microsoft Word, ©2010-11, Rene Hamburger." & vbCrLf & vbCrLf & _
                   "This program is free software. Please distribute it together with 'Blinx-Readme.rtf'. Please email bugs/suggestions to: blinx.add.in@gmail.com." & vbCrLf & vbCrLf & _
                   "I have dedicated this work to our Lord and Saviour! And to the fantastic theological college I have the privilege to train at: www.oakhill.ac.uk. My prayer is that this tool might be as useful to many other Christians and students of God's Word as it has been to me." & vbCrLf & vbCrLf & _
                   "Full copyright notice:" & vbCrLf & _
                   "This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, version 3 or above (see www.gnu.org/licenses), with the additional restriction of non-commercial use only." & vbCrLf & _
                   "This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details." & vbCrLf & vbCrLf & _
                   "Version: " & BX_VERSION_FULL
  tbx_about.SelStart = 0
  tbx_about.SelLength = 0
End Sub

Private Sub UserForm_Activate()
  bx_sFunction = "OptionsForm_Activate"
  SelectItem cbx_Translation, GetSetting("Blinx", "Options", "Translation", Split(BX_TRANSLATION, "#")(0))
  SelectItem cbx_OnlineBible, GetSetting("Blinx", "Options", "OnlineBible", Split(BX_ONLINE_BIBLE, "#")(0))
  SelectItem cbx_BlinkPreviewLength, GetSetting("Blinx", "Options", "BlinkPreviewLength", Split(BX_BLINK_PREVIEW_LENGTH, "#")(0))
  ReloadBookList
End Sub

Private Sub btn_Cancel_Click()
  bx_sFunction = "btn_Cancel_Click"
  Me.hide
End Sub

Private Sub btn_OK_Click()
  bx_sFunction = "btn_OK_Click"
  SaveSetting "Blinx", "Options", "Translation", cbx_Translation.Value
  SaveSetting "Blinx", "Options", "OnlineBible", cbx_OnlineBible.Value
  SaveSetting "Blinx", "Options", "BlinkPreviewLength", cbx_BlinkPreviewLength.Value
  Me.hide
End Sub

Private Sub btn_Reset_Click()
  bx_sFunction = "btn_Reset_Click"
  ThisDocument.AutoExec
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  bx_sFunction = "OptionsForm_QueryClose"
  If (CloseMode = vbFormControlMenu) Then
    Cancel = 1
    btn_Cancel_Click
  End If
End Sub

Private Sub AddRows(ByVal oCombo As ComboBox, ByVal sRow As String)
  bx_sFunction = "OptionsForm_AddRows"
  Dim nI As Integer
  Dim nJ As Integer
  Dim sString As String
  
  For nI = 1 To BX_CountInStr(sRow, "#")
    sString = Split(sRow, "#")(nI)
    oCombo.AddItem
    For nJ = 0 To BX_CountInStr(sString, "|")
      oCombo.List(nI - 1, nJ) = Split(sString, "|")(nJ)
    Next
  Next
End Sub

Private Sub SelectItem(ByVal oCombo As ComboBox, ByVal vValue As Variant)
  bx_sFunction = "OptionsForm_SelectItem"
  Dim nI As Integer
  For nI = 0 To oCombo.ListCount - 1
    If (oCombo.List(nI) = vValue) Then
      oCombo.ListIndex = nI
      Exit Sub
    End If
  Next
End Sub
