VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BX_EditBookNamesForm 
   OleObjectBlob   =   "BX_EditBookNamesForm.frx":0000
   Caption         =   "Edit Book Names"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4575
   StartUpPosition =   1  'Fenstermitte
   TypeInfoVer     =   14
End
Attribute VB_Name = "BX_EditBookNamesForm"
Attribute VB_Base = "0{B10DA30A-7C6C-45F7-BB1D-5D8DF9151E3D}{84F3D2FB-A83A-42E5-99F7-1263DE6604C9}"
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

Public m_nItem As Integer
Private m_aoBoxes(10) As TextBox

Private Sub btn_Cancel_Click()
  Me.hide
End Sub

Private Sub UserForm_Activate()
  Dim nI As Integer
  
  Set m_aoBoxes(1) = TextBox1
  Set m_aoBoxes(2) = TextBox2
  Set m_aoBoxes(3) = TextBox3
  Set m_aoBoxes(4) = TextBox4
  Set m_aoBoxes(5) = TextBox5
  Set m_aoBoxes(6) = TextBox6
  Set m_aoBoxes(7) = TextBox7
  Set m_aoBoxes(8) = TextBox8
  Set m_aoBoxes(9) = TextBox9
  Set m_aoBoxes(10) = TextBox10
  
  For nI = 1 To 10
    m_aoBoxes(nI).Text = bx_asBooks(bx_eLanguage, m_nItem + 1, nI + 1)
  Next
End Sub

Private Sub btn_OK_Click()
  Dim nI As Integer
  Dim nSize As Integer
  nSize = 0
  For nI = 1 To 10
   bx_asBooks(bx_eLanguage, m_nItem + 1, nI + 1) = m_aoBoxes(nI).Text
   If (m_aoBoxes(nI).Text = "" And nSize = 0) Then nSize = nI - 1
  Next
  bx_asBooks(bx_eLanguage, m_nItem + 1, 1) = nSize
  BX_SaveVariables
  BX_LoadVariables
  bx_oOptionsForm.ReloadBookList
  Me.hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  bx_sFunction = "EditBookNames_QueryClose"
  If (CloseMode = vbFormControlMenu) Then
    Cancel = 1
    btn_Cancel_Click
  End If
End Sub

Private Sub TextBox2_Change()
  If (TextBox2.Text = "" And TextBox3.Text <> "") Then
    TextBox2.Text = TextBox3.Text
    TextBox3.Text = ""
  End If
End Sub

Private Sub TextBox3_Change()
  Dim sText As String
  If (TextBox2.Text = "" And TextBox3.Text <> "") Then
    TextBox2.SetFocus
    sText = TextBox3.Text
    TextBox3.Text = ""
    TextBox2.Text = sText
  End If
  If (TextBox3.Text = "" And TextBox4.Text <> "") Then
    TextBox3.Text = TextBox4.Text
    TextBox4.Text = ""
  End If
End Sub

Private Sub TextBox4_Change()
  Dim sText As String
  If (TextBox3.Text = "" And TextBox4.Text <> "") Then
    TextBox3.SetFocus
    sText = TextBox4.Text
    TextBox4.Text = ""
    TextBox3.Text = sText
  End If
  If (TextBox4.Text = "" And TextBox5.Text <> "") Then
    TextBox4.Text = TextBox5.Text
    TextBox5.Text = ""
  End If
End Sub

Private Sub TextBox5_Change()
  Dim sText As String
  If (TextBox4.Text = "" And TextBox5.Text <> "") Then
    TextBox4.SetFocus
    sText = TextBox5.Text
    TextBox5.Text = ""
    TextBox4.Text = sText
  End If
  If (TextBox5.Text = "" And TextBox6.Text <> "") Then
    TextBox5.Text = TextBox6.Text
    TextBox6.Text = ""
  End If
End Sub

Private Sub TextBox6_Change()
  Dim sText As String
  If (TextBox5.Text = "" And TextBox6.Text <> "") Then
    TextBox5.SetFocus
    sText = TextBox6.Text
    TextBox6.Text = ""
    TextBox5.Text = sText
  End If
  If (TextBox6.Text = "" And TextBox7.Text <> "") Then
    TextBox6.Text = TextBox7.Text
    TextBox7.Text = ""
  End If
End Sub

Private Sub TextBox7_Change()
  Dim sText As String
  If (TextBox6.Text = "" And TextBox7.Text <> "") Then
    TextBox6.SetFocus
    sText = TextBox7.Text
    TextBox7.Text = ""
    TextBox6.Text = sText
  End If
  If (TextBox7.Text = "" And TextBox8.Text <> "") Then
    TextBox7.Text = TextBox8.Text
    TextBox8.Text = ""
  End If
End Sub

Private Sub TextBox8_Change()
  Dim sText As String
  If (TextBox7.Text = "" And TextBox8.Text <> "") Then
    TextBox7.SetFocus
    sText = TextBox8.Text
    TextBox8.Text = ""
    TextBox7.Text = sText
  End If
  If (TextBox8.Text = "" And TextBox9.Text <> "") Then
    TextBox8.Text = TextBox9.Text
    TextBox9.Text = ""
  End If
End Sub

Private Sub TextBox9_Change()
  Dim sText As String
  If (TextBox8.Text = "" And TextBox9.Text <> "") Then
    TextBox8.SetFocus
    sText = TextBox9.Text
    TextBox9.Text = ""
    TextBox8.Text = sText
  End If
  If (TextBox9.Text = "" And TextBox10.Text <> "") Then
    TextBox9.Text = TextBox10.Text
    TextBox10.Text = ""
  End If
End Sub

Private Sub TextBox10_Change()
  Dim sText As String
  If (TextBox9.Text = "" And TextBox10.Text <> "") Then
    TextBox9.SetFocus
    sText = TextBox10.Text
    TextBox10.Text = ""
    TextBox9.Text = sText
  End If
End Sub
