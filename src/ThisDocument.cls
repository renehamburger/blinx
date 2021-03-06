VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_Base = "0{00020906-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Option Base 1
'==============================================================================
' See https://github.com/renehamburger/blinx for source code, manual & license
'==============================================================================

Dim m_oEvents As New clsEvents

Public Sub AutoExec()
  Set m_oEvents.appWord = Word.Application
  BX_Reset
End Sub

Public Sub AutoExit()
  Set bx_oB = Nothing
  Set bx_oClip = Nothing
End Sub

Public Sub BX_ShowOptionsForm()
  BX_CheckForms
  'If (bx_oOptionsForm Is Nothing) Then Set bx_oOptionsForm = New BX_OptionsForm
  bx_oOptionsForm.Show
End Sub

Private Sub BX_ChangeToolTip()
  CommandBars("Blinx").Controls("Create Blink").TooltipText = "Create Blink (Alt-B)"
  CommandBars("Blinx").Controls("Create Blink with Tooltip").TooltipText = "Create Blink && tooltip (Alt-B)"
  CommandBars("Blinx").Controls("Create Blink with Text").TooltipText = "Create Blink && insert text (Alt-Shift-B)"
  CommandBars("Blinx").Controls("Unlink All Links").TooltipText = "Unlink Blinks and hyperlinks (Alt-U)"
  CommandBars("Blinx").Controls("Show Options").TooltipText = "Show Options (Ctrl-Alt-Shift-B)"
End Sub
