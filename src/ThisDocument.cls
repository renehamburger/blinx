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
' ---------------------------------------------------------------------------------------
' Blinx Add-In, v0.9 (February 2012)
' Copyright 2010-12
' Rene Hamburger (rene.hamburger@gmail.com)
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, version 3 or above (see www.gnu.org/licenses),
' with the additional restriction of non-commercial use only.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
' ---------------------------------------------------------------------------------------

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
  CommandBars("Blinx").Controls("Create Blink").TooltipText = "Create Blink (AltGr-B)"
  CommandBars("Blinx").Controls("Create Blink with Tooltip").TooltipText = "Create Blink && tooltip (Alt-B)"
  CommandBars("Blinx").Controls("Create Blink with Text").TooltipText = "Create Blink && insert text (Alt-Shift-B)"
  CommandBars("Blinx").Controls("Unlink All Links").TooltipText = "Unlink Blinks and hyperlinks (Alt-U)"
  CommandBars("Blinx").Controls("Show Options").TooltipText = "Show Options (Ctrl-Alt-Shift-B)"
End Sub
