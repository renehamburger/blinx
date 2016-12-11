VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEvents"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
'==============================================================================
' Methods shared by BX module & classes
' See https://github.com/renehamburger/blinx for source code, manual & license
'==============================================================================

Public WithEvents appWord As Word.Application
Attribute appWord.VB_VarHelpID = -1

'==============================================================================
'   Events
'==============================================================================

Private Sub appWord_WindowBeforeRightClick(ByVal oSel As Selection, bCancel As Boolean)
  bx_sFunction = "clsWindows_WindowBeforeRightClick"
  Dim oReference As BX_Reference
  Dim sVersion As String
  
  bCancel = False
  If (oSel.Hyperlinks.Count() = 1) Then
    If (oSel.Hyperlinks(1).Range.Start <= oSel.Start + 1 And oSel.Hyperlinks(1).Range.End >= oSel.End - 2) Then
      If (BX_GetDataFromLink(oSel, oReference, sVersion)) Then
        If (bx_oB Is Nothing) Then Set bx_oB = New clsBibleBW
        If (bx_oB.IsApplicationOK() And TypeName(bx_oB) = "clsBibleBW") Then
          bx_oB.GotoPassage BX_ReferenceToString(oReference), sVersion
          oSel.End = oSel.Start
          BX_FocusBW
          bCancel = True
        End If
      End If
    End If
  End If
End Sub
