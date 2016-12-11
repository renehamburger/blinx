VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBibleBW"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'<VBA_INSPECTOR_RUN />
Option Explicit
'==============================================================================
' BibleWorks implementation of IBible interface
' (See "ThisDocument" for license)
'==============================================================================
Implements IBible

Private m_oBW As Object
Private m_sOldDisplayVersions As String
Private m_sOldRef As String
Private m_bToggleBrowseMode As Boolean
Private m_bToggleNotes As Boolean
Private m_bToggleStrongsNumbers As Boolean
Private m_bClipboardSaved As Boolean
Private m_sLastRefIn As String
Private m_sLastRefOut As String
Private m_sLastVersion As String
Private m_sLastText As String

'==============================================================================
'   Initialisation/Termination
'==============================================================================

Private Sub Class_Initialize()
  bx_sFunction = "clsBibleBW_Initialize"
  If (m_oBW Is Nothing) Then Set m_oBW = BX_GetBWObject1()
  If (m_oBW Is Nothing) Then Set m_oBW = BX_GetBWObject2()
  If (Not m_oBW Is Nothing) Then
    If (Not TypeName(m_oBW) = "IAutomation") Then
      Set m_oBW = Nothing
    End If
  End If
End Sub

Private Sub Class_Terminate()
  bx_sFunction = "clsBibleBW_Terminate"
  Set m_oBW = Nothing
End Sub

'==============================================================================
'   Public Functions
'==============================================================================

Public Function IBible_IsApplicationOK() As Boolean
  bx_sFunction = "clsBibleBW_IsApplicationOK"
  If (Not m_oBW Is Nothing) Then
    If (TypeName(m_oBW) <> "Object") Then
      IBible_IsApplicationOK = True
    Else
      IBible_IsApplicationOK = False
    End If
  Else
    IBible_IsApplicationOK = False
  End If
End Function

Public Function IBible_IsReady() As Boolean
  bx_sFunction = "clsBibleBW_IsReady"
  IBible_IsReady = IBible_IsApplicationOK()
End Function

Public Function IBible_BeginProcess(ByVal sVersionIn As String) As Boolean
  bx_sFunction = "clsBibleBW_BeginProcess"
  Dim sText As String
  Dim bOK As Boolean
  Dim sVersion As String
  
  sVersion = CheckVersion(sVersionIn)
  If (Not IBible_IsApplicationOK()) Then
    IBible_BeginProcess = False
  Else
  '---Prep
    m_bToggleBrowseMode = False
    m_bToggleNotes = False
    m_bToggleStrongsNumbers = False
    m_sOldDisplayVersions = m_oBW.GetVersions()
    m_sOldRef = m_oBW.GetCurrentVerse()
    m_sOldRef = Mid(m_sOldRef, InStr(m_sOldRef, " ") + 1, 99)
    'm_bClipboardSaved = bx_oClip.Save()
    m_oBW.ClipGoToVerse True
    m_oBW.IgnoreErrors True
    
  '---Check browse mode & translation notes
    m_oBW.SetVersions "NIV NIV"
    m_oBW.GoToVerse "3John 1:1"
    sText = bx_oClip.GetText()
    If (Len(sText) > 100) Then
      BX_ToggleBrowseMode
      m_bToggleBrowseMode = True
    End If
    If (InStr(1, sText, "(a )", vbBinaryCompare) > 0) Then
      BX_ToggleNotes
      m_bToggleNotes = True
    End If
    If (sVersion = "KJV") Then
      m_oBW.SetVersions "KJV KJV"
      m_oBW.GoToVerse "3John 1:1"
      sText = bx_oClip.GetText()
      If (InStr(1, sText, "<4245>", vbBinaryCompare) > 0) Then
        BX_ToggleStrongsNumbers
        m_bToggleStrongsNumbers = True
      End If
    End If
    IBible_BeginProcess = True
  End If
End Function

Public Sub IBible_EndProcess()
  bx_sFunction = "clsBibleBW_EndProcess"
  m_oBW.ClipGoToVerse False
  m_oBW.SetVersions m_sOldDisplayVersions
  m_oBW.GoToVerse m_sOldRef
  If (m_bToggleBrowseMode) Then BX_ToggleBrowseMode
  If (m_bToggleNotes) Then BX_ToggleNotes
  If (m_bToggleStrongsNumbers) Then BX_ToggleStrongsNumbers
'  If m_bClipboardSaved Then
'    bx_oClip.Restore
'  Else
    bx_oClip.EmptyIt
'  End If
End Sub

Public Function IBible_CheckPassage(ByVal sRef As String, ByVal sVersionIn As String, Optional ByVal bFirstVerseOnly As Boolean = False) As Boolean
  Dim oRef As BX_Reference
  Dim sVersion As String
  
  sVersion = CheckVersion(sVersionIn)
  If (bFirstVerseOnly) Then
    oRef = BX_StringtToReference(sRef)
    If (oRef.nVerse1 = 0) Then oRef.nVerse1 = 1
    IBible_CheckPassage = InternalCheckPassage(oRef.sBook & " " & oRef.nChapter1 & ":" & oRef.nVerse1, sVersion)
  Else
    IBible_CheckPassage = IBible_GetPassage(sRef, sVersion, True) <> ""
  End If
End Function


Public Function IBible_GetPassage(ByRef sRef As String, ByVal sVersionIn As String, Optional ByVal bSuppressError As Boolean = False) As String
  bx_sFunction = "clsBibleBW_GetPassage"
  Dim sNewRef As String
  Dim sData As String
  Dim sTest As String
  Dim nPos1 As Long
  Dim nPos2 As Long
  Dim nPos3 As Long
  Dim nPos4 As Long
  Dim nPos5 As Long
  Dim bOK As Boolean
  Dim bRefHas3 As Boolean
  Dim sVersion As String
  
  sVersion = CheckVersion(sVersionIn)
  If (sRef <> "") Then
    If (sRef = m_sLastRefIn And sVersion = m_sLastVersion) Then
      sData = m_sLastText
      sRef = m_sLastRefOut
    Else
      bOK = True
      
    '---Go to verse and check whether it has worked (if not, BW will remain in 3Jn (or Obad))
      bRefHas3 = False
      If (Left(sRef, 1) = "3") Then bRefHas3 = True
      If (Len(sRef) > 4) Then
        If (LCase(Left(sRef, 4)) = "iii ") Then bRefHas3 = True
      End If
      
      m_oBW.ClipGoToVerse False
      If (bRefHas3) Then
        m_oBW.GoToVerse "Obadiah 1:1"
      Else
        m_oBW.GoToVerse "3John 1:1"
      End If
      
      m_oBW.SetVersions sVersion & " " & sVersion
      m_oBW.ClipGoToVerse True
      m_oBW.GoToVerse sRef
      sData = bx_oClip.GetText()
      If (Len(sData) > 6) Then
      
        If (Left(sData, 1) = " ") Then
          bx_sFunction = "clsBibleBW_GetPassage--IncorrectBWRequest"
        End If
            
        nPos1 = InStr(sData, "  ")
        If (bRefHas3) Then
          If (Mid(sData, nPos1 + 2, 1) = "O") Then bOK = False
        Else
          If (Mid(sData, nPos1 + 2, 1) = "3") Then bOK = False
        End If
      Else
        bOK = False
      End If
      
      If (bOK) Then
        m_sLastRefIn = sRef
        m_sLastVersion = sVersion
      
      '---Emulate BW 7 format:
        'sData = Left(sData, 4) & Mid(sData, 6)
      '---Prepare reference
        nPos1 = InStr(sData, " ")
        nPos3 = BX_InStr(1, sData, ":")
        
''---debug:
'If (nPos3 = 0 Or nPos3 > Len(sRef) + 6) Then
' If (GetSetting("Blinx", "Options", "Debugger", "") = "1") Then
'  Stop
' End If
'End If

        nPos2 = InStrRev(sData, " ", nPos3)
      '---Get reference
        sNewRef = Trim(Mid(sData, nPos1, nPos2 - nPos1)) & " "
        nPos4 = BX_InStr(1, sRef, ":")
        If (nPos4 = 0) Then
          nPos5 = InStrRev(sRef, " ")
        Else
          nPos5 = InStrRev(sRef, " ", nPos4)
        End If
        sNewRef = sNewRef & Trim(Mid(sRef, nPos5, Len(sRef) - nPos5 + 1))
        sRef = sNewRef
      '---Prepare passage
        sData = Mid(sData, nPos2 + 1)
        ConvertPassage sRef, sData
        m_sLastRefOut = sRef
        m_sLastText = sData
      Else
        If Not (bSuppressError) Then bx_oGeneralForm.MsgBox "BibleWorks was not able to determine this Bible reference", vbOKOnly
        sData = ""
      End If
    End If
  End If
  
  IBible_GetPassage = sData
End Function

Public Function IBible_GotoPassage(ByVal sRef As String, ByVal sVersionIn As String) As Boolean
  bx_sFunction = "clsBibleBW_GotoVerse"
  Dim oRef As BX_Reference
  Dim sText As String
  Dim sVersion As String
  
  sVersion = CheckVersion(sVersionIn)
  If (Not m_oBW Is Nothing) Then
    oRef = BX_StringtToReference(sRef)
    If (oRef.nChapter2 > oRef.nChapter1) Then
      m_oBW.SetVersions "NIV NIV"
      m_oBW.ClipGoToVerse True
      m_oBW.GoToVerse "3John 1:1"
      sText = bx_oClip.GetText()
      If (Len(sText) < 100) Then
        BX_ToggleBrowseMode
      End If
      If (oRef.nVerse1 = 0) Then oRef.nVerse1 = 1
      m_oBW.ClipGoToVerse False
      m_oBW.GoToVerse oRef.sBook & " " & oRef.nChapter1 & ":" & oRef.nVerse1
    Else
      m_oBW.ClipGoToVerse False
      m_oBW.GoToVerse sRef
    End If
    IBible_GotoPassage = True
  Else
    IBible_GotoPassage = False
  End If
End Function

'==============================================================================
'   Private Functions
'==============================================================================

Private Function BX_GetBWObject1() As Object
  On Error GoTo FAILED
  Set BX_GetBWObject1 = CreateObject("bibleworks.automation")
  Exit Function
FAILED:
  Set BX_GetBWObject1 = Nothing
End Function

Private Function BX_GetBWObject2() As Object
  On Error GoTo FAILED
  Set BX_GetBWObject2 = CreateObject("bwcom.document")
  Exit Function
FAILED:
  Set BX_GetBWObject2 = Nothing
End Function

Private Function CheckVersion(ByVal sVersion As String) As String
  If (sVersion = "NASB") Then
    CheckVersion = "NAS"
  Else
    CheckVersion = sVersion
  End If
End Function

Private Sub ConvertPassage(ByVal sRefString As String, ByRef sPass As String)
  bx_sFunction = "clsBibleBW_ConvertPassage"
  Dim oRef As BX_Reference
  Dim anPos() As Long
  Dim nChap As Integer
  Dim nVerse As Integer
  Dim nP As Long
  Dim nP2 As Long
  Dim bOK As Boolean
  Dim nLen As Long
  Dim sChar As String
  Dim sPass2 As String
  
'---Mark chapter numbers with "@@" and verse numbers with "##"
 '--Determine chapters & verses
  oRef = BX_StringtToReference(sRefString, True, ":")
  If (oRef.nChapter1 <= oRef.nChapter2) Then
    nLen = Len(sPass)
   '--Determine positions of first verse in each chapter
    ReDim anPos(oRef.nChapter1 - 1 To oRef.nChapter2 + 1) As Long
    anPos(oRef.nChapter1 - 1) = 0
    anPos(oRef.nChapter2 + 1) = Len(sPass) + 1
    For nChap = oRef.nChapter1 To oRef.nChapter2
      anPos(nChap) = InStr(anPos(nChap - 1) + 1, sPass, nChap & ":")
    Next
   '--Loop through each chapter and replace verse numbers
    nP = 0
    For nChap = oRef.nChapter1 To oRef.nChapter2
      nVerse = Val(Mid(sPass, InStr(anPos(nChap) + 1, sPass, ":") + 1, InStr(anPos(nChap) + 1, sPass, " ") - InStr(anPos(nChap) + 1, sPass, ":") - 1))
      nP = anPos(nChap)
      nP2 = nP
     '--chapter numbers
      Do
        nP2 = nP2 + 1
      Loop Until (Mid(sPass, nP2, 1) = ":" Or nP2 >= Len(sPass))
      sPass2 = sPass2 & "@@" & Mid(sPass, nP, nP2 - nP) & "@@"
     '--Loop through all verses of current chapter
      bOK = True
      nP2 = 0
      Do
        nP = InStr(nP + 2, sPass, CStr(nVerse))
        If (nP = 0 Or nP >= anPos(nChap + 1)) Then
          bOK = False
        Else
          If (nP2 > 0) Then sPass2 = sPass2 & Mid(sPass, nP2 + Len(CStr(nVerse - 1)) + 1, nP - nP2 - Len(CStr(nVerse - 1)) - 1)
          sPass2 = sPass2 & "##" & CStr(nVerse) & "##"
          nVerse = nVerse + 1
          nP2 = nP
        End If
      Loop While (bOK)
      If (nP2 > 0) Then sPass2 = sPass2 & Mid(sPass, nP2 + Len(CStr(nVerse - 1)) + 1, anPos(nChap + 1) - nP2 - Len(CStr(nVerse - 1)) - 1)
    Next
   '---Some cosmetic changes
    sPass2 = Replace(sPass2, ChrW(13) & ChrW(10), " ")
    sPass2 = Replace(sPass2, ".`", ". `")
    sPass2 = Replace(sPass2, ",`", ", `")
    sPass2 = Trim(sPass2)
    sPass = sPass2
  Else
    sPass = ""
  End If
End Sub

Private Function InternalCheckPassage(ByVal sRef As String, ByVal sVersion As String) As Boolean
  Dim sData As String
  Dim bOK As Boolean
  Dim nPos As Long
  Dim bRefHas3 As Boolean
  
  bRefHas3 = False
  If (Left(sRef, 1) = "3") Then bRefHas3 = True
  If (Len(sRef) > 4) Then
    If (LCase(Left(sRef, 4)) = "iii ") Then bRefHas3 = True
  End If
  
  m_oBW.ClipGoToVerse False
  If (bRefHas3) Then
    m_oBW.GoToVerse "Obadiah 1:1"
  Else
    m_oBW.GoToVerse "3John 1:1"
  End If
  
  bOK = True
  m_oBW.SetVersions sVersion & " " & sVersion
  m_oBW.ClipGoToVerse True
  m_oBW.GoToVerse sRef
  sData = bx_oClip.GetText()
  If (Len(sData) > 6) Then
    nPos = InStr(sData, "  ")
    If (bRefHas3) Then
      If (Mid(sData, nPos + 2, 1) = "O") Then bOK = False
    Else
      If (Mid(sData, nPos + 2, 1) = "3") Then bOK = False
    End If
  Else
    bOK = False
  End If
  
  InternalCheckPassage = bOK
End Function
