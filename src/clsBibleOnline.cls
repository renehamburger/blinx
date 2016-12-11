VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBibleOnline"
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
' OnlineBible implementation of IBible interface
' (See "ThisDocument" for license)
'==============================================================================
Implements IBible

#If Win64 Then
  Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (lpdwFlags As LongPtr, ByVal dwReserved As Long) As Boolean
#Else
  Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
#End If

Private m_oIE As Object
Private m_sHTML As String
Private m_sLastRefIn As String
Private m_sLastRefOut As String
Private m_sLastVersion As String
Private m_sLastText As String
Private m_bFirstRun As Boolean

'==============================================================================
'   Initialisation/Termination
'==============================================================================

Private Sub Class_Initialize()
  On Error GoTo FAILED
  
  bx_sFunction = "clsBibleOnline_Initialize"
  Set m_oIE = CreateObject("InternetExplorer.Application")
  m_oIE.Visible = False
  m_bFirstRun = True
  
  Exit Sub
FAILED:
  Set m_oIE = Nothing
End Sub

Private Sub Class_Terminate()
  bx_sFunction = "clsBibleOnline_Terminate"
  If (IBible_IsApplicationOK()) Then m_oIE.Quit
  Set m_oIE = Nothing
End Sub

'==============================================================================
'   Public Functions
'==============================================================================

Public Function IBible_IsApplicationOK() As Boolean
  bx_sFunction = "clsBibleOnline_IsApplicationOK"
  
  If (Not m_oIE Is Nothing) Then
    If (TypeName(m_oIE) <> "Object") Then
      IBible_IsApplicationOK = True
    Else
      IBible_IsApplicationOK = False
    End If
  Else
    IBible_IsApplicationOK = False
  End If
End Function

Public Function IBible_IsReady() As Boolean
  bx_sFunction = "clsBibleOnline_IsReady"

  If (IBible_IsApplicationOK() And InternetGetConnectedState(0&, 0&)) Then
    IBible_IsReady = True
  Else
    If (IsEmpty(bx_vTimeWarningOnline)) Then bx_vTimeWarningOnline = Time - 0.1
    If ((Time - bx_vTimeWarningOnline) * 3600 * 24 > 30) Then
      bx_oGeneralForm.MsgBox "No internet connection detected.", vbExclamation
      bx_vTimeWarningOnline = Time
    End If
    IBible_IsReady = False
  End If
End Function


Public Function IBible_BeginProcess(ByVal sVersion As String) As Boolean
  bx_sFunction = "clsBibleOnline_BeginProcess"
  IBible_BeginProcess = IBible_IsReady()
End Function

Public Sub IBible_EndProcess()
  bx_sFunction = "clsBibleOnline_EndProcess"
End Sub

Public Function IBible_CheckPassage(ByVal sRef As String, ByVal sVersion As String, Optional ByVal bFirstVerseOnly As Boolean = False) As Boolean
  Dim oRef As BX_Reference
  
  If (bFirstVerseOnly) Then
    oRef = BX_StringtToReference(sRef)
    If (oRef.nVerse1 = 0) Then oRef.nVerse1 = 1
    IBible_CheckPassage = (GetOnlinePassage(oRef.sBook & " " & oRef.nChapter1 & ":" & oRef.nVerse1, sVersion, True) <> "")
  Else
    IBible_CheckPassage = (GetOnlinePassage(sRef, sVersion, True) <> "")
  End If
End Function

Public Function IBible_GetPassage(ByRef sRef As String, ByVal sVersion As String, Optional ByVal bSuppressError As Boolean = False) As String
  bx_sFunction = "clsBibleOnline_GetPassage"
  IBible_GetPassage = GetOnlinePassage(sRef, sVersion, bSuppressError)
End Function

Public Function IBible_GotoPassage(ByVal sRef As String, ByVal sVersion As String) As Boolean
  bx_sFunction = "clsBibleOnline_GotoVerse"
  Dim sLink As String
  If (GetSetting("Blinx", "Options", "OnlineBible", Split(BX_ONLINE_BIBLE, "#")(0)) = "biblegateway.com") Then
    sLink = "https://www.biblegateway.com/passage/?search=" & Replace(sRef, " ", "%20") & "&version=" & sVersion
  Else
    sLink = "http://www.esvbible.org/search/" & sRef
  End If
  
  IBible_GotoPassage = NavigateTo(sLink)
End Function


'==============================================================================
'   Private Functions
'==============================================================================

Private Function GetOnlinePassage(ByRef sRefString As String, ByVal sVersionIn As String, Optional ByVal bSuppressError As Boolean = False) As String
  bx_sFunction = "clsBibleOnline_GetOnlinePassage"
  Dim sLink As String
  Dim sHeading As String
  Dim sText As String
  Dim sVersion As String
  Dim nPosHeading(2) As Long
  Dim nPosBlock(2) As Long
  Dim oRef As BX_Reference
  
  If (sRefString = m_sLastRefIn And sVersion = m_sLastVersion) Then
    sText = m_sLastText
    sRefString = m_sLastRefOut
  Else
    sVersion = CheckVersion(sVersionIn)
    sText = ""
    sLink = "https://www.biblegateway.com/passage/index.php?search=" & Replace(sRefString, " ", "%20") & "&version=" & sVersion & "&interface=print"
    If (NavigateTo(sLink)) Then
      m_sHTML = m_oIE.Document.body.innerHTML
      m_sHTML = Replace(m_sHTML, ChrW(10), "")
      m_sHTML = Replace(m_sHTML, ChrW(13), "")
      
'---debug: (for loading older HTML)
'Debug_LoadHTMLfromFile "D:\\Rene\\Misc\\programming\\VBA\\Blinx\\resources\\WinXP-IE8\\John1-2.txt"

      nPosHeading(1) = FindInHTML(1, "<span class=""passage-display-bcv"">", True)
      nPosHeading(2) = FindInHTML(nPosHeading(1), "</span>", False)
      nPosBlock(1) = FindInHTML(nPosHeading(2), "</h1>", True)
      nPosBlock(2) = FindInHTML(nPosBlock(1), "<div class=""footnotes"">", False)
      If (nPosHeading(1) > 0 And nPosHeading(2) > 0 And nPosBlock(1) > 0 And nPosBlock(2) > 0) Then
        m_sLastRefIn = sRefString
        m_sLastVersion = sVersion
        sHeading = Mid(m_sHTML, nPosHeading(1), nPosHeading(2) - nPosHeading(1) + 1)
        sText = Mid(m_sHTML, nPosBlock(1), nPosBlock(2) - nPosBlock(1) + 1)
        ProcessHTML sText
        oRef = BX_StringtToReference(sHeading, False)
       '---Add initial chapter number, if missing
        If (Left(sText, 2) <> "@@") Then sText = "@@" & oRef.nChapter1 & "@@" & sText
        sRefString = oRef.sBook & " " & Right(sRefString, Len(sRefString) - InStrRev(sRefString, " "))
        m_sLastText = sText
        m_sLastRefOut = sRefString
      End If
    End If
  End If
  
  GetOnlinePassage = sText
End Function

Private Function NavigateTo(ByVal sLink As String) As Boolean
  bx_sFunction = "clsBibleOnline_NavigateTo"
  Dim bOK As Boolean
  Dim vTime As Variant
  bOK = False
  vTime = Time
  
  m_oIE.navigate sLink
  Do
    DoEvents
    If (Not m_oIE.ReadyState <> 4 And m_oIE.LocationURL = sLink) Then
      bOK = True
    End If
    If (Not bOK) Then
      If (m_bFirstRun And (Time - vTime) * 24 * 3600 > 20) Then
        bx_oGeneralForm.MsgBox "Passage could not be accessed online. Please try again.", vbExclamation
        NavigateTo = False
        m_bFirstRun = False
        Exit Function
      ElseIf (Not m_bFirstRun And (Time - vTime) * 24 * 3600 > 10) Then
        bx_oGeneralForm.MsgBox "Passage could not be accessed online.", vbExclamation
        NavigateTo = False
        Exit Function
      End If
    End If
  Loop Until (bOK)
  NavigateTo = True
  m_bFirstRun = False
End Function

Private Function FindInHTML(ByVal nFrom As Long, ByVal sStringIn As String, ByVal bLeftEnd As Boolean)
  bx_sFunction = "clsBibleOnline_FindInHTML"
  Dim nPos As Long
  Dim sStringTemp As String
  
  '--Older versions of IE appear to have different quotations marks (or even none) for strings within html tags
  nPos = 0
  sStringTemp = sStringIn
  If (nFrom > 0) Then
    nPos = InStr(nFrom, m_sHTML, sStringTemp, vbTextCompare)
    If (nPos = 0) Then
      sStringTemp = Replace(sStringTemp, "'", """")
      nPos = InStr(nFrom, m_sHTML, sStringTemp, vbTextCompare)
      If (nPos = 0) Then
        sStringTemp = Replace(sStringTemp, "'", "")
        nPos = InStr(nFrom, m_sHTML, sStringTemp, vbTextCompare)
      End If
    End If
  End If
  If (nPos > 0) Then
    If (bLeftEnd) Then
      nPos = nPos + Len(sStringTemp)
    Else
      nPos = nPos - 1
    End If
  End If
  FindInHTML = nPos
End Function

Private Sub ProcessHTML(ByRef sText As String)
  bx_sFunction = "clsBibleOnline_ProcessHTML"
  Dim nPos As Long
  Dim nI As Long
  Dim vContent As Variant
  Dim sTemp As String
  
'---Remove all further headings
  nPos = 1
  Do While (nPos > 0)
    RemoveNextPair sText, nPos, "<h3>", "</h3>"
  Loop
'---Reduce and mark chapter numbers
  nPos = 1
  Do While (nPos > 0)
    RemoveNextPair sText, nPos, "<span class=""chapternum"">", "</span>", vContent
    nI = InStrRev(vContent, " ")
    If (nI > 0) Then
      sText = Left(sText, nPos - 1) & " @@" & Right(vContent, Len(vContent) - nI) & "@@" & Right(sText, Len(sText) - nPos + 1)
    End If
    vContent = ""
  Loop
'---Reduce and mark verse numbers and remove footnotes
  nPos = 1
  Do While (nPos > 0)
    RemoveNextPair sText, nPos, "<sup", "</sup>", vContent
    sTemp = "class='versenum'>"
    nI = InStrHTML(1, vContent, sTemp)
    If (nI > 0) Then
      sText = Left(sText, nPos - 1) & "##" & Right(vContent, Len(vContent) - nI - Len(sTemp) + 1) & "##" & Right(sText, Len(sText) - nPos + 1)
    End If
    vContent = ""
  Loop
'---Remove all remaining brackets plus content
  nPos = 1
  Do While (nPos > 0)
    RemoveNextPair sText, nPos, "<", ">"
  Loop
  nPos = 1
  Do While (nPos > 0)
    RemoveNextPair sText, nPos, "[", "]"
  Loop
'---Remove spaces & cr/lf
  sText = Replace(sText, Chr(10), " ")
  sText = Replace(sText, Chr(13), " ")
  sText = Replace(sText, "&nbsp;", " ")
  sText = Trim(sText)
  Do While (InStr(sText, "  ") > 0)
    sText = Replace(sText, "  ", " ")
  Loop
  sText = Replace(sText, "## ", "##")
  sText = Replace(sText, "@@ ", "@@")
End Sub

Private Sub RemoveNextPair(ByRef sText As String, ByRef nPos As Long, ByVal sLeftIn As String, ByVal sRightIn As String, Optional vContent As Variant)
  bx_sFunction = "clsBibleOnline_RemoveNextPair"
  Dim nPosRight As Long
  Dim nPosLeft As Long
  Dim sLeft As String
  Dim sRight As String
  
  sRight = sRightIn
  nPosRight = InStrHTML(nPos, sText, sRight)
  If (nPosRight > 0) Then
    sLeft = sLeftIn
    nPosLeft = InStrHTML(nPosRight, sText, sLeft, True)
    If (nPosLeft > 0) Then
      If (Not IsMissing(vContent)) Then vContent = Mid(sText, nPosLeft + Len(sLeft), nPosRight - nPosLeft - Len(sLeft))
      sText = Left(sText, nPosLeft - 1) & Right(sText, Len(sText) - nPosRight - Len(sRight) + 1)
      nPos = nPosLeft
    Else
      nPos = 0
    End If
  Else
    nPos = 0
  End If
End Sub

Private Function InStrHTML(ByVal nPos As Long, ByVal sSearchIn As String, ByRef sSearchFor As String, Optional ByVal bReverse As Boolean = False) As Long
  bx_sFunction = "clsBibleOnline_InStrHTML"
  Dim nResult As Long
  Dim sTemp As String
  
'--Search for string as it is
  sTemp = sSearchFor
  If (Not bReverse) Then
    nResult = InStr(nPos, sSearchIn, sTemp, vbTextCompare)
  Else
    nResult = InStrRev(sSearchIn, sTemp, nPos, vbTextCompare)
  End If
'--Search for string with " instead of '
  If (nResult = 0) Then
    sTemp = Replace(sSearchFor, "'", """")
    If (Not bReverse) Then
      nResult = InStr(nPos, sSearchIn, sTemp, vbTextCompare)
    Else
      nResult = InStrRev(sSearchIn, sTemp, nPos, vbTextCompare)
    End If
  End If
'--Search for string without '
  If (nResult = 0) Then
    sTemp = Replace(sSearchFor, "'", "")
    If (Not bReverse) Then
      nResult = InStr(nPos, sSearchIn, sTemp, vbTextCompare)
    Else
      nResult = InStrRev(sSearchIn, sTemp, nPos, vbTextCompare)
    End If
  End If
  
  sSearchFor = sTemp
  InStrHTML = nResult
End Function

Private Function CheckVersion(ByVal sVersion As String) As String
  bx_sFunction = "clsBibleOnline_CheckVersion"
  If (sVersion = "NIB") Then
    CheckVersion = "NIVUK"
  Else
    CheckVersion = sVersion
  End If
End Function

Private Sub Debug_LoadHTMLfromFile(ByVal sFile As String)
  bx_sFunction = "clsBibleOnline_LoadHTMLfromFile"
  Dim oFSO As New FileSystemObject
  Dim oTextStream As TextStream

  Set oTextStream = oFSO.OpenTextFile(sFile, ForReading, False, TristateTrue)
  m_sHTML = oTextStream.ReadAll
End Sub
