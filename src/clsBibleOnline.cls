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
Option Explicit
Option Base 1
'==============================================================================
' OnlineBible implementation of IBible interface
' See https://github.com/renehamburger/blinx for source code, manual & license
'==============================================================================
Implements IBible

#If Win64 Then
  Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (lpdwFlags As LongPtr, ByVal dwReserved As Long) As Boolean
#Else
  Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
#End If

Private m_oIE As InternetExplorer
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
  Set m_oIE = New InternetExplorer
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
    If (TypeName(m_oIE) <> "InternetExplorer") Then
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
  If (GetSetting("Blinx", "Options", "OnlineBible", Split(BX_ONLINE_BIBLES, "#")(0)) = "biblegateway.com") Then
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
  Dim oRef As BX_Reference
  Dim oPassageReference As IHTMLElement
  Dim oPassageContent As HTMLDivElement
  Dim oHTML As New HTMLDocument
    
  If (sRefString = m_sLastRefIn And sVersion = m_sLastVersion) Then
    sText = m_sLastText
    sRefString = m_sLastRefOut
  Else
    sVersion = CheckVersion(sVersionIn)
    sText = ""
    sLink = "https://www.biblegateway.com/passage/?search=" & Replace(sRefString, " ", "") & "&version=" & sVersion & "&interface=print"
    If (NavigateTo(sLink)) Then
      oHTML.body.innerHTML = m_oIE.Document.body.innerHTML
      Set oPassageReference = oHTML.querySelector(".passage-display .bcv")
      Set oPassageContent = oHTML.getElementsByClassName("passage-content").Item()
      
      If (Not oPassageReference Is Nothing And Not oPassageContent Is Nothing) Then
        m_sLastRefIn = sRefString
        m_sLastVersion = sVersion
        sHeading = Replace(oPassageReference.innerText, vbCrLf, "")
        sText = ProcessHTML(oPassageContent)
        oRef = BX_StringtToReference(sHeading, False)
        '---Add initial chapter number, if missing
        If (Left(sText, 2) <> "@@") Then sText = "@@" & oRef.nChapter1 & "@@" & sText
        sRefString = oRef.sBook & " " & Right(sRefString, Len(sRefString) - InStrRev(sRefString, " "))
        m_sLastText = sText
        m_sLastRefOut = sRefString
      Else
        bx_oGeneralForm.MsgBox "Unable to retrieve the passage from BibleGateway. If the issue persist and, please check at https://github.com/renehamburger/blinx if your version (" & BX_VERSION_FULL & ") is the latest version. If it is, please create an issue on that page.", vbExclamation
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

Private Function ProcessHTML(ByVal oPassage As HTMLDivElement) As String
  bx_sFunction = "clsBibleOnline_ProcessHTML"
  Dim nPos As Long
  Dim nI As Long
  Dim vContent As Variant
  Dim sTemp As String
  Dim sText As String
  Dim Node As Object
  
  '---Remove headings and crossreference & footnote superscripts
  RemoveNodes oPassage.getElementsByTagName("h3")
  RemoveNodes oPassage.getElementsByClassName("crossreference")
  RemoveNodes oPassage.getElementsByClassName("footnote")

  '---Mark chapter numbers
  For Each Node In oPassage.getElementsByClassName("chapternum")
    Node.innerText = "@@" & Node.innerText & "@@"
  Next Node
  
  '---Mark verse numbers
  For Each Node In oPassage.getElementsByClassName("versenum")
    Node.innerText = "##" & Node.innerText & "##"
  Next node

  '---Convert actual passage text to string
  For Each node In oPassage.getElementsByClassName("text")
    sText = sText & " " & node.innerText
  Next Node

  '---Clean string
  sText = Trim(sText)
  Do While (InStr(sText, "  ") > 0)
    sText = Replace(sText, "  ", " ")
  Loop
  
  ProcessHTML = sText
End Function

Private Sub RemoveNodes(ByRef nodes As Object)
  Dim Node As Object
  For Each Node In nodes
    Node.ParentNode.RemoveChild Node
  Next Node
End Sub

Private Function CheckVersion(ByVal sVersion As String) As String
  bx_sFunction = "clsBibleOnline_CheckVersion"
  If (sVersion = "NIB") Then
    CheckVersion = "NIVUK"
  Else
    CheckVersion = sVersion
  End If
End Function

Private Function Debug_LoadHTMLfromFile(ByVal sFile As String)
  bx_sFunction = "clsBibleOnline_LoadHTMLfromFile"
  Dim oFSO As New FileSystemObject
  Dim oTextStream As TextStream

  Set oTextStream = oFSO.OpenTextFile(sFile, ForReading, False, TristateTrue)
  Debug_LoadHTMLfromFile = oTextStream.ReadAll
End Function
