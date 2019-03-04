Attribute VB_Name = "BX"
Option Explicit
Option Base 1
'==============================================================================
' See https://github.com/renehamburger/blinx for source code, manual & license
'==============================================================================

Private m_eOptions As BX_Options

'==============================================================================
'   Macros
'==============================================================================

Public Sub BX_CreateBlink() 'AltGr-B
  bx_sFunction = "BX_CreatePureBlink"
  m_eOptions = BX_PURE
  BX_MacroEntryPoint
End Sub

Public Sub BX_CreateBlinkWithTooltip() 'Alt-B
  bx_sFunction = "BX_CreateBlinkWithTooltip"
  If (Selection.StoryType = wdMainTextStory) Then m_eOptions = BX_TOOLTIP Else m_eOptions = BX_PURE
  BX_MacroEntryPoint
End Sub

Public Sub BX_CreateBlinkWithText() 'Shift-Alt-B
  bx_sFunction = "BX_CreateBlinkWithText"
  m_eOptions = BX_TEXT
  BX_MacroEntryPoint
End Sub

Public Sub BX_UnlinkAllLinks() 'Alt-U
  bx_sFunction = "BX_UnlinkAllLinks"
  BX_UnlinkSelection
End Sub


'==============================================================================
'   Public Functions
'==============================================================================

Public Function BX_GetPassage(ByRef sRef As String, Optional ByVal vDisplayVer As Variant, Optional ByVal bSuppressError As Boolean = False) As String
  On Error GoTo ERR_HANDLER
  bx_sFunction = "BX_GetPassage"
  
  BX_CheckForms
  If (bx_sVariablesLoaded <> "true") Then BX_LoadVariables
  If (BX_InitializeBible()) Then
    If (bx_oB.BeginProcess()) Then
      If (IsMissing(vDisplayVer)) Then vDisplayVer = GetSetting("Blinx", "Options", "Version", Split(BX_TRANSLATIONS, "#")(0))
      BX_GetPassage = bx_oB.GetPassage(sRef, vDisplayVer, bSuppressError)
    End If
    bx_oB.EndProcess
  End If
  
EXIT_SUB:
  Exit Function
ERR_HANDLER:
  BX_ErrHandler
End Function



'==============================================================================
'   Private Functions
'==============================================================================

Private Sub BX_ErrHandler()
  If (bx_oGeneralForm Is Nothing) Then Set bx_oGeneralForm = New BX_GeneralForm
  
  If (bx_oGeneralForm Is Nothing) Then
    InputBox "An unexpected error has occured. Please contact blinx.add.in@gmail.com with the following details:", _
             Default:=BX_VERSION_FULL & ";" & bx_sFunction & ";" & m_eOptions & ";" & Err.Number & " - " & Err.Description & ";" & Replace(Selection.Text, vbCr, "<vbCr>")
  Else
    bx_oGeneralForm.ErrorBox "An unexpected error has occured. Please contact blinx.add.in@gmail.com with the following details:", _
    "Blinx Error", BX_VERSION_FULL & ";" & bx_sFunction & ";" & m_eOptions & ";" & Err.Number & " - " & Err.Description & ";" & Replace(Selection.Text, vbCr, "<vbCr>")
  End If
End Sub

Private Sub BX_MacroEntryPoint()
  On Error GoTo ERR_HANDLER
  bx_sFunction = "BX_MacroEntryPoint"
  Dim sVersion As String
  sVersion = GetSetting("Blinx", "Options", "Version", Split(BX_TRANSLATIONS, "#")(0))

'---debug:
  bx_sCurrentDocument = ActiveDocument.FullName
  
  BX_CheckForms
  If (bx_sVariablesLoaded <> "true") Then BX_LoadVariables
  If (BX_InitializeBible()) Then
    If (bx_oB.BeginProcess(sVersion)) Then
      BX_HandleFindParameters True
      If (Selection.Start = Selection.End) Then
        BX_CreateBlinkToLeft m_eOptions
      Else
        BX_CreateAllBlinks m_eOptions
      End If
      BX_HandleFindParameters False
    End If
    bx_oB.EndProcess
  End If
  
EXIT_SUB:
  Exit Sub
ERR_HANDLER:
  BX_ErrHandler
End Sub

Private Function BX_InitializeBible() As Boolean
  bx_sFunction = "BX_InitializeBible"
  Dim sName As String
  Dim bOK As Boolean
  bOK = True

  '--Delete object, if connection to application not established
  If (Not bx_oB Is Nothing) Then
    If (Not bx_oB.IsApplicationOK()) Then Set bx_oB = Nothing
  End If
  
  '--Create BW object first
  If (bx_oB Is Nothing) Then Set bx_oB = New clsBibleBW
  
  '--If connection to BW application cannot established, create IE object
  If (Not bx_oB.IsApplicationOK()) Then
    Set bx_oB = Nothing
    Set bx_oB = New clsBibleOnline
    If (Not bx_oB.IsApplicationOK()) Then
      bx_oGeneralForm.MsgBox "Neither BibleWorks nor BibleGateway (via Internet Explorer) could be found. Blinx cannot proceed.", vbExclamation
      bOK = False
    End If
  End If
  
  BX_InitializeBible = bOK
End Function

Private Sub BX_CreateAllBlinks(ByVal m_eOptions As BX_Options)
  bx_sFunction = "BX_CreateAllBlinks"
  
  Dim eReturn As BX_ReturnValue
  Dim bAcceptAll As Boolean
  Dim bCheckForPartial As Boolean
  Dim bOK As Boolean
  Dim bSkip As Boolean
  Dim bWithChVsSeparator As Boolean
  Dim nI As Long
  Dim nJ As Long
  Dim nLinksCreated As Long
  Dim nPosDiff As Long
  Dim nOldLength As Long
  Dim nOldRangeLength As Long
  Dim oLink As Hyperlink
  Dim oRange As Range
  Dim oOldRange As Range
  Dim oRef As BX_Reference
  Dim sRef As String
  Dim sText As String
  Dim sPass As String
  Dim sVersion As String

  '---Prep
  sVersion = GetSetting("Blinx", "Options", "Version", Split(BX_TRANSLATIONS, "#")(0))
  Set oOldRange = Selection.Range
  nOldRangeLength = oOldRange.End - oOldRange.Start
  nLinksCreated = 0
  bCheckForPartial = False
  bAcceptAll = False
  Set oLink = Nothing
  nPosDiff = 0
  bx_oRefPositions.Clear
  
  '---If no selection, get main text range
  If (Selection.Start = Selection.End) Then
    For Each oRange In ActiveDocument.StoryRanges
      If oRange.StoryType = wdMainTextStory Then oRange.Select
    Next
  End If
  Set oRange = Selection.Range
  
  '---Replace all BibleWorks javascript links
  nLinksCreated = nLinksCreated + BX_ReplaceBWHyperlinks(m_eOptions)
      
  '---Loop backwards and replace all complete references
  Selection.Start = Selection.End
  Do While (Selection.Start > oRange.Start)
    eReturn = BX_CreateBlinkToLeft(m_eOptions, True, True, oRange.Start)
    Select Case eReturn
      Case BX_FAILED: bCheckForPartial = True
      Case BX_LINK_CREATED: nLinksCreated = nLinksCreated + 1
      Case BX_ALREADY_LINK: nLinksCreated = nLinksCreated + 1
    End Select
    Selection.End = Selection.Start
  Loop
  
  Selection.Start = oRange.Start
  
  '---Loop forwards through all references and replace all partial references,
  '--- taking book (and chapter) from the last successful reference
  If (bCheckForPartial) Then
    bOK = True
    For nI = bx_oRefPositions.Size() To 1 Step -3
      If (bx_oRefPositions.data(nI - 2)) Then
        nPosDiff = nPosDiff + bx_oRefPositions.data(nI - 1)
        bx_oLastValidRef = BX_StringtToReference(bx_oRefPositions.data(nI), False)
      Else
        Selection.Start = bx_oRefPositions.data(nI - 1) + nPosDiff
        Selection.End = bx_oRefPositions.data(nI) + nPosDiff
          
        '---Check for further expansions of partial reference, some of which would be discarded
        bSkip = Not BX_ExpandSelectedPartialReference() 'expansions to the left might have already been recognised

        '---Further checks if just numbers
        If (Not bSkip) Then
          If (BX_TestChar(Selection.Text, BX_NUMBER)) Then
            '--reject if superscript
            If (Selection.Font.Superscript) Then bSkip = True
            '--reject if at beginning of line if not in table, as almost always numbering; verse would be OK if "v1" or "v.1" or "verse 1")
            If (Not bSkip And Selection.Tables.Count() = 0) Then
              If (Selection.Start = 0) Then
                bSkip = True
              Else
                Selection.Start = Selection.Start - 1
                If (BX_TestChar(Left(Selection.Text, 1), BX_RETURN)) Then bSkip = True
                Selection.Start = Selection.Start + 1
              End If
            End If
          End If
        End If
        
        If (Not bSkip) Then
          nOldLength = Selection.End - Selection.Start
          BX_CheckReference Selection.Text, oRefbook
          BX_CompletePartialReference oRef
          sRef = BX_ReferenceToString(oRef)
          If (sRef <> "" And Not bAcceptAll) Then
            If (Not bx_oProcessInputForm.oWindow Is ActiveWindow) Then
              Set bx_oProcessInputForm = Nothing
              Set bx_oProcessInputForm = New BX_ProcessInputForm
              bx_oProcessInputForm.SetWindow ActiveWindow
            End If
            bx_oProcessInputForm.bProcess = True
            bx_oProcessInputForm.sRef = BX_ReferenceToString(oRef, bx_eLanguage)
            bx_oProcessInputForm.Show vbModal
            sRef = bx_oProcessInputForm.sRef
            bAcceptAll = bx_oProcessInputForm.bAcceptAll
            bOK = Not bx_oProcessInputForm.bAbort
          End If
          If (bOK And sRef <> "") Then
            BX_CheckReference sRef, oRef
            sRef = BX_ReferenceToString(oRef, BX_ENGLISH)
            sPass = bx_oB.GetPassage(sRef, sVersion, bAcceptAll)
            If (sPass <> "") Then
              sText = Selection.Text
              Set oLink = BX_FillBlink(m_eOptions, sText, sPass, sRef, sVersion)
              If (Not oLink Is Nothing) Then
                bx_oLastValidRef = BX_StringtToReference(sRef, False)
                nLinksCreated = nLinksCreated + 1
                nPosDiff = nPosDiff + (Selection.End - oLink.Range.Start - nOldLength) 'Difference due to hyperlink field and/or endnote symbol and/or passage text
              End If
            End If
          End If
        End If
      End If
      If (Not bOK) Then Exit For
    Next
  End If
  Selection.Start = oOldRange.Start
  Selection.End = oOldRange.Start + nOldRangeLength + nPosDiff
  '--- If only one link within selection and the reference at the end, include whole selection in link
  If (Not oLink Is Nothing And nLinksCreated = 1 And Selection.End <> Selection.Start) Then
    If ((oLink.Range.End = Selection.End) Or (oLink.Range.End = Selection.End - 2 And BX_TestChar(Right(Selection.Text, 1), BX_RETURN))) Then
      If (vbYes = bx_oGeneralForm.MsgBox("Should the blink be expanded to contain the whole selection?", vbYesNo)) Then
        Selection.End = oLink.Range.Start
        oLink.TextToDisplay = Selection.Text & oLink.TextToDisplay
        Selection.Delete
        oOldRange.Select
      End If
    End If
  End If
End Sub

Private Function BX_CreateBlinkToLeft(ByVal m_eOptions As BX_Options, Optional ByVal bSuppressError As Boolean = False, Optional ByVal bPartOfAllLinks As Boolean = False, Optional ByVal nStartOfScope = 0) As BX_ReturnValue
  bx_sFunction = "BX_CreateBlinkToLeft"
  Dim nI As Long
  Dim nJ As Long
  Dim nK As Long
  Dim nOldStart As Long
  Dim nOldEnd As Long
  Dim nOldStartDiff As Long
  Dim nOldEndDiff As Long
  Dim nOldLength As Long
  Dim sRef As String
  Dim oLink As Hyperlink
  Dim bValid As Boolean
  Dim bPartialReference As Boolean
  Dim eReturn As BX_ReturnValue
  Dim sVersion As String
  Dim nCount As Long
  Dim bTemp As Boolean
  Dim oRef As BX_Reference
    
'---Prep
  sVersion = GetSetting("Blinx", "Options", "Version", Split(BX_TRANSLATIONS, "#")(0))
  Set oLink = Nothing
  eReturn = BX_UNDEFINED
  bValid = True
  bPartialReference = False
  nOldStart = Selection.Start
  nOldEnd = Selection.End

'---Move let until possible reference found
  Selection.End = Selection.Start
  '--go back to first number
  Selection.Find.Forward = False
  Selection.Find.Text = "^#"
  If (Selection.Find.Execute) Then
    If (Selection.Start >= nStartOfScope) Then
      'If (bPartOfAllLinks) Then nNumberPos = Selection.Start
      nOldStartDiff = nOldStart - Selection.End
      nOldEndDiff = nOldEnd - Selection.End
      If (Selection.Hyperlinks.Count() > 0) Then
        If (Len(Selection.Hyperlinks(1).Address) > 30) Then
          If (Left(Selection.Hyperlinks(1).Address, 27) = "http://www.biblegateway.com") Then bValid = False
          If (Left(Selection.Hyperlinks(1).Address, 28) = "http://www.esvstudybible.org") Then bValid = False 'former website of esvonline
          If (Left(Selection.Hyperlinks(1).Address, 24) = "http://www.esvonline.org") Then bValid = False
          If (Not bValid) Then
            BX_GetDataFromLink Selection, bx_oLastValidRef
            eReturn = BX_ALREADY_LINK
          End If
        End If
        Selection.Start = Selection.Hyperlinks(1).Range.Start
        Selection.End = Selection.Start
      End If
      
      If (bValid) Then
        bValid = BX_ExpandReference()
       '--Recalculate diff to old position
        nOldStartDiff = nOldStart - Selection.End
        nOldEndDiff = nOldEnd - Selection.End
      End If
    Else
      bValid = False
      eReturn = BX_NOTHING
    End If
  Else
    bValid = False
    eReturn = BX_NOTHING
    Selection.End = 0
    Selection.Start = 0
  End If
  
'---Attempt to convert selection
  If (bValid) Then
    nOldLength = Selection.End - Selection.Start
    BX_CheckReference Selection.Text, oRef
    '--For invalid book name, reduce selection to numbers
    If (oRef.sBook = "invalid") Then
      For nI = Len(Selection.Text) To 1 Step -1
        If (BX_TestChar(Mid(Selection.Text, nI, 1), BX_SPACE)) Then
          Selection.Start = Selection.Start + nI
          Exit For
        End If
      Next
      oRef.sBook = ""
    End If
    
    '--If links valid, create them
    If (oRef.sBook <> "invalid" And oRef.sBook <> "") Then
      Set oLink = BX_CreateBlinkHere(m_eOptions, BX_ReferenceToString(oRef, BX_ENGLISH), sVersion)
      If (Not oLink Is Nothing) Then
        eReturn = BX_LINK_CREATED
      Else
        eReturn = BX_FAILED
      End If
    '--Partial link: user input if not part of "convert all"
    ElseIf (Not bPartOfAllLinks) Then
      BX_CompletePartialReference oRef
      
      If (oRef.sBook <> "invalid") Then
        If (Not bx_oProcessInputForm.oWindow Is ActiveWindow) Then
          Set bx_oProcessInputForm = Nothing
          Set bx_oProcessInputForm = New BX_ProcessInputForm
          bx_oProcessInputForm.SetWindow ActiveWindow
        End If
        bx_oProcessInputForm.bProcess = False
        bx_oProcessInputForm.sRef = BX_ReferenceToString(oRef, bx_eLanguage)
        bx_oProcessInputForm.Show vbModal
        sRef = bx_oProcessInputForm.sRef
      End If
      If (sRef <> "") Then
        BX_CheckReference sRef, oRef
        sRef = BX_ReferenceToString(oRef, BX_ENGLISH)
        Set oLink = BX_CreateBlinkHere(m_eOptions, sRef, sVersion)
        If (Not oLink Is Nothing) Then
          eReturn = BX_LINK_CREATED
        Else
          eReturn = BX_FAILED
        End If
      Else
        eReturn = BX_FAILED
      End If
    Else
      eReturn = BX_FAILED
    End If
  End If

'--- Move to best starting position for next search or back to original position
  If (bPartOfAllLinks) Then
    If (eReturn = BX_LINK_CREATED) Then
      bx_oRefPositions.Add True 'Valid link
      bx_oRefPositions.Add Selection.End - oLink.Range.Start - nOldLength 'Difference due to hyperlink field and/or endnote symbol and/or inserted text
      bx_oRefPositions.Add BX_ReferenceToString(oRef) 'Reference
      Selection.End = oLink.Range.Start
      Selection.Start = oLink.Range.Start
    ElseIf (eReturn = BX_ALREADY_LINK) Then
      bx_oRefPositions.Add True 'Valid link
      bx_oRefPositions.Add 0 'No difference
      bx_oRefPositions.Add BX_ReferenceToString(bx_oLastValidRef) 'Reference
    Else
      If (bValid) Then
        bx_oRefPositions.Add False 'Invalid link
        bx_oRefPositions.Add Selection.Start
        bx_oRefPositions.Add Selection.End
      End If
      Selection.End = Selection.Start
      Selection.Start = Selection.Start
    End If
  Else
    nI = Selection.End
    Selection.Start = nI + nOldStartDiff
    Selection.End = nI + nOldEndDiff
  End If
  
  BX_CreateBlinkToLeft = eReturn
End Function

Private Function BX_ReplaceBWHyperlinks(ByVal m_eOptions As BX_Options) As Integer
  bx_sFunction = "BX_ReplaceBWHyperlinks"
  Dim oLink As Hyperlink
  Dim nPos As Long
  Dim sRef As String
  Dim sPass As String
  Dim sPass2 As String
  Dim oRange As Range
  Dim nLinksCreated As Integer
  Dim sVersion As String
  
  sVersion = GetSetting("Blinx", "Options", "Version", Split(BX_TRANSLATIONS, "#")(0))
  Set oRange = Selection.Range
  nLinksCreated = 0
  
  For Each oLink In ActiveDocument.Hyperlinks
    If (oLink.Range.Start >= oRange.Start And oLink.Range.End <= oRange.End) Then
      If (Left(oLink.Address, 14) = "javascript:R('") Then
        oLink.Range.Select
        nPos = InStr(oLink.Address, "')")
        sRef = Mid(oLink.Address, 15, nPos - 15)
        BX_ReplaceReservedCharacters sRef
        sPass = bx_oB.GetPassage(sRef, sVersion & " " & sVersion, True)
        oLink.Range.Select
        BX_FillBlink m_eOptions, oLink.TextToDisplay, sPass, sRef, sVersion
        nLinksCreated = nLinksCreated + 1
      End If
    End If
  Next
  
  oRange.Select
End Function

Private Function BX_CreateBlinkHere(ByVal m_eOptions As BX_Options, ByVal sRef As String, ByVal sVersion As String) As Hyperlink
  bx_sFunction = "BX_CreateBlinkHere"
  Dim sPass As String
  Dim sText As String
  Dim oLink As Hyperlink
  
  sPass = ""
  If (m_eOptions = BX_PURE) Then
    If (bx_oB.CheckPassage(sRef, sVersion, True)) Then
      sPass = "OK"
    End If
  Else
    sPass = bx_oB.GetPassage(sRef, sVersion, True)
  End If
  If (Len(sPass) > 0) Then
    sText = Selection.Text
    Set oLink = BX_FillBlink(m_eOptions, sText, sPass, sRef, sVersion)
    bx_oLastValidRef = BX_StringtToReference(sRef, False)
    Set BX_CreateBlinkHere = oLink
  Else
    Set BX_CreateBlinkHere = Nothing
  End If
End Function

Private Function BX_FillBlink(ByVal m_eOptions As BX_Options, ByVal sText As String, ByVal sPassIn As String, ByVal sRef As String, ByVal sVersion As String) As Hyperlink
  bx_sFunction = "BX_FillBlink"
  Const sLINKVER = "BX_BibRef_0.6" 'introduced in v0.6
  Dim oLink As Hyperlink
  Dim oEndnote As Endnote
  Dim sPass As String
  Dim sBeforePass As String
  Dim sAfterPass As String
  Dim sAddress As String
  Dim sSubAddress As String
  Dim sTextToDisplay As String
  Dim sTarget As String
  Dim sScreenTip As String
  Dim vPreviewLimit As Variant
  Dim nPreviewLimit As Integer
  Dim oRange As Range
  Dim nVertical As Long
  Dim nHorizontal As Long
  Dim sStandardFontStyle As String
  Dim sBibleTextFontStyle As String
  Dim sHTML As String
  
  Application.ScreenUpdating = False
  
'---Create content
  vPreviewLimit = GetSetting("Blinx", "Options", "BlinkPreviewLength", Split(BX_BLINK_PREVIEW_LENGTHS, "#")(0))
  sBeforePass = sRef & " " & ChrW(&H2013&) & " "
  sAfterPass = " (" & sVersion & ")"
  Select Case GetSetting("Blinx", "Options", "OnlineBible", Split(BX_ONLINE_BIBLES, "#")(0))
    Case "esvbible.org"
      sAddress = "http://www.esvonline.org/search/" & sRef
    Case "bibleserver.com"
      sAddress = "http://www.bibleserver.com/text/" & sVersion & "/" & sRef   'Matth%C3%A4us5%2C44-45
    Case Default
      sAddress = "http://www.biblegateway.com/passage/?search=" & sRef & "&version=" & sVersion
  End Select
  sAddress = Replace(sAddress, " ", "%20")
  sSubAddress = ""
  sTextToDisplay = Replace(sText, " ", ChrW(&H2008&))  'this special space character will keep the whole reference together
  sTarget = sLINKVER & ";" & sVersion & ";" & sRef
  If (m_eOptions = BX_TOOLTIP) Then
    sScreenTip = sRef & vbCr & "Preview passage: move onto link (not from the left)" & vbCr & "Look up in BibleWorks: right-click" & vbCr & "Look up online:"
  Else
    sScreenTip = sRef & vbCr & "Look up in BibleWorks: right-click" & vbCr & "Look up online:"
  End If
  
  
'---debug:
If (ActiveDocument.FullName <> bx_sCurrentDocument) Then
 If (GetSetting("Blinx", "Options", "Debugger", "") = "1") Then
  Stop
 End If
End If
  
   
  Set oLink = AddHyperlink(sAddress, sSubAddress, sScreenTip, sTextToDisplay, sTarget)
  oLink.Range.Fields.Update
  oLink.Range.Select
  
  If (m_eOptions = BX_TOOLTIP) Then
    '---Determine font
    sStandardFontStyle = "font-family:Calibri, sans-serif; font-size: 6pt;"
    Select Case (sVersion)
      Case "BGT"
        sBibleTextFontStyle = "font-family:Bwgrkl; font-size: 6pt;"
      Case "WTT"
        sBibleTextFontStyle = "font-family:Bwhebl; font-size: 6pt;"
      Case Else
        sBibleTextFontStyle = sStandardFontStyle
    End Select
    
    '---Save scroll position
    nVertical = ActiveWindow.ActivePane.VerticalPercentScrolled
    nHorizontal = ActiveWindow.ActivePane.HorizontalPercentScrolled
    '---Add & format endnote
    sPass = sPassIn
    If (vPreviewLimit <> "unlimited") Then
      If (vPreviewLimit = "5000") Then vPreviewLimit = 4950
      BX_TrimToLength sPass, CInt(vPreviewLimit) - Len(sAfterPass) - Len(sBeforePass)
    End If
    BX_ConvertTextIntoHTML sBeforePass, True
    BX_ConvertTextIntoHTML sPass, True
    BX_ConvertTextIntoHTML sAfterPass, True
    
    sHTML = "<font style='" & sStandardFontStyle & "'>" & sBeforePass & "</font>"
    sHTML = sHTML & "<font style='" & sBibleTextFontStyle & "'>" & sPass & "</font>"
    sHTML = sHTML & "<font style='" & sStandardFontStyle & "'>" & sAfterPass & "</font>"
    
    'Selection.EndnoteOptions.Location = wdEndOfDocument
    Set oEndnote = Selection.Endnotes.Add(Selection.Range, ChrW(&H203A&))
    Selection.MoveLeft Unit:=wdWord, Count:=1, Extend:=wdExtend
    Selection.Font.Size = 6
    Selection.Start = Selection.End
    bx_oClip.SetHTML sHTML
    oEndnote.Range.Paste
    oLink.Range.Select
    Selection.End = Selection.End + 1
    '---Restore scroll position (unfortunately still  jumping in long documents, due to scroll positions 1-100 in long
    ActiveWindow.ActivePane.VerticalPercentScrolled = nVertical
    ActiveWindow.ActivePane.HorizontalPercentScrolled = nHorizontal
    ActiveWindow.ScrollIntoView oLink.Range
  ElseIf (m_eOptions = BX_TEXT) Then
    sPass = sPassIn & sAfterPass
    BX_ConvertTextIntoHTML sPass
    bx_oClip.SetHTML "&nbsp;&#x2013;&nbsp;" & sPass
    Selection.Start = Selection.End
    Selection.Paste
  End If
  
  Application.ScreenUpdating = True
  
  Set BX_FillBlink = oLink
End Function

Private Function AddHyperlink(ByVal sAddress As String, ByVal sSubAddress As String, ByVal sScreenTip As String, ByVal sTextToDisplay As String, ByVal sTarget As String) As Object
  On Error Resume Next
  Dim oLink As Object
  Set oLink = Nothing
  
  ' Due to a bug in Word 2007 (and possibly others?), the add hyperlink function call crashes, if TextToDisplay is included
  If (Application.Version <> "12.0") Then
    Set oLink = Selection.Hyperlinks.Add(Anchor:=Selection.Range, Address:=sAddress, SubAddress:=sSubAddress, ScreenTip:=sScreenTip, Target:=sTarget, TextToDisplay:=sTextToDisplay)
  Else
    Set oLink = Selection.Hyperlinks.Add(Anchor:=Selection.Range, Address:=sAddress, SubAddress:=sSubAddress, ScreenTip:=sScreenTip, Target:=sTarget)
    If (Not oLink Is Nothing) Then oLink.TextToDisplay = sTextToDisplay
  End If
  
  If (Err.Number > 0) Then
    Err.Clear
    If (oLink Is Nothing) Then Set oLink = Selection.Hyperlinks.Add(Anchor:=Selection.Range, Address:=sAddress, SubAddress:=sSubAddress, ScreenTip:=sScreenTip, Target:=sTarget)
    If (Not oLink Is Nothing) Then oLink.TextToDisplay = sTextToDisplay
  End If
  
  Set AddHyperlink = oLink
End Function

Private Sub BX_TrimToLength(ByRef sText As String, ByVal nLength As Long)
  bx_sFunction = "BX_TrimToLength"
  Dim nI As Long
  Dim nJ As Long
  
  If (Len(sText) > nLength + BX_CountInStr(sText, "@") + BX_CountInStr(sText, "#")) Then
    nJ = 0
    For nI = 1 To Len(sText)
      If (Mid(sText, nI, 1) <> "#" And Mid(sText, nI, 1) <> "@") Then nJ = nJ + 1
      If (nJ = nLength - 1) Then Exit For
    Next
    If (nI < Len(sText)) Then sText = Left(sText, nI) & ChrW(&H2026&) '&H2026& is horizontal ellipsis: …
  End If
End Sub


Private Function BX_ExpandReference() As Boolean
  Dim bValid As Boolean
  Dim bTemp As Boolean
  Dim bValidBooknumber As Boolean
  Dim nI As Long
  Dim nJ As Long
  Dim nK As Long
  Dim nCount As Long
  'Dim nMoveRight As Long
  Dim nSuperScript As Long
  
  bValid = True
  nSuperScript = Selection.Font.Superscript
        
  '--Extend to left from final number
  '  (Possible references: Jn 1-3, Jn 1:2a-45b, Jn 1:2a-3:45b)
  BX_ExpandSelectionType BX_LEFT, BX_NUMBER  'further end verse numbers (4)
  If (BX_ExpandSelectionType(BX_LEFT, BX_CH_VS_SEPARATOR, 1) > 0) Then 'end chapter-verse separator (:)
    If (BX_ExpandSelectionType(BX_LEFT, BX_NUMBER) = 0) Then  'end chapter numbers (3)
      Selection.Start = Selection.Start + 1
    End If
  End If
  If (BX_ExpandSelectionType(BX_LEFT, BX_DASH, 1) > 0) Then   'single verse dash (-)
    BX_ExpandSelectionType BX_LEFT, BX_SENTENCE_LETTER, 1  'start verse sentence letter (a)
    If (BX_ExpandSelectionType(BX_LEFT, BX_NUMBER) > 0) Then 'start verse numbers (2)
      If (BX_ExpandSelectionType(BX_LEFT, BX_CH_VS_SEPARATOR, 1) > 0) Then 'start chapter-verse separator (:)
        If (BX_ExpandSelectionType(BX_LEFT, BX_NUMBER) = 0) Then  'start chapter numbers
          Selection.Start = Selection.Start + 1
        End If
      End If
    Else
      Selection.Start = Selection.Start + 1
    End If
  End If
  
  '--Extend to left through spaces
  BX_ExpandSelectionType BX_LEFT, BX_SPACE
  '--Extend to left through 1 full stop (often before bookname)
  BX_ExpandSelectionString ".", BX_LEFT, 1
  '--Extend to left for book names with spaces (e.g. "Song of Songs")
  bTemp = False
  For nI = 1 To bx_oCompoundBooks.Size()
    bTemp = BX_ExpandSelectionString(bx_oCompoundBooks.data(nI), BX_LEFT)
    If (bTemp) Then Exit For
  Next
  If (Not bTemp) Then
    '--Extend to left through letters
    BX_ExpandSelectionType BX_LEFT, BX_LETTER
    '--Extend to left through 0-2 spaces & 0 or 1 full stop & book number
    nI = BX_ExpandSelectionType(BX_LEFT, BX_SPACE, 2)
    If (BX_ExpandSelectionString(".", BX_LEFT, 1)) Then
      nJ = 1
    Else
      nJ = 0
    End If
    nK = BX_ExpandSelectionType(BX_LEFT, BX_BOOK_NUMBER)
    bValidBooknumber = True
    Select Case nK
    Case 0
        bValidBooknumber = False
    Case 2
      If (LCase(Left(Selection.Text, 2)) <> "ii") Then
        bValidBooknumber = False
      End If
    Case 3
      If (LCase(Left(Selection.Text, 3)) <> "iii") Then
        bValidBooknumber = False
      End If
    Case Is > 3
        bValidBooknumber = False
    End Select
    If (Not bValidBooknumber) Then
      Selection.Start = Selection.Start + nI + nJ + nK
    End If
  End If
  Do While (BX_TestChar(Left(Selection.Text, 1), BX_SPACE))
   Selection.Start = Selection.Start + 1
  Loop
  '--Extend to right through f or ff or single sentence letter
  If (bValid) Then
    If (Not BX_ExpandSelectionString("ff", BX_RIGHT, 1)) Then _
      If (Not BX_ExpandSelectionString("f", BX_RIGHT, 1)) Then _
        BX_ExpandSelectionType BX_RIGHT, BX_SENTENCE_LETTER, 1
  End If
  '--Check for numbers or other letters directly to right
  If (bValid) Then
    If (BX_ExpandSelectionType(BX_RIGHT, BX_LETTER Or BX_NUMBER) > 0) Then bValid = False
  End If

  '--Check for superscript
  If (Selection.Font.Superscript <> nSuperScript) Then bValid = False
  
  BX_ExpandReference = bValid
End Function

Private Function BX_ExpandSelectionString(ByVal sStr As String, ByVal eDir As BX_Direction, Optional ByVal bEitherCase = True) As Boolean
  bx_sFunction = "BX_ExpandSelectionString"
  Dim nLen As Long
  Dim bOK As Boolean
  Dim oOldRange As Range
  Dim nCompare As Integer
  Dim nOldStart As Long
  Dim nOldEnd As Long
  
  bOK = False
  nLen = Len(sStr)
  Set oOldRange = Selection.Range
  If (bEitherCase) Then nCompare = 1 Else nCompare = 0
  
  If (eDir = BX_LEFT) Then
    If (Selection.Start > nLen) Then
      Selection.Start = Selection.Start - nLen
      If (StrComp(Replace(Left(Selection.Text, nLen), ChrW(&H2008&), " "), sStr, nCompare) = 0 And Selection.Hyperlinks.Count() = 0) Then
        bOK = True
        If (Selection.Start > 1) Then
          nOldStart = Selection.Start
          nOldEnd = Selection.End
          Selection.Start = Selection.Start - 1
          If (BX_TestChar(Left(sStr, 1), BX_LETTER) And BX_TestChar(Left(Selection.Text, 1), BX_LETTER)) Then bOK = False 'do not allow expansion into a word
          Selection.Start = Selection.Start + 1
          Selection.Start = nOldStart
          Selection.End = nOldEnd 'Might have been moved if within a table
        End If
      End If
    End If
  Else
    If (Selection.End + nLen < Selection.StoryLength) Then
      Selection.End = Selection.End + nLen
      If (StrComp(Replace(Right(Selection.Text, nLen), ChrW(&H2008&), " "), sStr, nCompare) = 0 And Selection.Hyperlinks.Count() = 0) Then
        bOK = True
        If (Selection.End + 1 < Selection.StoryLength) Then
          nOldStart = Selection.Start
          nOldEnd = Selection.End
          Selection.End = Selection.End + 1
          If (BX_TestChar(Right(sStr, 1), BX_LETTER) And BX_TestChar(Right(Selection.Text, 1), BX_LETTER)) Then bOK = False 'do not allow expansion into a word
          Selection.End = nOldEnd
          Selection.Start = nOldStart 'Might have been moved if within a table
        End If
      End If
    End If
  End If
  
  If (Not bOK) Then oOldRange.Select

  BX_ExpandSelectionString = bOK
End Function

Private Function BX_ExpandSelectionType(ByVal eDir As BX_Direction, ByVal nCharTypes As Integer, Optional ByVal nMaxOccurences As Integer = 0) As Long
  bx_sFunction = "BX_ExpandSelectionType"
  Dim nExpanded As Long
  Dim oOldRange As Range
  Dim oLastValidRange As Range
  
  nExpanded = 0
  Set oOldRange = Selection.Range
  Set oLastValidRange = Selection.Range
  
  If (Selection.Hyperlinks.Count = 0) Then
    If (eDir = BX_LEFT) Then
      Do While (Selection.Start > 0 And (nMaxOccurences = 0 Or nExpanded <= nMaxOccurences))
        Selection.Start = Selection.Start - 1
        If (Selection.Tables.Count > 0) Then
          If (BX_CountInStr(Selection.Text, ChrW(7)) > 0) Then
            oLastValidRange.Select
            Exit Do
          End If
        End If
        If (Selection.Hyperlinks.Count <> 0) Then
          Selection.Start = Selection.Hyperlinks(1).Range.End
          Exit Do
        ElseIf (Not BX_TestChar(Left(Selection.Text, 1), nCharTypes)) Then
          Selection.Start = Selection.Start + 1
          Exit Do
        Else
          nExpanded = nExpanded + 1
          Set oLastValidRange = Selection.Range
        End If
      Loop
    Else
      Do While (Selection.End < Selection.StoryLength And (nMaxOccurences = 0 Or nExpanded <= nMaxOccurences))
        Selection.End = Selection.End + 1
        If (Selection.Tables.Count > 0) Then
          If (BX_CountInStr(Selection.Text, ChrW(7)) > 0) Then
            oLastValidRange.Select
            Exit Do
          End If
        End If
        If (Selection.Hyperlinks.Count <> 0) Then
          Selection.End = Selection.Hyperlinks(1).Range.Start
          Exit Do
        ElseIf (Not BX_TestChar(Right(Selection.Text, 1), nCharTypes)) Then
          Selection.End = Selection.End - 1
          Exit Do
        Else
          nExpanded = nExpanded + 1
          Set oLastValidRange = Selection.Range
        End If
      Loop
    End If
  End If
  
  If (nMaxOccurences <> 0 And nExpanded > nMaxOccurences) Then
    oOldRange.Select
    nExpanded = 0
  End If
  
  BX_ExpandSelectionType = nExpanded
End Function

Private Function BX_ExpandSelectedPartialReference() As Boolean
  bx_sFunction = "BX_ExpandSelectedPartialReference"
  Dim bExtended As Boolean
  Dim bSkip As Boolean
  
  bSkip = False
  bExtended = False
  
''---Check for expanding selection to the left - chapters
'  If (Not bSkip And Not bExtended) Then
'    bExtended = True
'    If (Not BX_ExpandSelectionString("chapters ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("chapters", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("chapter ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("chapter", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("chaps. ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("chaps.", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("chaps ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("chaps", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("chap. ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("chap.", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("chap ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("chap", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("chs. ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("chs.", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("chs ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("chs", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("ch. ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("ch.", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("ch ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("ch", BX_LEFT)) Then _
'      bExtended = False
''    bChaptersOnly = bExtended
'  End If
'
''---Check for expanding selection to the left - verses
'  If (Not bSkip And Not bExtended) Then
'    bExtended = True
'    If (Not BX_ExpandSelectionString("verses ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("verses", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("verse ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("verse", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("vv. ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("vv.", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("vv ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("vv", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("vss. ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("vss.", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("vss ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("vss", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("vs. ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("vs.", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("vs ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("vs", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("v. ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("v.", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("v ", BX_LEFT)) Then _
'    If (Not BX_ExpandSelectionString("v", BX_LEFT)) Then _
'      bExtended = False
'  End If

'---Check for unwanted expansion to the left (p: pages, c: circa (very rarely chapter))
  If (Not bSkip And Not bExtended) Then
    bExtended = True
    If (Not BX_ExpandSelectionString("pp. ", BX_LEFT)) Then _
    If (Not BX_ExpandSelectionString("pp.", BX_LEFT)) Then _
    If (Not BX_ExpandSelectionString("pp", BX_LEFT)) Then _
    If (Not BX_ExpandSelectionString("p. ", BX_LEFT)) Then _
    If (Not BX_ExpandSelectionString("p.", BX_LEFT)) Then _
    If (Not BX_ExpandSelectionString("p", BX_LEFT)) Then _
    If (Not BX_ExpandSelectionString("c.", BX_LEFT)) Then _
    If (Not BX_ExpandSelectionString("c ", BX_LEFT)) Then _
    If (Not BX_ExpandSelectionString("AD ", BX_LEFT)) Then _
    If (Not BX_ExpandSelectionString("AD", BX_LEFT)) Then _
    If (Not BX_ExpandSelectionString("A.D. ", BX_LEFT)) Then _
    If (Not BX_ExpandSelectionString("A.D.", BX_LEFT)) Then _
      bExtended = False
    bSkip = bExtended
  End If
  
'---Check for unwanted expansion to the right
  If (Not bSkip And Not bExtended) Then
    bExtended = True
    If (Not BX_ExpandSelectionString(" AD", BX_RIGHT)) Then _
    If (Not BX_ExpandSelectionString("AD", BX_RIGHT)) Then _
    If (Not BX_ExpandSelectionString(" A.D.", BX_RIGHT)) Then _
    If (Not BX_ExpandSelectionString("A.D.", BX_RIGHT)) Then _
    If (Not BX_ExpandSelectionString(" CE", BX_RIGHT)) Then _
    If (Not BX_ExpandSelectionString("CE", BX_RIGHT)) Then _
    If (Not BX_ExpandSelectionString(" BCE", BX_RIGHT)) Then _
    If (Not BX_ExpandSelectionString("BCE", BX_RIGHT)) Then _
    If (Not BX_ExpandSelectionString(" BC", BX_RIGHT)) Then _
    If (Not BX_ExpandSelectionString("BC", BX_RIGHT)) Then _
    If (Not BX_ExpandSelectionString(" B.C.", BX_RIGHT)) Then _
    If (Not BX_ExpandSelectionString("B.C.", BX_RIGHT)) Then _
      bExtended = False
    bSkip = bExtended
  End If
  
'---Check for expanding selection to the right
  bExtended = False
  If (Not bSkip And Not bExtended) Then
    If (Not BX_ExpandSelectionString("ff", BX_RIGHT)) Then _
    BX_ExpandSelectionString "f", BX_RIGHT
  End If

  BX_ExpandSelectedPartialReference = Not bSkip
End Function
          

Private Sub BX_ConvertTextIntoHTML(ByRef sText As String, Optional ByVal bEndnoteFormat = False)
  bx_sFunction = "BX_ConvertTextIntoHTML"
  Dim nP1 As Long
  Dim nP2 As Long
  Dim nI As Long
  
 '---Remove single chapter number
   If (BX_CountInStr(sText, "@@") = 2) Then
     sText = Split(sText, "@@")(0) & Split(sText, "@@")(2)
   End If
 '---Remove single verse number
   If (BX_CountInStr(sText, "##") = 2) Then
     sText = Split(sText, "##")(0) & Split(sText, "##")(2)
   End If
 '---Replace chapter numbers
  nP1 = -1
  Do
    nP1 = InStr(nP1 + 2, sText, "@@")
    nP2 = InStr(nP1 + 2, sText, "@@")
    If (nP1 > 0) Then
      If (nP2 > 0) Then
        If (bEndnoteFormat) Then
          sText = Left(sText, nP1 - 1) & Mid(sText, nP1 + 2, nP2 - nP1 - 2) & "&nbsp;" & Right(sText, Len(sText) - nP2 - 1)
        Else
          sText = Left(sText, nP1 - 1) & "<b>" & Mid(sText, nP1 + 2, nP2 - nP1 - 2) & "</b>&nbsp;" & Right(sText, Len(sText) - nP2 - 1)
        End If
      Else
        nP2 = InStr(nP1, sText, ChrW(&H2026&))
        If (nP2 > 0) Then
          sText = Left(sText, nP1 - 1) & Mid(sText, nP2)
        Else
          sText = Left(sText, nP1 - 1)
        End If
      End If
    End If
  Loop While (nP1 <> 0)
  sText = Replace(sText, "@" & ChrW(&H2026&), ChrW(&H2026&))
 '---Replace verse numbers
  nP1 = -1
  Do
    nP1 = InStr(nP1 + 2, sText, "##")
    nP2 = InStr(nP1 + 2, sText, "##")
    If (nP1 > 0) Then
      If (nP2 > 0) Then
        If (bEndnoteFormat) Then
          sText = Left(sText, nP1 - 1) & BX_Superscript(Mid(sText, nP1 + 2, nP2 - nP1 - 2), True) & Right(sText, Len(sText) - nP2 - 1)
        Else
          sText = Left(sText, nP1 - 1) & "<b><sup>" & Mid(sText, nP1 + 2, nP2 - nP1 - 2) & "</sup></b>" & Right(sText, Len(sText) - nP2 - 1)
        End If
      Else
        nP2 = InStr(nP1, sText, ChrW(&H2026&))
        If (nP2 > 0) Then
          sText = Left(sText, nP1 - 1) & Mid(sText, nP2)
        Else
          sText = Left(sText, nP1 - 1)
        End If
      End If
    End If
  Loop While (nP1 <> 0)
  sText = Replace(sText, "#" & ChrW(&H2026&), ChrW(&H2026&))
 '---Convert all characters above ASCII 127
  For nI = 1 To Len(sText)
    If (AscW(Mid(sText, nI, 1)) > 127) Then
      sText = Left(sText, nI - 1) & "&#" & AscW(Mid(sText, nI, 1)) & ";" & Mid(sText, nI + 1)
    End If
  Next
  'sText = Replace(sText, ChrW(&H2013&), "&#x2013;") 'en dash
  'sText = Replace(sText, ChrW(&H2014&), "&#x2013;") 'em dash
  'sText = Replace(sText, ChrW(&H2026&), "&#x2026;") 'horizontal ellipsis
End Sub

Private Sub BX_UnlinkSelection()
  bx_sFunction = "BX_UnlinkSelection"
  Dim oOldRange As Range
  Dim oRange As Range
  Dim oHyperlink As Hyperlink
  Dim nLastPos As Long
  Dim bOK As Boolean

  Set oOldRange = Selection.Range
  Set oRange = Selection.Range
  nLastPos = 0
  bOK = True
  Selection.Start = Selection.End
  Application.Browser.Target = wdBrowseField
  If (Selection.Hyperlinks.Count = 0) Then Application.Browser.Previous
  Do
    If (bOK) Then
      nLastPos = Selection.Start
      Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
      If (Selection.Hyperlinks.Count = 1) Then
        Set oHyperlink = Selection.Hyperlinks(1)
        Selection.Move wdCharacter, 1
        oHyperlink.Delete
        If (BX_ExpandSelectionString(ChrW(&H203A&), BX_RIGHT, 1)) Then
          If (Selection.Endnotes.Count() > 0) Then
            Selection.Delete
          End If
        End If
      End If
    End If
    Application.Browser.Previous
    bOK = (Selection.Start <> nLastPos And Selection.Start >= oRange.Start And Selection.End <= oRange.End) 'checking of ended needed if initial selection was not in main text
  Loop While (bOK)
  
  oOldRange.Select
End Sub



Sub BX_HandleFindParameters(ByVal bBefore As Boolean)
' Of all possible format options, only the font format is currently preserved.
  bx_sFunction = "BX_HandleFindParameters"
  Static vText As Variant
  Static vForward As Variant
  Static vWrap As Variant
  Static vFormat As Variant
  Static vLanguageID As Variant
  Static vLanguageIDFarEast As Variant
  Static vLanguageIDOther As Variant
  Static bMatchCase As Boolean
  Static bMatchWholeWord As Boolean
  Static bMatchKashida As Boolean
  Static bMatchDiacritics As Boolean
  Static bMatchAlefHamza As Boolean
  Static bMatchControl As Boolean
  Static bMatchWildcards As Boolean
  Static bMatchSoundsLike As Boolean
  Static bMatchAllWordForms As Boolean
  'Static sStyle As String
  Static oFont As Font
  'Static oParagraphFormat As ParagraphFormat
  'Static oFrame As Frame
  
  If (bBefore) Then
    If (vFormat) Then
      If (oFont Is Nothing) Then Set oFont = New Font
      BX_CloneFont Selection.Find.Font, oFont, True
      vLanguageID = Selection.Find.LanguageID
      vLanguageIDFarEast = Selection.Find.LanguageIDFarEast
      vLanguageIDOther = Selection.Find.LanguageIDOther
      'If (oParagraphFormat Is Nothing) Then Set oParagraphFormat = New ParagraphFormat
      'If (oFrame Is Nothing) Then Set oFrame = New Frame
      'BX_CloneParagraphFormat Selection.Find.ParagraphFormat, oParagraphFormat
      'BX_CloneFrame Selection.Find.Frame, oFrame
      'sStyle = CStr(Selection.Find.Style) 'There seems to be a bug in Word here: this field remains set to a style when it was used in an earlier search, even if the current search does not use it.
    End If
    With Selection.Find
      vText = .Text
      vForward = .Forward
      vWrap = .Wrap
      vFormat = .Format
      bMatchCase = .MatchCase
      bMatchWholeWord = .MatchWholeWord
      bMatchKashida = .MatchKashida
      bMatchDiacritics = .MatchDiacritics
      bMatchAlefHamza = .MatchAlefHamza
      bMatchControl = .MatchControl
      bMatchWildcards = .MatchWildcards
      bMatchSoundsLike = .MatchSoundsLike
      bMatchAllWordForms = .MatchAllWordForms
      
      .MatchCase = False
      .MatchWholeWord = False
      .MatchKashida = False
      .MatchDiacritics = False
      .MatchAlefHamza = False
      .MatchControl = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      .ClearFormatting
      .Wrap = wdFindStop
    End With
  Else
    With Selection.Find
      .ClearFormatting
      .Text = vText
      .Forward = vForward
      .Wrap = vWrap
      .Format = vFormat
      .MatchCase = bMatchCase
      .MatchWholeWord = bMatchWholeWord
      .MatchKashida = bMatchKashida
      .MatchDiacritics = bMatchDiacritics
      .MatchAlefHamza = bMatchAlefHamza
      .MatchControl = bMatchControl
      .MatchWildcards = bMatchWildcards
      .MatchSoundsLike = bMatchSoundsLike
      .MatchAllWordForms = bMatchAllWordForms
    End With
    If (vFormat) Then
      BX_CloneFont oFont, Selection.Find.Font, True
      Selection.Find.LanguageID = vLanguageID
      Selection.Find.LanguageIDFarEast = vLanguageIDFarEast
      Selection.Find.LanguageIDOther = vLanguageIDOther
      'Selection.Find.Style = sStyle
      'BX_CloneParagraphFormat oParagraphFormat, Selection.Find.ParagraphFormat
      'BX_CloneFrame oFrame, Selection.Find.Frame
    End If
  End If
End Sub
