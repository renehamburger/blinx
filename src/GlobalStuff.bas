Attribute VB_Name = "GlobalStuff"
Option Explicit
Option Base 1

'==============================================================================
' See https://github.com/renehamburger/blinx for source code, manual & license
'==============================================================================

'------------------------------------------------------------------------------
' settings

Public Const BX_LANGUAGES = "English#English|e.g., John 3:16,18#German|e.g., Johannes 3,16.18"
'Public Const BX_TRANSLATIONS = "ESV#ESV|English Standard Version#NIV|New International Version#NIB|New International Version (UK)#TNIV|Today's New International Version#NIRV|New International Reader's Version#NASB|New American Standard Bible#KJV|King James Version#NLT|New Living Translation#YLT|Young's Literal Translation"
Public Const BX_TRANSLATIONS = "ESV#ESV|English Standard Version#NIV|New International Version#NIB|New International Version (UK)#TNIV|Today's New International Version#NIRV|New International Reader's Version#NASB|New American Standard Bible#KJV|King James Version" & _
                "#NLT|New Living Translation#YLT|Young's Literal Translation#BGT|BW LXX/Greek NT#ELB|Elberfelder#ZUR|Zürcher#SCL|Schlachter 2000#EIN|Einheitsübersetzung#LUO|Luther 1912#NEU|Neue evangelistische Übersetzung"  'WTT|BW Hebrew OT# leads to crash...
Public Const BX_ONLINE_BIBLES = "esvbible.org#esvbible.org|(ESV with commentary)#biblegateway.com|(All major English Bible versions)#bibleserver.com|(Ideal for German Bible versions)"
Public Const BX_TEXT_SOURCES = "biblegateway.com#biblegateway.com#Logos#BibleWorks"
Public Const BX_BLINK_PREVIEW_LENGTHS = "5000#100#200#500#1000#2000#5000#unlimited"
Public Const BX_VERSION = "0.11"
Public Const BX_VERSION_FULL = "v0.11.1 (24 April 2021)"
Public Const BX_MAX_CHAPTER = 152
Public Const BX_MAX_VERSE = 176
Public Const BX_MAX_NUMBER = 176
'Public Const BX_BOOK_NAMES_EN = "Genesis|Exodus|Leviticus|Numbers|Deuteronomy|Joshua|Judges|Ruth|1 Samuel|2 Samuel|1 Kings|2 Kings|1 Chronicles|2 Chronicles|Ezra|Nehemiah|Esther" & _
'                "|Job|Psalm|Proverbs|Ecclesiastes|Song|Isaiah|Jeremiah|Lamentations|Ezekiel|Daniel|Hosea|Joel|Amos|Obadiah|Jonah|Micah|Nahum|Habakkuk|Zephaniah|Haggai|Zechariah|Malachi" & _
'                "|Matthew|Mark|Luke|John|Acts|Romans|1 Corinthians|2 Corinthians|Galatians|Ephesians|Philippians|Colossians|1 Thessalonians|2 Thessalonians|1 Timothy|2 Timothy|Titus|Philemon|Hebrews|James|1 Peter|2 Peter|1 John|2 John|3 John|Jude|Revelation"
Public Const BX_BOOK_NAMES_EN = "Gen|Exod|Lev|Numb|Deut|Josh|Judg|Rut|1 Sam|2 Sam|1 King|2 King|1 Chron|2 Chron|Ezr|Neh|Est|Jb" & _
                "|Psa|Prov|Eccl|Song of Solomon|Isa|Jer|Lam|Ezek|Dan|Hos|Joe|Amo|Obad|Jon|Mic|Nah|Hab|Zeph|Hagg|Zech|Mal" & _
                "|Matth|Mar|Luk|Joh|Act|Rom|1 Cor|2 Cor|Gal|Eph|Phil|Col|1 Thess|2 Thess|1 Tim|2 Tim|Tit|Phlm|Hebr|Jam|1 Pet|2 Pet|1 Jo|2 Jo|3 Jo|Jud|Rev"
Public Const BX_DEFAULT_BOOK_NAMES_EN = _
                "Genesis|Gen|Gn#Exodus|Exod|Exo|Ex#Leviticus|Lev|Lv#Numbers|Numb|Num|Nu#Deuteronomy|Deut|Deu|Dt#Joshua|Josh|Jos|Jsh#Judges|Judg|Jud|Jdg" & _
                "#Ruth|Rut|Ru|Rt#1 Samuel|1 Sam|1 Sa|1 S|1 Sm#2 Samuel|2 Sam|2 Sa|2 S|2 Sm#1 Kings|1 King|1 Kin|1 Ki|1 Kgs|1 Kg|1 K#2 Kings|2 King|2 Kin|2 Ki|2 Kgs|2 Kg|2 K" & _
                "#1 Chronicles|1 Chron|1 Chro|1 Chr|1 Ch#2 Chronicles|2 Chron|2 Chro|2 Chr|2 Ch#Ezra|Ezr#Nehemiah|Neh|Ne#Esther|Est|Es#Job|Jb#Psalm|Psa|Ps|Psl|Psm" & _
                "#Proverbs|Prov|Pro|Prv|Pr#Ecclesiastes|Eccl|Ecc#Song|Song of Solomon|Song of Songs|Songs|Sol|Cant#Isaiah|Isa|Is#Jeremiah|Jer|Je#Lamentations|Lam|La" & _
                "#Ezekiel|Ezek|Eze|Ezk#Daniel|Dan|Dn#Hosea|Hos|Ho#Joel|Joe|Jo#Amos|Amo|Am#Obadiah|Obad|Oba|Ob|Obd#Jonah|Jon|Jnh#Micah|Mic|Mi|Mc#Nahum|Nah|Na#Habakkuk|Hab|Hb" & _
                "#Zephaniah|Zeph|Zep|Zp|Zph#Haggai|Hagg|Hag|Hg#Zechariah|Zech|Zec|Zch|Zc#Malachi|Mal|Ml" & _
                "#Matthew|Matth|Matt|Mat|Mt#Mark|Mar|Mk#Luke|Luk|Lu|Lk#John|Joh|Jn|Jh#Acts|Act|Ac#Romans|Rom|Ro|Rm|Rms#1 Corinthians|1 Cor|1 Co|1 C#2 Corinthians|2 Cor|2 Co|2 C" & _
                "#Galatians|Gal|Ga#Ephesians|Eph#Philippians|Phil|Phi|Phl#Colossians|Col|Co#1 Thessalonians|1 Thess|1 Thes|1 Th#2 Thessalonians|2 Thess|2 Thes|2 Th#1 Timothy|1 Tim|1 Ti|1 Tm" & _
                "#2 Timothy|2 Tim|2 Ti|2 Tm#Titus|Tit|Tt#Philemon|Phlm|Phm#Hebrews|Hebr|Heb|Hb#James|Jam|Jas|Js#1 Peter|1 Pet|1 Pe|1 P|1 Pt#2 Peter|2 Pet|2 Pe|2 P|2 Pt#1 John|1 Jo|1 Jn|1 J" & _
                "#2 John|2 Jo|2 Jn|2 J#3 John|3 Jo|3 Jn|3 J#Jude|Jud|Jd#Revelation|Rev|Rvl|Rv"
Public Const BX_DEFAULT_BOOK_NAMES_DE = _
                "1 Mose|1 Mos|1 Mo|Genesis|Gen|Gn#2 Mose|2 Mos|2 Mo|Exodus|Exod|Ex#3 Mose|3 Mos|3 Mo|Levitikus|Lev|Lv#4 Mose|4 Mos|4 Mo|Numeri|Num|Nm#5 Mose|5 Mos|5 Mo|Deuteronomium|Deut|Dtn|Dt" & _
                "#Josua|Jos#Richter|Rich|Ri#Rut|Ru#1 Samuel|1 Sam|1 Sa#2 Samuel|2 Sam|2 Sa#1 Könige|1 Kön|1 Kng|1 Kö#2 Könige|2 Kön|2 Kng|2 Kö#1 Chronik|1 Chron|1 Chr#2 Chronik|2 Chron|2 Chr" & _
                "#Esra|Esr#Nehemia|Neh#Ester|Est#Hiob|Ijob|Job#Psalmen|Psalm|Pslm|Ps#Sprüche|Sprichwörter|Spr#Prediger|Pred|Kohelet|Koh|Ecclesiastes#Hohelied|Hoheslied|Hld#Jesaja|Jes#Jeremia|Jer#Klagelieder|Klgl" & _
                "#Hesekiel|Hesek|Hes|Ezechiel|Ez#Daniel|Dan#Hosea|Hos#Joel#Amos#Obadja|Obad|Obd#Jona#Micha|Mich#Nahum|Nah#Habakuk|Hab#Zefanja|Zef#Haggai|Hagg|Hag#Sacharja|Sach#Maleachi|Mal" & _
                "#Matthäus|Matt|Mt#Markus|Mark|Mk#Lukas|Luk|Lk#Johannes|Joh#Apostelgeschichte|Apg#Römer|Röm|Rm#1 Korinther|1 Kor|1 Ko#2 Korinther|2 Kor|2 Ko#Galater|Gal" & _
                "#Epheser|Eph#Philipper|Phil#Kolosser|Kol#1 Thessalonicher|1 Thess|1 Th#2 Thessalonicher|2 Thess|2 Th#1 Timotheus|1 Tim#2 Timotheus|2 Tim#Titus|Tit" & _
                "#Philemon|Phlm|Phm#Hebräer|Hebr|Heb|Hb#Jakobus|Jak|Jk#1 Petrus|1 Petr|1 Pet|1 Pt#2 Petrus|2 Petr|2 Pet|2 Pt" & _
                "#1 Johannes|1 Joh|1 Jo#2 Johannes|2 Joh|2 Jo#3 Johannes|3 Joh|3 Jo#Judas|Jud#Offenbarung|Offb"
Public Const BX_MAX_SENTENCE_LETTER = "e"

'------------------------------------------------------------------------------
#If Win64 Then
  Public Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As LongPtr) As Long
  Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
  Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
  Public Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
  Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
  Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
  Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
  Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  Public Declare Function GetForegroundWindow Lib "user32" () As Long
  Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
  'Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
  'Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
  'Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
  'Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If
'------------------------------------------------------------------------------

Public Enum BX_Language
  BX_NO_LANGUAGE = 0
  BX_ENGLISH = 1
  BX_GERMAN = 2
End Enum

Public Enum BX_ReturnValue
  BX_UNDEFINED
  BX_FAILED
  BX_NOTHING
  BX_LINK_CREATED
  BX_ALREADY_LINK
End Enum

Public Enum BX_Direction
  BX_LEFT
  BX_RIGHT
End Enum

Public Enum BX_Options
  BX_TOOLTIP
  BX_TEXT
  BX_PURE
End Enum

Public Enum BX_CHAR_TYPE
  BX_NUMBER = 1
  BX_LETTER = 2
  BX_SPACE = 4
  BX_DASH = 8
  BX_RETURN = 16
  BX_SEPARATOR = 32
  BX_CH_VS_SEPARATOR = 64
  BX_SENTENCE_LETTER = 128
  BX_BOOK_NUMBER = 256
End Enum

Public Enum BX_BOOKS 'First 3 characters, except for Judges (JDG) and Philemon (PHM)
  BX_BK_GEN = 1
  BX_BK_EXO = 2
  BX_BK_LEV = 3
  BX_BK_NUM = 4
  BX_BK_DEU = 5
  BX_BK_JOS = 6
  BX_BK_JDG = 7
  BX_BK_RUT = 8
  BX_BK_1SA = 9
  BX_BK_2SA = 10
  BX_BK_1KI = 11
  BX_BK_2KI = 12
  BX_BK_1CH = 13
  BX_BK_2CH = 14
  BX_BK_EZR = 15
  BX_BK_NEH = 16
  BX_BK_EST = 17
  BX_BK_JOB = 18
  BX_BK_PSA = 19
  BX_BK_PRO = 20
  BX_BK_ECC = 21
  BX_BK_SON = 22
  BX_BK_ISA = 23
  BX_BK_JER = 24
  BX_BK_LAM = 25
  BX_BK_EZE = 26
  BX_BK_DAN = 27
  BX_BK_HOS = 28
  BX_BK_JOE = 29
  BX_BK_AMO = 30
  BX_BK_OBA = 31
  BX_BK_JON = 32
  BX_BK_MIC = 33
  BX_BK_NAH = 34
  BX_BK_HAB = 35
  BX_BK_ZEP = 36
  BX_BK_HAG = 37
  BX_BK_ZEC = 38
  BX_BK_MAL = 39
  BX_BK_MAT = 40
  BX_BK_MAR = 41
  BX_BK_LUK = 42
  BX_BK_JOH = 43
  BX_BK_ACT = 44
  BX_BK_ROM = 45
  BX_BK_1CO = 46
  BX_BK_2CO = 47
  BX_BK_GAL = 48
  BX_BK_EPH = 49
  BX_BK_PHI = 50
  BX_BK_COL = 51
  BX_BK_1TH = 52
  BX_BK_2TH = 53
  BX_BK_1TI = 54
  BX_BK_2TI = 55
  BX_BK_TIT = 56
  BX_BK_PHM = 57
  BX_BK_HEB = 58
  BX_BK_JAM = 59
  BX_BK_1PE = 60
  BX_BK_2PE = 61
  BX_BK_1JO = 62
  BX_BK_2JO = 63
  BX_BK_3JO = 64
  BX_BK_JUD = 65
  BX_BK_REV = 66
End Enum

Public Type BX_Reference
  sBook As String
  nBook As Integer
  nChapter1 As Integer
  nChapter2 As Integer
  nVerse1 As Integer
  nVerse2 As Integer
  nSentence1 As Integer
  nSentence2 As Integer
End Type


'------------------------------------------------------------------------------

Public bx_eLanguage As BX_Language
Public bx_oLastValidRef As BX_Reference
Public bx_oB As IBible
Public bx_oClip As New clsClipboard
Public bx_sFunction As String
Public bx_vTimeWarningOnline As Variant
Public bx_sVariablesLoaded As String
Public bx_asBooks(2, 66, 11) As String
Public bx_asChVsSeparators() As String
Public bx_asSeparators() As String
Public bx_oCompoundBooks As New clsVector
Public bx_oRefPositions As New clsVector
Public bx_oProcessInputForm As BX_ProcessInputForm
Public bx_oOptionsForm As BX_OptionsForm
Public bx_oEditBookNamesForm As BX_EditBookNamesForm
Public bx_oGeneralForm As BX_GeneralForm

'---debug:
Public bx_sCurrentDocument As String

'------------------------------------------------------------------------------
Public Sub BX_Reset()
  Set bx_oB = Nothing
  BX_LoadVariables
  BX_SaveVariables
  BX_CheckForms
End Sub

Public Sub BX_CheckForms()
  If (bx_oProcessInputForm Is Nothing) Then Set bx_oProcessInputForm = New BX_ProcessInputForm
  If (bx_oOptionsForm Is Nothing) Then Set bx_oOptionsForm = New BX_OptionsForm
  If (bx_oEditBookNamesForm Is Nothing) Then Set bx_oEditBookNamesForm = New BX_EditBookNamesForm
  If (bx_oGeneralForm Is Nothing) Then Set bx_oGeneralForm = New BX_GeneralForm
End Sub

Public Sub BX_LoadVariables()
  bx_sFunction = "BX_LoadVariables"
  Dim nI As Integer
  Dim nJ As Integer
  Dim sBooks As String
  Dim nSize As Integer
  Dim eLanguage As BX_Language
  Dim sChapterVerseSeparators As String
  Dim sSeparators As String
  Dim sCurrentLanguage As String
  
  sCurrentLanguage = GetSetting("Blinx", "Options", "Language", Split(BX_LANGUAGES, "#")(0))
   
  Select Case sCurrentLanguage
    Case "English"
      bx_eLanguage = BX_ENGLISH
      sChapterVerseSeparators = GetSetting("Blinx", "Options", "ChapterVerseSeparators.en", ":|.")
      sSeparators = GetSetting("Blinx", "Options", "Separators.en", ",|;")
    Case "German"
      bx_eLanguage = BX_GERMAN
      sChapterVerseSeparators = GetSetting("Blinx", "Options", "ChapterVerseSeparators.de", ",")
      sSeparators = GetSetting("Blinx", "Options", "Separators.de", ".|;")
  End Select
  
  For eLanguage = BX_ENGLISH To BX_GERMAN
    For nI = 1 To 66
      If eLanguage = BX_ENGLISH Then
        sBooks = GetSetting("Blinx", "Books.en", "Book" & Format(nI, "00"), Split(BX_DEFAULT_BOOK_NAMES_EN, "#")(nI - 1))
      Else
        sBooks = GetSetting("Blinx", "Books.de", "Book" & Format(nI, "00"), Split(BX_DEFAULT_BOOK_NAMES_DE, "#")(nI - 1))
      End If
      nSize = CInt(BX_CountInStr(sBooks, "|") + 1)
      If (nSize > 10) Then nSize = 10
      bx_asBooks(eLanguage, nI, 1) = nSize
      For nJ = 2 To nSize + 1
        bx_asBooks(eLanguage, nI, nJ) = Split(sBooks, "|")(nJ - 2)
        If (BX_InStr(1, bx_asBooks(eLanguage, nI, nJ), " ") > 0 And Not BX_TestChar(Left(bx_asBooks(eLanguage, nI, nJ), 1), BX_BOOK_NUMBER)) Then bx_oCompoundBooks.Add (bx_asBooks(eLanguage, nI, nJ))
      Next
    Next
  Next
  
  nSize = CInt(BX_CountInStr(sChapterVerseSeparators, "|") + 1)
  ReDim bx_asChVsSeparators(nSize)
  For nI = 1 To nSize
    bx_asChVsSeparators(nI) = Split(sChapterVerseSeparators, "|")(nI - 1)
  Next
  
  nSize = CInt(BX_CountInStr(sSeparators, "|") + 1)
  ReDim bx_asSeparators(nSize)
  For nI = 1 To nSize
    bx_asSeparators(nI) = Split(sSeparators, "|")(nI - 1)
  Next
  
  bx_sVariablesLoaded = "true"
End Sub

Public Sub BX_SaveVariables()
  bx_sFunction = "BX_SaveVariables"
  Dim nI As Integer
  Dim nJ As Integer
  Dim sString As String
  Dim nSize As Integer
  Dim eLanguage As BX_Language
  
  For eLanguage = BX_ENGLISH To BX_GERMAN
    For nI = 1 To 66
      nSize = CInt(bx_asBooks(eLanguage, nI, 1))
      sString = CStr(bx_asBooks(eLanguage, nI, 2))
      For nJ = 3 To nSize + 1
        sString = sString & "|" & bx_asBooks(eLanguage, nI, nJ)
      Next
      If eLanguage = BX_ENGLISH Then
        SaveSetting "Blinx", "Books.en", "Book" & Format(nI, "00"), sString
      Else
        SaveSetting "Blinx", "Books.de", "Book" & Format(nI, "00"), sString
      End If
    Next
  Next
  
  sString = ""
  sString = bx_asChVsSeparators(1)
  For nI = 2 To UBound(bx_asChVsSeparators)
    sString = sString & "|" & bx_asChVsSeparators(nI)
  Next
  SaveSetting "Blinx", "Options", "ChapterVerseSeparators", sString
  
  sString = ""
  sString = bx_asSeparators(1)
  For nI = 2 To UBound(bx_asSeparators)
    sString = sString & "|" & bx_asSeparators(nI)
  Next
  SaveSetting "Blinx", "Options", "Separators", sString
End Sub


Public Sub BX_FocusBW()
  bx_sFunction = "BX_FocusBW"
  #If Win64 Then
    Dim nHandle As LongPtr
    Dim nRet As LongPtr
  #Else
    Dim nHandle As Long
    Dim nRet As Long
  #End If
  Const WM_SYSCOMMAND = &H112&
  Const SC_MAXIMIZE = &HF030&
  Const SC_RESTORE = &HF120&
  Const SC_CLOSE As Long = &HF060&
  Const SC_MINIMIZE As Long = &HF020&

  nHandle = FindWindow("BWFrame", vbNullString)
  If (nHandle <> 0) Then
    nRet = BringWindowToTop(nHandle)
    SendMessage nHandle, WM_SYSCOMMAND, SC_RESTORE, 0
  End If
  DoEvents
End Sub

Public Sub BX_ToggleBrowseMode()
  bx_sFunction = "BX_ToggleBrowseMode"
  #If Win64 Then
    Dim nHandle1 As LongPtr
    Dim nHandle2 As LongPtr
  #Else
    Dim nHandle1 As Long
    Dim nHandle2 As Long
  #End If
 
  nHandle1 = GetForegroundWindow()
  nHandle2 = FindWindow("BWFrame", vbNullString)
  If (nHandle1 <> 0 And nHandle2 <> 0) Then
    BX_FocusBW
    DoEvents
    SendKeys "{F6}"
    SendKeys "b"
    SetForegroundWindow (nHandle1)
    DoEvents
  End If
End Sub

Public Sub BX_ToggleNotes()
  bx_sFunction = "BX_ToggleNotes"
  Dim nHandle1 As Long
  Dim nHandle2 As Long
 
  nHandle1 = GetForegroundWindow()
  nHandle2 = FindWindow("BWFrame", vbNullString)
  If (nHandle1 <> 0 And nHandle2 <> 0) Then
    BX_FocusBW
    DoEvents
    SendKeys "{F6}"
    SendKeys "n"
    SetForegroundWindow (nHandle1)
  End If
End Sub

Public Sub BX_ToggleStrongsNumbers()
  bx_sFunction = "BX_ToggleStrongsNumbers"
  Dim nHandle1 As Long
  Dim nHandle2 As Long
 
  nHandle1 = GetForegroundWindow()
  nHandle2 = FindWindow("BWFrame", vbNullString)
  If (nHandle1 <> 0 And nHandle2 <> 0) Then
    BX_FocusBW
    DoEvents
    SendKeys "{F6}"
    SendKeys "r"
    SetForegroundWindow (nHandle1)
  End If
End Sub

Public Function BX_GetDataFromLink(ByVal oSel As Selection, ByRef oRef As BX_Reference, Optional ByRef sVersion As Variant) As Boolean
  bx_sFunction = "BX_GetDataFromLink"
  Dim sAdd As String
  Dim bVersion As Boolean
  Dim nStart1 As Long
  Dim nStart2 As Long
  Dim nEnd2 As Long
  Dim nStart3 As Long
  Dim nEnd3 As Long
  Dim sReference As String
  Dim bReturn As Boolean
  Dim sTarget As String
  Dim asParts() As String
  
  bReturn = False
  sReference = ""
  If (Not IsMissing(sVersion)) Then sVersion = ""
  
  If (oSel.Hyperlinks.Count() = 1) Then
    sTarget = oSel.Hyperlinks(1).Target
    asParts = Split(sTarget, ";")
    '---Blinx links contain the passage as hyperlink target
    If (Left(sTarget, 14) = "BX_BibRef_0.6;") Then
        sVersion = asParts(1)
        sReference = asParts(2)
        oRef = BX_StringtToReference(sReference, False, ":")
        bReturn = True
    '---Other hyperlinks could be parsed, too
    Else
      sAdd = oSel.Hyperlinks(1).Address()
      BX_ReplaceReservedCharacters sAdd
    '---Check for biblegateway
      nStart1 = InStr(1, sAdd, "www.biblegateway.com")
      If (nStart1 > 0) Then
        nStart2 = InStr(nStart1, sAdd, "search=")
        nStart3 = InStr(nStart1, sAdd, "version=")
        If (nStart2 > 0) Then
          nEnd2 = InStr(nStart2, sAdd, "&")
          If (nStart3 > 0 And Not IsMissing(sVersion)) Then
            nEnd3 = InStr(nStart3, sAdd, "&")
            If (nEnd3 > 0) Then
              sVersion = Mid(sAdd, nStart3 + 8, nEnd3 - nStart3)
            Else
              sVersion = Mid(sAdd, nStart3 + 8, Len(sAdd) - nStart3)
            End If
          End If
          If (nEnd2 > 0) Then
            sReference = Mid(sAdd, nStart2 + 7, nEnd2 - nStart2 - 7)
          Else
            sReference = Mid(sAdd, nStart2 + 7, Len(sAdd) - nStart2)
          End If
          bReturn = True
        End If
      Else
    '---Check for esvstudybible (former website of esvonline)
        nStart1 = InStr(1, sAdd, "www.esvstudybible.org")
        If (nStart1 > 0) Then
          nStart2 = InStr(nStart1, sAdd, "search?q=")
          If (nStart2 > 0) Then
            nEnd2 = Len(sAdd)
            sReference = Mid(sAdd, nStart2 + 9, nEnd2 - nStart2 - 8)
            If (IsMissing(sVersion)) Then sVersion = "ESV"
            bReturn = True
          End If
        Else
      '---Check for esvonline with "/search"
          nStart1 = InStr(1, sAdd, "www.esvonline.org/search/")
          If (nStart1 > 0) Then
            nStart2 = nStart1 + Len("www.esvonline.org/search/")
            nEnd2 = Len(sAdd)
            sReference = Mid(sAdd, nStart2, nEnd2 - nStart2 + 1)
            If (IsMissing(sVersion)) Then sVersion = "ESV"
            bReturn = True
          Else
      '---Check for esvonline without "/search"
            nStart1 = InStr(1, sAdd, "www.esvonline.org")
            If (nStart1 > 0) Then
              nStart2 = nStart1 + Len("www.esvonline.org") + 1
              nEnd2 = Len(sAdd)
              sReference = Mid(sAdd, nStart2, nEnd2 - nStart2 + 1)
              If (IsMissing(sVersion)) Then sVersion = "ESV"
              bReturn = True
            Else
      '---Check for esvonline without "/search"
              nStart1 = InStr(1, sAdd, "www.bibleserver.com/text/")
              If (nStart1 > 0) Then
                nStart2 = nStart1 + Len("www.bibleserver.com/text/")
                nEnd2 = Len(sAdd)
                sReference = Mid(sAdd, nStart2 + 4, nEnd2 - nStart2 + 1)
                sVersion = Mid(sAdd, nStart2, nStart2 + 3)
                If (IsMissing(sVersion)) Then sVersion = "ESV"
                bReturn = True
              End If
            End If
          End If
        End If
      End If
      oRef = BX_StringtToReference(sReference, False)
    End If
  End If
  
  
  BX_GetDataFromLink = bReturn
End Function

Public Sub BX_ReplaceReservedCharacters(sStr As String)
  bx_sFunction = "BX_ReplaceReservedCharacters"
  Dim nPos As Long
  Dim sCode As String
  Dim nCode As Long
  
  Do While (InStr(1, sStr, "%") > 0 And InStr(1, sStr, "%") < Len(sStr))
    nPos = InStr(1, sStr, "%")
    sCode = Mid(sStr, nPos + 1, 2)
    nCode = CLng("&H" & sCode)
    sStr = Left(sStr, nPos - 1) & Chr(nCode) & Mid(sStr, nPos + 3)
  Loop
End Sub

Public Sub BX_CleanReference(ByRef sRef As String, Optional bAllLowerCase As Boolean = True)
  bx_sFunction = "BX_CleanReference"
  Dim nI As Integer
  Dim nNameEnd As Integer
  
  If (sRef <> "") Then
    '--For whole reference: convert spaces
    For nI = 1 To Len(sRef)
      If (BX_TestChar(Mid(sRef, nI, 1), BX_SPACE)) Then
        sRef = Left(sRef, nI - 1) & " " & Mid(sRef, nI + 1)
      ElseIf (BX_TestChar(Mid(sRef, nI, 1), BX_DASH)) Then
        sRef = Left(sRef, nI - 1) & "-" & Mid(sRef, nI + 1)
      End If
    Next
   
    '--Determine end of bookname
    ' Rule: book name must be separated from rest of reference by " " or "."
    ' It would be possible to recognise references without either (e.g. Jn3:16),
    ' but that would also allow for more false alarms (e.g. a1)
    nNameEnd = InStrRev(sRef, " ")
    If (nNameEnd = 0 And BX_CountInStr(sRef, ".") > 0) Then
      nI = 0
      Do
        nI = InStr(nI + 1, sRef, ".")
        If (nI > 0 And nI < Len(sRef)) Then
          If (BX_TestChar(Mid(sRef, nI - 1, 1), BX_LETTER) And BX_TestChar(Mid(sRef, nI + 1, 1), BX_NUMBER)) Then
            nNameEnd = nI
            nI = 0
          End If
        End If
      Loop Until (nI = 0)
    End If
     
   '--For book name: replace full stops
   For nI = 1 To nNameEnd
     If (Mid(sRef, nI, 1) = ".") Then
       sRef = Left(sRef, nI - 1) & " " & Mid(sRef, nI + 1)
     End If
   Next
   
'   '--For rest: convert chapter-verse-separator
'   For nI = nNameEnd + 1 To Len(sRef)
'     If (BX_TestChar(Mid(sRef, nI, 1), BX_CH_VS_SEPARATOR)) Then
'       sRef = Left(sRef, nI - 1) & ":" & Mid(sRef, nI + 1)
'     End If
'   Next
  
   '--Remove superfluous spaces
   Do While (InStr(sRef, "  ") > 0)
     sRef = Replace(sRef, "  ", " ")
   Loop
   sRef = Trim(sRef)
   
   '--Transform all into lower case
   If (bAllLowerCase) Then sRef = LCase(sRef)
  End If
End Sub

Public Function BX_StringtToReference(ByVal sRefIn As String, Optional ByVal bAllLowerCase As Boolean = True, Optional ByVal sChapterVerseSeparator As String = "") As BX_Reference
  bx_sFunction = "BX_StringtToReference"
  Dim sRef As String
  Dim nStart As Integer
  Dim nEnd As Integer
  Dim nColumn1 As Integer
  Dim nColumn2 As Integer
  Dim nDash As Integer
  Dim bSentence As Boolean
  Dim oRef As BX_Reference
  Dim nI As Integer
  
  oRef.sBook = ""
  oRef.nChapter1 = 0
  oRef.nChapter2 = 0
  oRef.nVerse1 = 0
  oRef.nVerse2 = 0
  oRef.nSentence1 = 0
  oRef.nSentence2 = 0
  
  If (sRefIn <> "") Then
    sRef = sRefIn
    BX_CleanReference sRef, bAllLowerCase
    '--Determine end of book name
    ' BX_CleanReference allows for full stop between name & numbers
    nStart = InStrRev(sRef, " ")
    If (nStart > 0) Then
      oRef.sBook = Left(sRef, nStart - 1)
      sRef = Mid(sRef, nStart + 1)
    Else
      oRef.sBook = ""
    End If
    '-- Determine rest of reference
    nEnd = Len(sRef)
    If (sChapterVerseSeparator = "") Then
      nColumn1 = BX_InStr(1, sRef, bx_asChVsSeparators, vbBinaryCompare)
      nColumn2 = BX_InStr(nColumn1 + 1, sRef, bx_asChVsSeparators)
    Else
      nColumn1 = BX_InStr(1, sRef, ":", vbBinaryCompare)
      nColumn2 = BX_InStr(nColumn1 + 1, sRef, ":")
    End If
    nDash = InStr(1, sRef, "-")
    
    If (nColumn1 = 0) Then            '[Jn 1] or [Jn 1-2]
      If (nDash = 0) Then             '[Jn 1]
        oRef.nChapter1 = Val(Mid(sRef, 1, nEnd))
        oRef.nChapter2 = oRef.nChapter1
      Else                            '[Jn 1-2]
        oRef.nChapter1 = Val(Mid(sRef, 1, nDash - 1))
        oRef.nChapter2 = Val(Mid(sRef, nDash + 1, nEnd - nDash))
      End If
    Else
      If (nDash = 0) Then             '[Jn 1:2]
        oRef.nChapter1 = Val(Mid(sRef, 1, nColumn1 - 1))
        oRef.nChapter2 = oRef.nChapter1
        oRef.nVerse1 = Val(Mid(sRef, nColumn1 + 1, nEnd - nColumn1))
        oRef.nVerse2 = oRef.nVerse1
      Else
        If (nColumn2 = 0) Then        '[Jn 1:2-3] or [Jn 1-2:3] (=[Jn 1:1-2:3])
          If (nColumn1 < nDash) Then  '[Jn 1:2-3]
            oRef.nChapter1 = Val(Mid(sRef, 1, nColumn1 - 1))
            oRef.nChapter2 = oRef.nChapter1
            oRef.nVerse1 = Val(Mid(sRef, nColumn1 + 1, nDash - nColumn1 - 1))
            oRef.nVerse2 = Val(Mid(sRef, nDash + 1, nEnd - nDash))
          Else                        '[Jn 1-2:3]
            oRef.nChapter1 = Val(Mid(sRef, 1, nDash - 1))
            oRef.nChapter2 = Val(Mid(sRef, nDash + 1, nColumn1 - nDash - 1))
            oRef.nVerse1 = 1
            oRef.nVerse2 = Val(Mid(sRef, nColumn1 + 1, nEnd - nColumn1))
          End If
        Else                          '[Jn 1:2-2:3]
          oRef.nChapter1 = Val(Mid(sRef, 1, nColumn1 - 1))
          oRef.nVerse1 = Val(Mid(sRef, nColumn1 + 1, nDash - nColumn1 - 1))
          oRef.nChapter2 = Val(Mid(sRef, nDash + 1, nColumn2 - nDash - 1))
          oRef.nVerse2 = Val(Mid(sRef, nColumn2 + 1, nEnd - nColumn2))
        End If
      End If
    End If
    
    '--Extract sentence letters
    bSentence = False
    If (BX_TestChar(Right(sRef, 1), BX_SENTENCE_LETTER)) Then
      oRef.nSentence1 = AscW(Right(sRef, 1)) - AscW("a") + 1
      bSentence = True
    End If
    If (nDash > 0) Then
      If (bSentence) Then oRef.nSentence2 = oRef.nSentence1
      If (BX_TestChar(Mid(sRef, nDash - 1, 1), BX_SENTENCE_LETTER)) Then
        oRef.nSentence1 = AscW(Mid(sRef, nDash - 1, 1)) - AscW("a") + 1
        bSentence = True
      End If
    End If
    
    '--Process "f" or "ff" after single chapter/verse (e.g. 1:2f = 1:2-3)
    If (Not bSentence And Right(sRef, 1) = "f" And nDash = 0) Then
      If (nColumn1 = 0) Then
        oRef.nChapter2 = oRef.nChapter1 + 1
      Else
        oRef.nVerse2 = oRef.nVerse1 + 1
      End If
    End If
    
    '--Check for chapter and verse "book" names
    If ((oRef.sBook = "" And nColumn1 = 0) Or oRef.sBook = "v" Or oRef.sBook = "vss" Or oRef.sBook = "vs" Or oRef.sBook = "vv" Or oRef.sBook = "verse" Or oRef.sBook = "verses") Then
      oRef.sBook = "" '"verse"
      oRef.nVerse1 = oRef.nChapter1
      oRef.nVerse2 = oRef.nChapter2
      oRef.nChapter1 = 0
      oRef.nChapter2 = 0
    End If
    
    If (oRef.sBook = "ch" Or oRef.sBook = "chs" Or oRef.sBook = "chap" Or oRef.sBook = "chaps" Or oRef.sBook = "chapter" Or oRef.sBook = "chapters") Then
      oRef.sBook = ""  '"chapter"
    End If
  End If
  
  BX_StringtToReference = oRef
End Function

Public Function BX_ReferenceToString(ByRef oRef As BX_Reference, Optional ByVal eLanguage As BX_Language = BX_NO_LANGUAGE, _
    Optional ByVal bIncludeSentenceLetters As Boolean = False) As String
  Dim sRef As String
  Dim oRefForBook As BX_Reference
  Dim sChVsSeparator As String
  sRef = ""
  
  If (oRef.sBook <> "invalid") Then
    If (eLanguage = BX_NO_LANGUAGE) Then
      sRef = oRef.sBook
      sChVsSeparator = ":"
    Else
      If (oRef.nBook > 0) Then
        sRef = bx_asBooks(eLanguage, oRef.nBook, 2)
      ElseIf (oRef.sBook <> "") Then
        ' oRef.sBook may have been entered by the user
        BX_CheckReference oRef.sBook & " 1", oRefForBook, bx_eLanguage
        If (oRefForBook.sBook = "invalid") Then
          ' ... or generated in English (Admittedly pretty messy!)
          BX_CheckReference oRef.sBook & " 1", oRefForBook, BX_ENGLISH
        Else
          sRef = sRef
        End If
        If (oRefForBook.sBook <> "invalid") Then
          sRef = bx_asBooks(eLanguage, oRefForBook.nBook, 2)
        End If
      End If
      If (eLanguage = bx_eLanguage) Then
        sChVsSeparator = bx_asChVsSeparators(1)
      ElseIf (eLanguage = BX_ENGLISH) Then
        sChVsSeparator = ":"
      Else
        sChVsSeparator = ","
      End If
    End If
    If (oRef.nChapter1 <> 0) Then sRef = sRef & " " & CStr(oRef.nChapter1)
    If (oRef.nVerse1 <> 0) Then
      If (oRef.nChapter1 <> 0) Then
        sRef = sRef & sChVsSeparator & CStr(oRef.nVerse1)
      Else
        sRef = sRef & " " & CStr(oRef.nVerse1)
      End If
    End If
    If (bIncludeSentenceLetters) Then sRef = sRef & ChrW(AscW("a") + oRef.nSentence1 - 1)
    If (oRef.nChapter2 <> oRef.nChapter1) Then sRef = sRef & "-" & CStr(oRef.nChapter2)
    If (oRef.nVerse2 <> oRef.nVerse1 Or (oRef.nChapter2 <> oRef.nChapter1 And oRef.nVerse1 <> 0)) Then
      If (oRef.nChapter2 <> oRef.nChapter1) Then
        sRef = sRef & sChVsSeparator & CStr(oRef.nVerse2)
      Else
        sRef = sRef & "-" & CStr(oRef.nVerse2)
      End If
    End If
    sRef = Trim(sRef)
  End If
  
  BX_ReferenceToString = sRef
End Function

Public Sub BX_CheckReference(ByVal sRef As String, ByRef oRef As BX_Reference, Optional eLanguage As BX_Language = 0)
  bx_sFunction = "BX_CheckReference"
  Dim sTemp As String
  Dim nI As Integer
  Dim nJ As Integer
  Dim bRecognized As Boolean
  bRecognized = False

  If (eLanguage = 0) Then eLanguage = bx_eLanguage

  oRef = BX_StringtToReference(sRef)
  If (oRef.sBook <> "") Then
    For nI = 1 To 66
      For nJ = 2 To bx_asBooks(eLanguage, nI, 1) + 1
        '--Compare all books (lowercase without spaces)
        If (Replace(oRef.sBook, " ", "") = Replace(LCase(bx_asBooks(eLanguage, nI, nJ)), " ", "")) Then bRecognized = True
        '--Allow Roman numbers for books with 1, 2, 3 at start
        sTemp = ""
        If (Left(bx_asBooks(eLanguage, nI, nJ), 1) = 1) Then
          sTemp = "i" & Mid(bx_asBooks(eLanguage, nI, nJ), 2)
        ElseIf (Left(bx_asBooks(eLanguage, nI, nJ), 1) = 2) Then
          sTemp = "ii" & Mid(bx_asBooks(eLanguage, nI, nJ), 2)
        ElseIf (Left(bx_asBooks(eLanguage, nI, nJ), 1) = 3) Then
          sTemp = "iii" & Mid(bx_asBooks(eLanguage, nI, nJ), 2)
        End If
        If (oRef.sBook = LCase(sTemp)) Then bRecognized = True
        If (bRecognized) Then
          oRef.sBook = bx_asBooks(eLanguage, nI, 2)
          'oRef.sBook = Split(BX_BOOK_NAMES_EN, "|")(nI - 1)
          oRef.nBook = nI
          nI = 66
          Exit For
        End If
      Next
    Next
    '---Resolve books with 1 chapter
    If (bRecognized) Then
      If (oRef.nBook = 31 Or oRef.nBook = 57 Or oRef.nBook = 63 Or oRef.nBook = 64 Or oRef.nBook = 65) Then
        If (oRef.nVerse1 = 0) Then
          oRef.nVerse1 = oRef.nChapter1
          oRef.nVerse2 = oRef.nChapter2
          oRef.nChapter1 = 1
          oRef.nChapter2 = 1
        End If
      End If
    Else
      oRef.sBook = "invalid"
    End If
  End If
End Sub

Public Sub BX_CompletePartialReference(ByRef oRef As BX_Reference)
  bx_sFunction = "BX_CompletePartialReference"
  Dim bSkip As Boolean
  bSkip = False
  
  '--Check if any of numbers too large
  If (oRef.nChapter2 > BX_MAX_CHAPTER) Then bSkip = True
  If (oRef.nVerse2 > BX_MAX_VERSE) Then bSkip = True
  
  '--Determine reference
  If (Not bSkip) Then
    If (oRef.sBook = "") Then oRef.sBook = bx_oLastValidRef.sBook
    If (oRef.nChapter1 = 0) Then
      oRef.nChapter1 = bx_oLastValidRef.nChapter2
      oRef.nChapter2 = bx_oLastValidRef.nChapter2
    End If
  Else
    oRef.sBook = "invalid"
  End If
End Sub

Public Function BX_Superscript(ByVal sText As String, Optional bAsHTML As Boolean = False) As String
  bx_sFunction = "BX_Superscript"
  Dim sRet As String
  Dim nP As Long
  
  For nP = 1 To Len(sText)
    Select Case Asc(Mid(sText, nP, 1))
      Case 48:       If (bAsHTML) Then sRet = sRet & "&#x2070;" Else sRet = sRet & ChrW(&H2070)
      Case 49:       If (bAsHTML) Then sRet = sRet & "&#xB9;" Else sRet = sRet & ChrW(&HB9)
      Case 50:       If (bAsHTML) Then sRet = sRet & "&#xB2;" Else sRet = sRet & ChrW(&HB2)
      Case 51:       If (bAsHTML) Then sRet = sRet & "&#xB3;" Else sRet = sRet & ChrW(&HB3)
      Case 52 To 57: If (bAsHTML) Then sRet = sRet & "&#x207" & CStr(4 + Asc(Mid(sText, nP, 1)) - 52) & ";" Else sRet = sRet & ChrW(&H2074 + Asc(Mid(sText, nP, 1)) - 52)
      Case Else:     sRet = sRet & Mid(sText, nP, 1)
    End Select
  Next
  BX_Superscript = sRet
End Function

Public Sub BX_CloneFont(ByVal oIn As Font, ByVal oOut As Font, Optional bFindParameters As Boolean = False)
  bx_sFunction = "BX_CloneFont"
  
  oOut.AllCaps = oIn.AllCaps
  oOut.Bold = oIn.Bold
  oOut.BoldBi = oIn.BoldBi
  oOut.Color = oIn.Color
  oOut.ColorIndex = oIn.ColorIndex
  oOut.ColorIndexBi = oIn.ColorIndexBi
  oOut.DoubleStrikeThrough = oIn.DoubleStrikeThrough
  oOut.Emboss = oIn.Emboss
  oOut.Engrave = oIn.Engrave
  oOut.Hidden = oIn.Hidden
  oOut.Italic = oIn.Italic
  oOut.ItalicBi = oIn.ItalicBi
  oOut.Kerning = oIn.Kerning
  oOut.Name = oIn.Name
  oOut.NameAscii = oIn.NameAscii
  oOut.NameBi = oIn.NameBi
  oOut.NameOther = oIn.NameOther
  oOut.Outline = oIn.Outline
  oOut.Position = oIn.Position
  oOut.Scaling = oIn.Scaling
  oOut.Shadow = oIn.Shadow
  oOut.Size = oIn.Size
  oOut.SizeBi = oIn.SizeBi
  oOut.SmallCaps = oIn.SmallCaps
  oOut.Spacing = oIn.Spacing
  oOut.StrikeThrough = oIn.StrikeThrough
  oOut.Subscript = oIn.Subscript
  oOut.Superscript = oIn.Superscript
  oOut.Underline = oIn.Underline
  oOut.UnderlineColor = oIn.UnderlineColor

  If (Not bFindParameters) Then
    oOut.Animation = oIn.Animation
    oOut.Borders = oIn.Borders
    oOut.DiacriticColor = oIn.DiacriticColor
    oOut.DisableCharacterSpaceGrid = oIn.DisableCharacterSpaceGrid
    oOut.EmphasisMark = oIn.EmphasisMark
    oOut.Shading.BackgroundPatternColorIndex = oIn.Shading.BackgroundPatternColorIndex
    oOut.Shading.ForegroundPatternColorIndex = oIn.Shading.ForegroundPatternColorIndex
    oOut.Shading.BackgroundPatternColor = oIn.Shading.BackgroundPatternColor
    oOut.Shading.ForegroundPatternColor = oIn.Shading.ForegroundPatternColor
    oOut.Shading.Texture = oIn.Shading.Texture
  End If
    
  'oOut.NameFarEast = oIn.NameFarEast
  'oOut.Application = oIn.Application
  'oOut.Creator = oIn.Creator
  'oOut.Parent = oIn.Parent
  'oOut.Duplicate = oIn.Duplicate
  'oOut.Shading.Application = oIn.Shading.Application
  'oOut.Shading.Creator = oIn.Shading.Creator
  'oOut.Shading.Parent = oIn.Shading.Parent
End Sub

Public Sub BX_CloneParagraphFormat(ByVal oIn As ParagraphFormat, ByVal oOut As ParagraphFormat)
End Sub

Public Sub BX_CloneFrame(ByVal oIn As Frame, ByVal oOut As Frame)
End Sub

Public Function BX_CountInStr(ByVal sSearchIn As String, ByVal sSearchFor As String) As Integer
  bx_sFunction = "BX_CountInStr"
  BX_CountInStr = UBound(Split(sSearchIn, sSearchFor))
End Function

Public Function BX_InStr(ByVal nStart As Long, ByVal sSearchIn As String, ByVal vSearchFor As Variant, Optional ByVal nCompare As VbCompareMethod = -1) As Long
  Dim nPos As Long
  Dim nI As Integer
  
  If (IsArray(vSearchFor)) Then
    nPos = 0
    For nI = 1 To UBound(vSearchFor)
      If (nCompare = -1) Then
        nPos = InStr(nStart, sSearchIn, CStr(vSearchFor(nI)))
      Else
        nPos = InStr(nStart, sSearchIn, CStr(vSearchFor(nI)), nCompare)
      End If
      If (nPos <> 0) Then Exit For
    Next
  Else
    If (nCompare = -1) Then
      nPos = InStr(nStart, sSearchIn, CStr(vSearchFor))
    Else
      nPos = InStr(nStart, sSearchIn, CStr(vSearchFor), nCompare)
    End If
   End If
  BX_InStr = nPos
End Function


Public Function BX_TestChar(ByVal sString As String, ByVal nCharTypes As Integer) As Boolean 'returns true if all characters in string are of one of the specific types
  Dim nPos As Long
  Dim nI As Long
  Dim bEqual As Boolean
  Dim sChar As String
  
  bEqual = False
  
  For nPos = 1 To Len(sString)
    sChar = Mid(sString, nPos, 1)
    bEqual = False
    
    If (Not bEqual And ((BX_CH_VS_SEPARATOR And nCharTypes) <> 0)) Then
      For nI = 1 To UBound(bx_asChVsSeparators)
        bEqual = (AscW(sChar) = AscW(bx_asChVsSeparators(nI)))
        If (bEqual) Then Exit For
      Next
    End If
    
    If (Not bEqual And ((BX_SEPARATOR And nCharTypes) <> 0)) Then
      For nI = 1 To UBound(bx_asSeparators)
        bEqual = (AscW(sChar) = AscW(bx_asSeparators(nI)))
        If (bEqual) Then Exit For
      Next
    End If
    
    If (Not bEqual And ((BX_NUMBER And nCharTypes) <> 0)) Then
      bEqual = (AscW(sChar) >= AscW("0") And AscW(sChar) <= AscW("9"))
    End If
    
    If (Not bEqual And ((BX_LETTER And nCharTypes) <> 0)) Then
      Select Case bx_eLanguage
        Case BX_ENGLISH
          bEqual = ((AscW(sChar) >= AscW("a") And AscW(sChar) <= AscW("z")) Or (AscW(sChar) >= AscW("A") And AscW(sChar) <= AscW("Z")))
        Case BX_GERMAN
          bEqual = ((AscW(sChar) >= AscW("a") And AscW(sChar) <= AscW("z")) Or (AscW(sChar) >= AscW("A") And AscW(sChar) <= AscW("Z"))) _
                    Or (AscW(sChar) = AscW("ä")) Or (AscW(sChar) = AscW("Ä")) Or (AscW(sChar) = AscW("ö")) Or (AscW(sChar) = AscW("Ö")) Or (AscW(sChar) = AscW("ü")) Or (AscW(sChar) = AscW("Ü")) Or (AscW(sChar) = AscW("ß"))
        Case Default 'other
          bEqual = ((AscW(sChar) >= AscW("a") And AscW(sChar) <= AscW("z")) Or (AscW(sChar) >= AscW("A") And AscW(sChar) <= AscW("Z")))
      End Select
    End If
    
    If (Not bEqual And ((BX_DASH And nCharTypes) <> 0)) Then
      bEqual = (AscW(sChar) = &H2D Or (AscW(sChar) >= &H2010 And AscW(sChar) <= &H2015))
    End If
    
    If (Not bEqual And ((BX_SPACE And nCharTypes) <> 0)) Then
      bEqual = (AscW(sChar) = &H20 Or AscW(sChar) = &HA0 Or (AscW(sChar) >= &H2000 And AscW(sChar) <= &H200D))
    End If
    
    If (Not bEqual And ((BX_RETURN And nCharTypes) <> 0)) Then
      bEqual = (AscW(sChar) = 10 Or AscW(sChar) = 13)
    End If
    
    If (Not bEqual And ((BX_SENTENCE_LETTER And nCharTypes) <> 0)) Then
      bEqual = ((AscW(UCase(sChar)) >= AscW("A") And AscW(UCase(sChar)) <= AscW(UCase(BX_MAX_SENTENCE_LETTER))))
    End If
    
    If (Not bEqual And ((BX_BOOK_NUMBER And nCharTypes) <> 0)) Then
      bEqual = ((AscW(sChar) >= AscW("1") And AscW(sChar) <= AscW("3")) Or AscW(UCase(sChar)) = AscW("I"))
    End If
    
    If (Not bEqual) Then Exit For
  Next
  
  BX_TestChar = bEqual
End Function

Public Function RegEx(strPattern As String, Optional GlobalSearch As Boolean, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean) As Object
    Dim objRegEx As Object
    Set objRegEx = CreateObject("vbscript.regexp")
    objRegEx.Global = GlobalSearch
    objRegEx.MultiLine = MultiLine
    objRegEx.IgnoreCase = IgnoreCase
    objRegEx.Pattern = strPattern
    Set RegEx = objRegEx
End Function
