VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClipboard"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
'==============================================================================
' Clipboard class
' See https://github.com/renehamburger/blinx for source code, manual & license
'---
' The code below is based on work by Daniel Cleveland (http://www.francescofoti.com/2013/12/
' share-the-clipboard-with-vba-and-the-windows-api/) and Leigh Webber
' (https://social.msdn.microsoft.com/Forums/office/en-US/ee9e0d28-0f1e-467f-
' 8d1d-1a86b2db2878/a-clipboard-object-for-vba-including-microsoft-word?forum=worddev)
' and LA Solutions (http://www.la-solutions.co.uk/content/MVBA/MVBA-Clipboard.htm)
'==============================================================================
#If Win64 Then
  'To copy text on the clipboard
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
  Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongLong) As Long
  Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongLong) As LongPtr
  Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongLong) As LongPtr
  Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongLong) As Long
  Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As String) As LongPtr
  Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
  Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongLong) As LongLong
  Private Declare PtrSafe Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
  Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
  Private Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
#Else
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
  Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
  Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
  Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
  Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
  Private Declare Function CloseClipboard Lib "user32" () As Long
  Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function EmptyClipboard Lib "user32" () As Long
  Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
  Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
  Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
  Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
  Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
  Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
#End If

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT) 'Use for hwnd
Private nCF_HTML As Long
Private nCF_RTF As Long

Public Enum eCBFormat
  CF_TEXT = 1
  CF_BITMAP = 2
  CF_METAFILEPICT = 3
  CF_SYLK = 4
  CF_DIF = 5
  CF_TIFF = 6
  CF_OEMTEXT = 7
  CF_DIB = 8
  CF_PALETTE = 9
  CF_PENDATA = 10
  CF_RIFF = 11
  CF_WAVE = 12
  CF_UNICODETEXT = 13
  CF_ENHMETAFILE = 14
  CF_HDROP = 15
  CF_LOCALE = 16
  CF_MAX = 17
  CF_OWNERDISPLAY = &H80
  CF_DSPTEXT = &H81
  CF_DSPBITMAP = &H82
  CF_DSPMETAFILEPICT = &H83
  CF_DSPENHMETAFILE = &H8E
  CF_PRIVATEFIRST = &H200
  CF_PRIVATELAST = &H2FF
  CF_GDIOBJFIRST = &H300
  CF_GDIOBJLAST = &H3FF
End Enum

' --------------------------------------------------------------------
'  Public Methods
' --------------------------------------------------------------------

#If Win64 Then
 Public Function HasFormat(ByVal peCBFormat As eCBFormat, Optional ByVal phWnd As LongLong = 0) As Boolean
#Else
 Public Function HasFormat(ByVal peCBFormat As eCBFormat, Optional ByVal phWnd As Long = 0) As Boolean
#End If
  Dim lRet    As Long

  If OpenClipboard(phWnd) > 0 Then
    lRet = EnumClipboardFormats(0)
    If lRet <> 0 Then
      Do
        If lRet = peCBFormat Then
          HasFormat = True
          Exit Do
        End If
        lRet = EnumClipboardFormats(lRet)
      Loop While lRet <> 0
    End If
    CloseClipboard
  Else
    'Problem: Cannot open clipboard
  End If
End Function

#If Win64 Then
 Public Function GetData(ByVal vFormat As Variant, Optional ByVal phWnd As LongLong = 0) As String
  Dim hData       As LongPtr
  Dim lByteLen    As LongPtr
  Dim lPointer    As LongPtr
  Dim lSize       As LongLong
#Else
 Public Function GetData(ByVal vFormat As Variant, Optional ByVal phWnd As Long = 0) As String
  Dim hData       As Long
  Dim lByteLen    As Long
  Dim lPointer    As Long
  Dim lSize       As Long
#End If
  Dim lRet        As Long
  Dim abData()    As Byte
  Dim sText       As String

  If VarType(vFormat) = vbString Then
    vFormat = RegisterClipboardFormat(vFormat)
  ElseIf VarType(vFormat) <> vbLong Then
    Exit Function
  End If

  lRet = OpenClipboard(phWnd)
  If lRet > 0 Then
    hData = GetClipboardData(vFormat)
    If hData <> 0 Then
      lByteLen = GlobalSize(hData)
      lSize = GlobalSize(hData)
      lPointer = GlobalLock(hData)
      If lSize > 0 Then
        ReDim abData(0 To CLng(lSize) - CLng(1)) As Byte
        CopyMemory abData(0), ByVal lPointer, lSize
        GlobalUnlock hData
        sText = StrConv(abData, vbUnicode)
      End If
    Else
      'Problem: Cannot open clipboard
    End If
    CloseClipboard
  End If

  GetData = sText
End Function

Public Function SetData(ByVal vFormat As Variant, psData As String) As Boolean
  #If Win64 Then
    Dim hGlobalMemory   As LongLong
    Dim lpGlobalMemory  As LongPtr
    Dim hClipMemory     As LongLong
  #Else
    Dim hGlobalMemory   As Long
    Dim lpGlobalMemory  As Long
    Dim hClipMemory     As Long
  #End If
  Dim fOK             As Boolean
  
  If VarType(vFormat) = vbString Then
    vFormat = RegisterClipboardFormat(vFormat)
  ElseIf VarType(vFormat) <> vbLong Then
    Exit Function
  End If

  fOK = True

  ' Allocate moveable global memory.
  #If Win64 Then
    hGlobalMemory = GlobalAlloc(GHND, LenB(psData) + 1)
  #Else
    hGlobalMemory = GlobalAlloc(GHND, Len(psData) + 1)
  #End If
  If hGlobalMemory = 0 Then
    Exit Function
  End If
  ' Lock the block to get a far pointer
  ' to this memory.
  lpGlobalMemory = GlobalLock(hGlobalMemory)
  ' Copy the string to this global memory.
  lpGlobalMemory = lstrcpy(lpGlobalMemory, psData)
  ' Unlock the memory.
  If GlobalUnlock(hGlobalMemory) <> 0 Then
    fOK = False
    GoTo OutOfHere2
  End If
  ' Open the Clipboard to copy data to.
  If OpenClipboard(0&) = 0 Then
    Exit Function
  End If
  ' Clear the Clipboard.
  Call EmptyClipboard
  ' Copy the data to the Clipboard.
  hClipMemory = SetClipboardData(vFormat, hGlobalMemory)

OutOfHere2:
   Call CloseClipboard
   SetData = fOK
End Function

#If Win64 Then
 Public Function GetText(Optional ByVal phWnd As LongLong = 0) As String
#Else
 Public Function GetText(Optional ByVal phWnd As Long = 0) As String
#End If
  Dim nFormat As Long
  If Me.HasFormat(CF_TEXT, phWnd) Then
    nFormat = CF_TEXT
  ElseIf Me.HasFormat(CF_OEMTEXT, phWnd) Then
    nFormat = CF_OEMTEXT
  ElseIf Me.HasFormat(CF_DSPTEXT, phWnd) Then
    nFormat = CF_DSPTEXT
  Else
    Exit Function
  End If
  GetText = Me.GetData(nFormat, phWnd)
End Function

Public Function SetText(psData As String) As Boolean
  SetText = Me.SetData(CF_TEXT, psData)
End Function

Public Sub EmptyIt()
  If OpenClipboard(0&) Then EmptyClipboard
  CloseClipboard
End Sub

Public Function Save() As Boolean
  Dim bOK As Boolean
  Dim nFormat As Long
  Dim nSize As Long
  Dim hData As Long
  Dim hData2 As Long
  Dim hData3 As Long

  bOK = False
  nBufferFormat = 0
  If OpenClipboard(0&) Then
    nFormat = 0
    If (Me.HasFormat(nCF_RTF)) Then
      nFormat = nCF_RTF
    ElseIf (Me.HasFormat(CF_TEXT)) Then
      nFormat = CF_TEXT
    End If

    If (nFormat <> 0) Then
      hData = GetClipboardData(nFormat)
      If (hData) Then
        nSize = GlobalSize(hData)
        hData2 = GlobalLock(hData)
        ReDim anBuffer(0 To (nSize - 1)) As Byte
        CopyMemory anBuffer(0), ByVal hData2, nSize
        GlobalUnlock (hData)
        nBufferFormat = nFormat
        bOK = True
      End If
    End If
    CloseClipboard
  End If
  Save = bOK
End Function

Public Function Restore() As Boolean
  Dim bOK As Boolean
  Dim hData As Long
  Dim hData2 As Long

  bOK = False
  If (nBufferFormat <> 0 And OpenClipboard(0&)) Then
    EmptyClipboard 'Necessary for taking ownership
    hData = GlobalAlloc(GHND, UBound(anBuffer) + 1)
    If hData Then
      hData2 = GlobalLock(hData)
      CopyMemory ByVal hData2, anBuffer(0), UBound(anBuffer) + 1
      GlobalUnlock (hData)
      bOK = CBool(Me.SetData(nBufferFormat, hData))
    End If
    CloseClipboard
  End If
  Restore = bOK
End Function

Public Function SetHTML(ByVal sText As String) As Boolean
  Dim sHTML As String
  Dim sNumber As String
  Const nStartHTML As Long = 105
  Const nBeforeFragment As Long = 38 '143
  Const nAfterFragment As Long = 36
  Dim nStartFragment As Long
  Dim nEndFragment As Long
  Dim nEndHTML As Long
  
  nStartFragment = nStartHTML + nBeforeFragment
  nEndFragment = nStartFragment + Len(sText)
  nEndHTML = nEndFragment + nAfterFragment
  
  sHTML = sHTML & "Version:0.9" & vbCrLf
  sHTML = sHTML & "StartHTML:" & Format(nStartHTML, "0000000000") & vbCrLf
  sHTML = sHTML & "EndHTML:" & Format(nEndHTML, "0000000000") & vbCrLf
  sHTML = sHTML & "StartFragment:" & Format(nStartFragment, "0000000000") & vbCrLf
  sHTML = sHTML & "EndFragment:" & Format(nEndFragment, "0000000000") & vbCrLf
  sHTML = sHTML & "<html>" & vbCrLf
  sHTML = sHTML & "<body>" & vbCrLf
  sHTML = sHTML & "<!--StartFragment-->" & vbCrLf
  sHTML = sHTML & sText & vbCrLf
  sHTML = sHTML & "<!--EndFragment-->" & vbCrLf
  sHTML = sHTML & "</body>" & vbCrLf
  sHTML = sHTML & "</html>" & vbCrLf
  
  If (nCF_HTML <> 0) Then
    SetHTML = Me.SetData(nCF_HTML, sHTML)
  End If
End Function

' --------------------------------------------------------------------
'  Private Methods
' --------------------------------------------------------------------
Private Sub Class_Initialize()
  nCF_HTML = RegisterClipboardFormat("HTML Format")
  nCF_RTF = RegisterClipboardFormat("Rich Text Format")
End Sub
