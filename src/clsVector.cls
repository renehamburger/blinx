VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVector"
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
' Interface for dynamic vector array
' See https://github.com/renehamburger/blinx for source code, manual & license
'==============================================================================

Private m_avData() As Variant
Private m_nTotalSize
Private m_nCurrentSize

Private Sub Class_Initialize()
  ReDim m_avData(8)
  m_nTotalSize = 8
  m_nCurrentSize = 0
End Sub

Public Sub Add(ByVal vAdd As Variant)
  m_avData(m_nCurrentSize + 1) = vAdd
  m_nCurrentSize = m_nCurrentSize + 1
  If (m_nCurrentSize = m_nTotalSize) Then
    m_nTotalSize = 2 * m_nTotalSize
    ReDim Preserve m_avData(m_nTotalSize)
  End If
End Sub

Public Function Size()
  Size = m_nCurrentSize
End Function

Public Sub Clear()
  ReDim m_avData(8)
  m_nTotalSize = 8
  m_nCurrentSize = 0
End Sub

Public Property Get data(ByVal nIndex As Integer)
  If (nIndex > 0 And nIndex <= m_nCurrentSize) Then
    data = m_avData(nIndex)
  Else
    data = Empty
  End If
End Property


Public Property Let data(ByVal nIndex As Integer, ByVal vValue As Variant)
  If (nIndex > m_nTotalSize) Then
    m_nTotalSize = Round(nIndex / 2# + 0.5) * 2
    ReDim Preserve m_avData(m_nTotalSize)
  End If
  
  m_avData(nIndex) = vValue
End Property
