VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type BITMAP
    bmType       As Long
    Width        As Long
    Height       As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type
Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr

Private Declare Function BitBlt Lib "gdi32" (ByVal hDstDC As LongPtr, ByVal xDst As Long, ByVal yDst As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
                                             ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr

Private Declare Function GetObjectA Lib "gdi32" (ByVal hObject As LongPtr, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long

Private m_PFN     As String
Private m_picDC   As LongPtr
Private m_picInfo As BITMAP
Private m_PicDisp As IPictureDisp

Friend Sub New_(aPFN As String)
    m_PFN = aPFN
    LoadFromFile m_PFN
End Sub

Public Sub LoadFromFile(aPFN As String)
    m_picDC = CreateCompatibleDC(GetDC(0))
    Set m_PicDisp = LoadPicture(aPFN)
    SelectObject m_picDC, m_PicDisp
    'GetObjectA LoadPicture(aPFN), Len(picInfo), picInfo
    GetObjectA m_PicDisp, Len(m_picInfo), m_picInfo
End Sub

Public Property Get hDC() As LongPtr
    hDC = m_picDC
End Property

Public Property Get Width() As Long
    Width = m_picInfo.Width
End Property
Public Property Get Height() As Long
    Height = m_picInfo.Height
End Property
Public Property Get Width2() As Long
    Width2 = m_picInfo.Width / 2
End Property
Public Property Get Height2() As Long
    Height2 = m_picInfo.Height / 2
End Property
'
'Public Sub Draw(TohDC As Long, ByVal X As Long, ByVal Y As Long)
'    BitBlt TohDC, X - Me.Width2, Y - Me.Height2, Me.Width, Me.Height, picDC, 0, 0, RasterOpConstants.vbSrcCopy
'End Sub
