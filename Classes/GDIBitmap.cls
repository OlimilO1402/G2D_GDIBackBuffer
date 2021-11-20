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
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDstDC As Long, ByVal xDst As Long, ByVal yDst As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function GetObjectA Lib "gdi32.dll" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private mPFN     As String
Private picDC    As Long
Private picInfo  As BITMAP
Private mPicDisp As IPictureDisp

Friend Sub New_(aPFN As String)
    mPFN = aPFN
    LoadFromFile mPFN
End Sub

Public Sub LoadFromFile(aPFN As String)
    picDC = CreateCompatibleDC(GetDC(0))
    Set mPicDisp = LoadPicture(aPFN)
    SelectObject picDC, mPicDisp
    'GetObjectA LoadPicture(aPFN), Len(picInfo), picInfo
    GetObjectA mPicDisp, Len(picInfo), picInfo
End Sub

Public Property Get hDC() As Long
    hDC = picDC
End Property

Public Property Get Width() As Long
    Width = picInfo.Width
End Property
Public Property Get Height() As Long
    Height = picInfo.Height
End Property
Public Property Get Width2() As Long
    Width2 = picInfo.Width / 2
End Property
Public Property Get Height2() As Long
    Height2 = picInfo.Height / 2
End Property
'
'Public Sub Draw(TohDC As Long, ByVal X As Long, ByVal Y As Long)
'    BitBlt TohDC, X - Me.Width2, Y - Me.Height2, Me.Width, Me.Height, picDC, 0, 0, RasterOpConstants.vbSrcCopy
'End Sub
