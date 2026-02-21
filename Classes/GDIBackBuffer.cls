VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIBackBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Private Declare Function BitBlt Lib "gdi32" (ByVal hDstDC As LongPtr, ByVal xDst As Long, ByVal yDst As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
                                             ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare Function GetObjectA Lib "gdi32" (ByVal hObject As LongPtr, ByVal nCount As Long, ByRef lpObject As Any) As LongPtr
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long

Private m_backDC As LongPtr
Private m_PB 'As PictureBox
Attribute m_PB.VB_VarHelpID = -1
Private m_Width  As Long
Private m_Height As Long
Private m_backBitmap As LongPtr

Friend Sub New_(aPB) ' As PictureBox)
    'foreDC = aPB.hDC
    Set m_PB = aPB
    'Create device context for the background
    'backDC = CreateCompatibleDC(GetDC(0))
    m_backDC = CreateCompatibleDC(m_PB.hDC)
End Sub
'Copy To MNew.bas:
'Public Function GDIBackBuffer(aPB) As GDIBackBuffer 'As PictureBox) As GDIBackBuffer
'    Set GDIBackBuffer = New GDIBackBuffer: GDIBackBuffer.New_ aPB
'End Function

Private Sub Class_Terminate()
    DeleteDC m_backDC
    DeleteObject m_backBitmap
End Sub

Public Sub Clear()
    m_backBitmap = CreateCompatibleBitmap(m_PB.hDC, m_Width, m_Height)
    SelectObject m_backDC, m_backBitmap
    DeleteObject m_backBitmap
End Sub

Public Property Get hDC() As LongPtr
    hDC = m_backDC
End Property

'VBRUN:
'Enum RasterOpConstants
'    vbNotSrcErase =  1114278 '(&H1100A6)
'    vbNotSrcCopy  =  3342344 '(&H330008)
'    vbSrcErase    =  4457256 '(&H440328)
'    vbDstInvert   =  5570569 '(&H550009)
'    vbPatInvert   =  5898313 '(&H5A0049)
'    vbSrcInvert   =  6684742 '(&H660046)
'    vbSrcAnd      =  8913094 '(&H8800C6)
'    vbMergePaint  = 12255782 '(&HBB0226)
'    vbMergeCopy   = 12583114 '(&HC000CA)
'    vbSrcCopy     = 13369376 '(&HCC0020)
'    vbSrcPaint    = 15597702 '(&HEE0086)
'    vbPatCopy     = 15728673 '(&HF00021)
'    vbPatPaint    = 16452105 '(&HFB0A09)
'End Enum

Public Sub DrawBitmap(bmp As GDIBitmap, ByVal X As Long, Y As Long)
    BitBlt m_backDC, X, Y, bmp.Width, bmp.Height, bmp.hDC, 0, 0, RasterOpConstants.vbSrcCopy
End Sub

Public Sub Paint()
    BitBlt m_PB.hDC, 0, 0, m_PB.ScaleWidth, m_PB.ScaleHeight, m_backDC, 0, 0, RasterOpConstants.vbSrcCopy
End Sub

Public Sub Resize()
    'Create bitmap of proper size for the background
    m_Width = m_PB.ScaleWidth
    m_Height = m_PB.ScaleHeight
    Clear
End Sub
'Private Sub mPB_Resize()
'    Resize
'    'Draw picture in the middle of background
'    'BitBlt backDC, Int((Me.ScaleWidth - picInfo.bmWidth) / 2), Int((Me.ScaleHeight - picInfo.bmHeight) / 2), picInfo.bmWidth, picInfo.bmHeight, picDC, 0, 0, vbSrcCopy
'
'End Sub
