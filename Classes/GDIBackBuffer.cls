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
Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDstDC As Long, ByVal xDst As Long, ByVal yDst As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetObjectA Lib "gdi32" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private backDC As Long
Private mPB 'As PictureBox
Attribute mPB.VB_VarHelpID = -1
Private mWidth  As Long
Private mHeight As Long
Private backBitmap As Long

Friend Sub New_(aPB) ' As PictureBox)
    'foreDC = aPB.hDC
    Set mPB = aPB
    'Create device context for the background
    'backDC = CreateCompatibleDC(GetDC(0))
    backDC = CreateCompatibleDC(mPB.hDC)
End Sub
Private Sub Class_Terminate()
    DeleteDC backDC
    DeleteObject backBitmap
End Sub

Public Sub Clear()
    backBitmap = CreateCompatibleBitmap(mPB.hDC, mWidth, mHeight)
    SelectObject backDC, backBitmap
    DeleteObject backBitmap
End Sub

Public Property Get hDC() As Long
    hDC = backDC
End Property

Public Sub DrawBitmap(bmp As GDIBitmap, ByVal X As Long, Y As Long)
    BitBlt backDC, X, Y, bmp.Width, bmp.Height, bmp.hDC, 0, 0, RasterOpConstants.vbSrcCopy
End Sub

Public Sub Paint()
    BitBlt mPB.hDC, 0, 0, mPB.ScaleWidth, mPB.ScaleHeight, backDC, 0, 0, RasterOpConstants.vbSrcCopy
End Sub

Public Sub Resize()
    'Create bitmap of proper size for the background
    mWidth = mPB.ScaleWidth
    mHeight = mPB.ScaleHeight
    Clear
End Sub
'Private Sub mPB_Resize()
'    Resize
'    'Draw picture in the middle of background
'    'BitBlt backDC, Int((Me.ScaleWidth - picInfo.bmWidth) / 2), Int((Me.ScaleHeight - picInfo.bmHeight) / 2), picInfo.bmWidth, picInfo.bmHeight, picDC, 0, 0, vbSrcCopy
'
'End Sub
