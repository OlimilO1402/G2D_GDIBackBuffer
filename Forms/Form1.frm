VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mGBB As GDIBackBuffer
Private mPic As GDIBitmap
Private Const filename As String = "WaterFall.jpg"

Private Sub Form_Load()
    Me.ScaleMode = vbPixels
    Set mPic = GDIBitmap(App.Path & "\Resources\" & filename)
    Set mGBB = GDIBackBuffer(Me)
End Sub

Public Function GDIBackBuffer(aPB) As GDIBackBuffer 'As PictureBox) As GDIBackBuffer
    Set GDIBackBuffer = New GDIBackBuffer: GDIBackBuffer.New_ aPB
End Function
Public Function GDIBitmap(aPFN As String) As GDIBitmap
    Set GDIBitmap = New GDIBitmap: GDIBitmap.New_ aPFN
End Function

Private Sub Form_Resize()
    mGBB.Resize
    mGBB.DrawBitmap mPic, Me.ScaleWidth / 2 - mPic.Width2, Me.ScaleHeight / 2 - mPic.Height2
    Form_Paint
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mGBB.DrawBitmap mPic, X - mPic.Width2, Y - mPic.Height2
    Form_Paint
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button > 0 Then
        mGBB.DrawBitmap mPic, X - mPic.Width2, Y - mPic.Height2
        Form_Paint
    End If
End Sub

Private Sub Form_Paint()
    mGBB.Paint
End Sub

