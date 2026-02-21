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
Private m_GBB As GDIBackBuffer
Private m_Pic As GDIBitmap
Private Const filename As String = "WaterFall.jpg"

Private Sub Form_Load()
    Me.ScaleMode = vbPixels
    Set m_Pic = MNew.GDIBitmap(App.Path & "\Resources\" & filename)
    Set m_GBB = MNew.GDIBackBuffer(Me)
End Sub

Private Sub Form_Resize()
    m_GBB.Resize
    m_GBB.DrawBitmap m_Pic, Me.ScaleWidth / 2 - m_Pic.Width2, Me.ScaleHeight / 2 - m_Pic.Height2
    Form_Paint
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_GBB.DrawBitmap m_Pic, X - m_Pic.Width2, Y - m_Pic.Height2
    Form_Paint
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button > 0 Then
        m_GBB.DrawBitmap m_Pic, X - m_Pic.Width2, Y - m_Pic.Height2
        Form_Paint
    End If
End Sub

Private Sub Form_Paint()
    m_GBB.Paint
End Sub

