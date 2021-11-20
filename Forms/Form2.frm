VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Doubleclick to switch on off"
   ClientHeight    =   10695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13215
   LinkTopic       =   "Form2"
   ScaleHeight     =   10695
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      FillColor       =   &H008080FF&
      ForeColor       =   &H00FFFF00&
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6075
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mGBB As GDIBackBuffer
Private mbIsRunning As Boolean
Private mCounter As Long
Private mTimer   As Single
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare Function Rectangle Lib "gdi32" (ByVal hhdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Sub Form_Load()
    Timer1.Interval = 1 '0 '10
    Me.ScaleMode = vbPixels
    Picture1.ScaleMode = vbPixels
    Set mGBB = GDIBackBuffer(Picture1)
End Sub
Public Function GDIBackBuffer(aPB) As GDIBackBuffer 'As PictureBox) As GDIBackBuffer
    Set GDIBackBuffer = New GDIBackBuffer: GDIBackBuffer.New_ aPB
End Function

Private Sub Form_Paint()
    Picture1_Resize
End Sub

Private Sub Form_Resize()
    Dim brdr As Single ': brdr = 8 '* Screen.TwipsPerPixelX
    Dim W As Single: W = Me.ScaleWidth - 2 * brdr
    Dim H As Single: H = Me.ScaleHeight - 2 * brdr
    If W > 0 And H > 0 Then Picture1.Move brdr, brdr, W, H
End Sub

Private Sub Picture1_DblClick()
    mbIsRunning = Not mbIsRunning
    AniTimer
End Sub

Private Sub Picture1_Paint()
    mGBB.Paint
End Sub

Private Sub Picture1_Resize()
    mGBB.Resize
    mGBB.Paint
End Sub

Private Sub AniTimer()
    Dim t As Long: t = 10
    If mbIsRunning Then
        Do
            If (mCounter Mod t) = 0 Then
                Dim d As Double: d = Timer - mTimer
                mTimer = Timer
                Me.Caption = "Generations: " & CStr(mCounter) & "   " & CStr(CLng(t / d)) & " per sec"
                DoEvents
            End If
            If Not mbIsRunning Then Exit Do
            DrawCells
            Sleep 1 '0
        Loop
    End If
End Sub

Private Sub DrawCells()
    mGBB.Clear
    'Picture1.ForeColor = vbBlue 'nope
    Dim n  As Long:    n = 180
    Dim sz As Double: sz = Picture1.ScaleHeight / n
    
    Dim X1 As Long, Y1 As Long
    Dim X2 As Long, Y2 As Long
    Dim hDC As Long: hDC = mGBB.hDC
    Dim i As Long, j As Long
    Dim rrnd As Double
    Randomize
    For i = 0 To n - 1
        For j = 0 To n - 1
            X1 = i * sz: X2 = X1 + sz
            Y1 = j * sz: Y2 = Y1 + sz
            rrnd = Rnd
            If 0.75 < (rrnd) Then
                Rectangle hDC, X1, Y1, X2, Y2
            End If
        Next
    Next
    mCounter = mCounter + 1
    mGBB.Paint
End Sub
