Attribute VB_Name = "MNew"
Option Explicit

Sub Main()
    Form1.Show
    Form2.Show
End Sub

Public Function GDIBackBuffer(aPB) As GDIBackBuffer 'As PictureBox) As GDIBackBuffer
    Set GDIBackBuffer = New GDIBackBuffer: GDIBackBuffer.New_ aPB
End Function

Public Function GDIBitmap(aPFN As String) As GDIBitmap
    Set GDIBitmap = New GDIBitmap: GDIBitmap.New_ aPFN
End Function

