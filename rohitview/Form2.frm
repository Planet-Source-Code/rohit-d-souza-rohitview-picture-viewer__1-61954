VERSION 5.00
Begin VB.Form showslide 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "slide"
   ClientHeight    =   7035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   5400
      Top             =   1200
   End
   Begin VB.Image imgSlide 
      Height          =   7095
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "showslide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Form_Click()
check
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 13:
check
Case 27:
Unload Me
Start.Show
End Select
End Sub

Private Sub Form_Load()
showslide.Top = 0
showslide.Left = 0
showslide.Height = Screen.Height
showslide.Width = Screen.Width
imgSlide.Height = showslide.Height
imgSlide.Width = showslide.Width
i = 0
Timer1.Enabled = False
userautoshow
End Sub

Private Sub imgSlide_Click()
check
End Sub

Private Sub Timer1_Timer()
If i <= SlideShow.lstslidefiles.ListCount - 1 Then
imgSlide.Picture = LoadPicture(SlideShow.lstslidefiles.List(i))
resize
i = i + 1
End If
If Not i <= SlideShow.lstslidefiles.ListCount - 1 Then
Timer1.Enabled = False
End If
End Sub

Public Function timedshow(interval1 As Double)
Timer1.Interval = (interval1 * 1000)
Timer1.Enabled = True
End Function

Public Function userautoshow()
If i <= SlideShow.lstslidefiles.ListCount - 1 Then
imgSlide.Picture = LoadPicture(SlideShow.lstslidefiles.List(i))
resize
i = i + 1
Else
Unload Me
Start.Show
End If
End Function


Public Sub check()
If Timer1.Enabled = True Then
Timer1.Enabled = False
Unload Me
Start.Show
End If
If Timer1.Enabled = False Then
userautoshow
End If
End Sub

Public Sub resize()
If imgSlide.Picture.Height <= imgSlide.Height Or imgSlide.Picture.Width <= imgSlide.Width Then
imgSlide.Stretch = False

End If
End Sub
