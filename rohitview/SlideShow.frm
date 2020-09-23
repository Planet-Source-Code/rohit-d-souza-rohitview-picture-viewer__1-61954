VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form SlideShow 
   BackColor       =   &H80000008&
   Caption         =   "Slide Show"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   Icon            =   "SlideShow.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5175
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   3600
      Top             =   6720
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      ItemData        =   "SlideShow.frx":0442
      Left            =   6840
      List            =   "SlideShow.frx":044F
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Picture Preview"
      ForeColor       =   &H00FF8080&
      Height          =   2775
      Left            =   360
      TabIndex        =   22
      Top             =   5040
      Width           =   2895
      Begin VB.Image imgpreview 
         Height          =   2175
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   1200
      Left            =   6840
      TabIndex        =   4
      Top             =   2520
      Width           =   1815
   End
   Begin VB.DirListBox dirdirectory 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   990
      Left            =   6840
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.DriveListBox drvdrive 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   6840
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Caption         =   "Slide Show Option"
      ForeColor       =   &H00FF8080&
      Height          =   1695
      Left            =   3720
      TabIndex        =   18
      Top             =   3960
      Width           =   5535
      Begin VB.OptionButton optuserauto 
         BackColor       =   &H80000008&
         Caption         =   "Show Next Image Automatic after Mouse/Keyboard Input"
         ForeColor       =   &H00FF8080&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   4575
      End
      Begin VB.OptionButton optTimeauto 
         BackColor       =   &H80000012&
         Caption         =   "Show Next Image Automic After"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txttime 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Seconds"
         ForeColor       =   &H00FF8080&
         Height          =   375
         Left            =   3840
         TabIndex        =   20
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FF8080&
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmddown 
      Appearance      =   0  'Flat
      Caption         =   "Move Down"
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdup 
      Appearance      =   0  'Flat
      Caption         =   "Move Up"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      Caption         =   "Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdPlay 
      Appearance      =   0  'Flat
      Caption         =   "Play"
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdRemoveall 
      Appearance      =   0  'Flat
      Caption         =   "Remove All"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdRemove 
      Appearance      =   0  'Flat
      Caption         =   "Remove"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdaddall 
      Appearance      =   0  'Flat
      Caption         =   "Add All"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   30
         Left            =   600
         TabIndex        =   24
         Top             =   4920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   53
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   1
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
      Begin VB.ListBox lstslidefiles 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   4125
         ItemData        =   "SlideShow.frx":0469
         Left            =   120
         List            =   "SlideShow.frx":046B
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton cmdAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         Caption         =   "Add"
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Slide Show Files"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2055
      End
   End
   Begin MSComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   4800
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8821
            MinWidth        =   8821
            Picture         =   "SlideShow.frx":046D
            Text            =   "Opened Image File:-->"
            TextSave        =   "Opened Image File:-->"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "SlideShow.frx":08BF
            Text            =   "Time:-->"
            TextSave        =   "Time:-->"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3353
            MinWidth        =   3353
            Picture         =   "SlideShow.frx":0D11
            Text            =   "Date:-->"
            TextSave        =   "Date:-->"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "Select File Type For Slide Show:"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "Look In:"
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   3720
      TabIndex        =   21
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "SlideShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    

Dim str1 As String
Public interval1 As Long


Private Sub cmdAdd_Click() 'to add files to the list for slide show
Static int1 As Integer
If File1.FileName <> "" Then
lstslidefiles.AddItem (File1.Path & "\" & File1.FileName)
Else
MsgBox "Plz Click on an Image File which you want To add to slide show", , "Rohitview"
File1.SetFocus
End If
End Sub

Private Sub cmdaddall_Click() 'to add all files from file list into the list for slide show
Dim int1 As Integer
If File1.ListCount > 0 Then
int1 = File1.ListCount
For i = 0 To int1 - 1
lstslidefiles.AddItem File1.Path & "\" & File1.List(i)
Next
Else
MsgBox "Plz Select an Image Folder which you want To add to slide show", , "RohitView"
drvdrive.SetFocus
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
Start.Show
End Sub

Private Sub cmddown_Click()
If Not lstslidefiles.ListIndex = lstslidefiles.ListCount - 1 Then
lstslidefiles.ListIndex = lstslidefiles.ListIndex + 1
End If
End Sub

Private Sub cmdPlay_Click() 'to start the slide show
Dim interval1 As Double
If lstslidefiles.ListCount > 0 Then
If optTimeauto.Value = True Then
If txttime.Text <> "" And IsNumeric(txttime.Text) Then
interval1 = CDbl(txttime.Text)
SlideShow.Hide
showslide.Show
showslide.timedshow (interval1) 'passing interval value to showslide form
Else
MsgBox "Plz enter a valid time interval", , "Rohitview"
txttime.SetFocus
End If
ElseIf optuserauto.Value = True Then
showslide.Show
SlideShow.Hide
Else
MsgBox "plz select a option for slide show image display", , "Rohitview"
optTimeauto.SetFocus
End If
Else
MsgBox "PLZ select files to be displayed in the slide show", , "Rohitview"
dirdirectory.SetFocus
End If
End Sub

Private Sub cmdRemove_Click()
If lstslidefiles.ListCount > 0 Then
lstslidefiles.RemoveItem lstslidefiles.ListIndex
Else
MsgBox "No file selected in the slide show to be removed", , "Rohitview"
End If
End Sub

Private Sub cmdRemoveall_Click()
If lstslidefiles.ListCount > 0 Then
lstslidefiles.Clear
Else
MsgBox "No files selected in the slide show to be removed", , "Rohitview"
End If

End Sub

Private Sub cmdup_Click()
If Not lstslidefiles.ListIndex = 0 Then
lstslidefiles.ListIndex = lstslidefiles.ListIndex - 1
End If
End Sub

Private Sub Combo1_Click()
str1 = Combo1.Text
File1.Pattern = str1
End Sub

Private Sub dirdirectory_Change()
File1.Path = dirdirectory.Path
End Sub

Private Sub drvdrive_Change()
Dim i As Integer
On Error GoTo eerr
dirdirectory.Path = drvdrive.Drive
eerr:
If Err.Number = 68 Then 'find if error is of type 68 ie. device unavailable
MsgBox "plz Insert a disk in " & " " & drvdrive.Drive, , "RohitView"
drvdrive.Drive = dirdirectory.Path
End If
End Sub

Private Sub File1_Click()
Dim str As String
str = File1.Path & "\" & File1.FileName
imgpreview.Picture = LoadPicture(str)
End Sub

Private Sub File1_DblClick()
lstslidefiles.AddItem File1.List(File1.ListIndex)
End Sub

Private Sub Form_Load()
StatusBar2.Panels(3).Text = "Date:" & "-->" & Date
StatusBar2.Panels(2).Text = "Time:" & "-->" & Time
optTimeauto.Value = False
optuserauto.Value = False
txttime.Text = ""
File1.Pattern = "*.jpg"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstslidefiles_Click()
Dim str As String
str = lstslidefiles.Text
imgpreview.Picture = LoadPicture(str)
End Sub

Private Sub Timer1_Timer()
If Not StatusBar2.Panels(2).Text = "Time" & "-->" & Time Then
StatusBar2.Panels(2).Text = "Time:" & "-->" & Time
End If
End Sub
