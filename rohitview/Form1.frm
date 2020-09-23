VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Start 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rohit View"
   ClientHeight    =   6120
   ClientLeft      =   2745
   ClientTop       =   2070
   ClientWidth     =   7350
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   7350
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   960
      Top             =   3120
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1000
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1452
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":176C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2010
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2462
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":28B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2D06
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3158
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":35AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Value           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Value           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "slideshow"
            Object.ToolTipText     =   "SlideShow"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   5625
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8821
            MinWidth        =   8821
            Picture         =   "Form1.frx":39FC
            Text            =   "Opened Image File:-->"
            TextSave        =   "Opened Image File:-->"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "Form1.frx":3E4E
            Text            =   "Time:-->"
            TextSave        =   "Time:-->"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3353
            MinWidth        =   3353
            Picture         =   "Form1.frx":42A0
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
   Begin VB.Image Image1 
      Height          =   7335
      Left            =   0
      Stretch         =   -1  'True
      Top             =   360
      Width           =   12120
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnufileopen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnufilesave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnusepat 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileslideshow 
         Caption         =   "SildeShow"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnufilemruds 
         Caption         =   "MRUDs"
         Index           =   0
         Visible         =   0   'False
         WindowList      =   -1  'True
      End
      Begin VB.Menu mnusepartor 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "Start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub resize()
If Image1.Picture.Height <= Image1.Height Or Image1.Picture.Width <= Image1.Width Then
Image1.Stretch = False
Else
Image.Stretch = True
End If
End Sub


Private Sub Form_Load()
StatusBar1.Panels(3).Text = "Date:" & "-->" & Date
StatusBar1.Panels(2).Text = "Time:" & "-->" & Time
End Sub

Private Sub Form_Resize() 'to resize the display when the form is resized
If WindowState <> vbMinimized Then
Image1.Height = Start.ScaleHeight
Image1.Width = Start.ScaleWidth
End If
End Sub

Private Sub mnufileopen_Click()
CommonDialog1.CancelError = True
On Error GoTo cancelopen
CommonDialog1.Filter = "JPEG FILES|*.jpg|BITMAP FILES|*.bmp|TIFF FILES|*.tif"

CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
Image1.Picture = LoadPicture(CommonDialog1.FileName)
StatusBar1.Panels(1).Text = "Opened File" & "-->" & " " & CommonDialog1.FileName
resize


End If
cancelopen:
Exit Sub
End Sub

Private Sub mnufilesave_Click() 'to save a file
CommonDialog1.CancelError = True
On Error GoTo cancelopen
CommonDialog1.DialogTitle = "File Save"
CommonDialog1.Filter = "JPEG FILES|*.jpg|BITMAP FILES|*.bmp|TIFF FILES|*.tif"

CommonDialog1.ShowSave
Dim str As String
str = CommonDialog1.FileName
If CommonDialog1.FileName <> "" Then

SavePicture Image1.Picture, str
End If
cancelopen:
Exit Sub
End Sub


Private Sub mnufileslideshow_Click()
SlideShow.Show
Start.Hide
End Sub


Private Sub Timer1_Timer()
If Not StatusBar1.Panels(2).Text = "Time" & "-->" & Time Then
StatusBar1.Panels(2).Text = "Time:" & "-->" & Time
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "open":
mnufileopen_Click
Case "save":
mnufilesave_Click
Case "slideshow":
mnufileslideshow_Click
Case "delete"
End Select
End Sub
