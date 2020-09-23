VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BF67E8C4-7E0E-41BC-9068-7A106BD576BD}#1.0#0"; "ANIMATEDGIF.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "HF-Solutions mp3"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "&Close Volume"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Open Volume"
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   5280
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin AnimatedGif.AnimatedGifCtl AnimatedGifCtl1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1720
      BorderStyle     =   1
      strGifFileName  =   "C:\My Documents\My Pictures\equalizer.gif"
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      _Version        =   327682
      BorderStyle     =   1
      TickStyle       =   2
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   5400
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
MediaPlayer1.FileName = CommonDialog1.FileTitle
MediaPlayer1.Play
Text1.Text = CommonDialog1.FileTitle
AnimatedGifCtl1.StartTimer
Slider1.Max = MediaPlayer1.Duration
End Sub

Private Sub Command2_Click()
MediaPlayer1.Stop
AnimatedGifCtl1.Refresh
AnimatedGifCtl1.StopGif
Slider1.Value = 0
End Sub

Private Sub Command3_Click()
CommonDialog1.Filter = "MP3 Files|*.MP3"
CommonDialog1.ShowOpen
End Sub



Private Sub Command4_Click()
frmVolume.Visible = False
End Sub

Private Sub Command5_Click()
frmVolume.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Me.WindowState = 0


    Do
        Me.Top = Me.Top + 10
        Me.Left = Me.Left + 10
        Me.Width = Me.Width - 20
        Me.Height = Me.Height - 20
    Loop Until Me.Top >= Screen.Height
    frmVolume.Enabled = False
    'you can change those numbers to make it
    '     faster
    'or slower. right now it is pretty slow.
    '
    'if the height and width #'s are twice a
    '     s much
    'as the top and left #'s, it will make a
    '
    ' "zooming out" effect and then will fal
    '     l to the
    'bottom of the screen.
End Sub


Private Sub Slider1_Scroll()
MediaPlayer1.CurrentPosition = Slider1.Value
End Sub

Private Sub Timer1_Timer()
Slider1.Value = MediaPlayer1.CurrentPosition
End Sub
