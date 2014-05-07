VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form video 
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   Picture         =   "video.frx":0000
   ScaleHeight     =   6540
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   2520
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      MouseIcon       =   "video.frx":A7BE
      MousePointer    =   99  'Custom
      Picture         =   "video.frx":AD48
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PLAY"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1920
      MouseIcon       =   "video.frx":D8CB
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SAVE AS"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3480
      MouseIcon       =   "video.frx":DE55
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "OPEN"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      MouseIcon       =   "video.frx":E3DF
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Height          =   735
      Left            =   7440
      MouseIcon       =   "video.frx":E969
      MousePointer    =   99  'Custom
      Picture         =   "video.frx":EEF3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Height          =   735
      Left            =   6720
      MouseIcon       =   "video.frx":1176C
      MousePointer    =   99  'Custom
      Picture         =   "video.frx":11CF6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   735
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   5655
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   8655
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   15266
      _cy             =   9975
   End
End
Attribute VB_Name = "video"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With cd
.FileName = ""
.Filter = "Image(*.mp3;*.mp4;*.wav;*.wma)|*.mp3;*.mp4;*.wav;*.wma"
.ShowOpen

If Len(.FileName) <> 0 Then
strpic = .FileName
wmp.URL = cd.FileName
End If
End With
End Sub

Private Sub Command3_Click()
wmp.URL = cd.FileName

wmp.Visible = True
End Sub

Private Sub Command4_Click()
gallery.Show
Me.Hide
End Sub

Private Sub Command6_Click()
End
End Sub
