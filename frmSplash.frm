VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5790
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   5790
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5280
      Top             =   2760
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   3840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      MousePointer    =   11
      Scrolling       =   1
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   $"frmSplash.frx":5F53
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   4
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   855
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "All rights reserved."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Starting...."
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   3360
      Width           =   3615
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Timer1_Timer()
Dim i As Integer
If ProgressBar1.Value >= ProgressBar1.Max Then
login.Show
Unload Me
Timer1.Enabled = False
End If

i = ProgressBar1.Value
ProgressBar1.Value = ProgressBar1.Value + 1

Select Case i
Case 1
Label1.Caption = "Loading Forms..."
'Print i
Case 19
Label1.Caption = "Connecting Database.."
'Print i

Case 29
Label1.Caption = "Preparing User Interface..."
'Print i

Case 49
Label1.Caption = "Checking Connectivity..."
Case 69
Label1.Caption = "Preparing Accounts Info.."
Case 95
Label3.Caption = "Welcome.."

End Select

End Sub

