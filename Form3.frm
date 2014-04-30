VERSION 5.00
Begin VB.Form afterlog 
   Caption         =   "Form3"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15360
   LinkTopic       =   "Form3"
   MouseIcon       =   "Form3.frx":0000
   ScaleHeight     =   8310
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
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
      Left            =   840
      MouseIcon       =   "Form3.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "Form3.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "VIDEOS"
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
      Left            =   12240
      MaskColor       =   &H000040C0&
      MouseIcon       =   "Form3.frx":3697
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "MUSIC"
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
      Left            =   6480
      MouseIcon       =   "Form3.frx":3C21
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "IMAGE GALLERY"
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
      Left            =   1800
      MouseIcon       =   "Form3.frx":41AB
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   8115
      Left            =   120
      Picture         =   "Form3.frx":4735
      Top             =   0
      Width           =   15600
   End
End
Attribute VB_Name = "afterlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
Me.Hide
gallery.Show
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H80FF&

End Sub



Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = &H80FF&
End Sub
Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.BackColor = &H80FF&
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
login.Visible = False
End Sub

Private Sub Form_resize()

    Set Image1.Picture = LoadPicture("C:\Users\SHREYA-lappy\Desktop\outline\images\instagram41.jpg")
    
    If Me.WindowState <> vbMinimized Then
        Image1.Width = Me.Width
        Image1.Height = Me.Height
    End If

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H80FFFF
Command2.BackColor = &H80FFFF
Command3.BackColor = &H80FFFF


End Sub
