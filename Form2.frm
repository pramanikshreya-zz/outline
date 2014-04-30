VERSION 5.00
Begin VB.Form gallery 
   Caption         =   "Form2"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13500
   BeginProperty Font 
      Name            =   "Segoe Print"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   6090
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Height          =   735
      Left            =   1560
      Picture         =   "Form2.frx":218A4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Height          =   735
      Left            =   720
      Picture         =   "Form2.frx":2411D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   915
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   2160
      Top             =   1440
      Width           =   2655
   End
End
Attribute VB_Name = "gallery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set gallery.Picture = LoadPicture(App.Path & "\images\galopen.jpg")
    
    If Me.WindowState <> vbMinimized Then
        'gallery.Picture.Width = Me.Width
        'gallery.Picture.Height = Me.Height
    End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set gallery.Picture = LoadPicture(App.Path & "\images\galsave.jpg")
    
    If Me.WindowState <> vbMinimized Then
        'gallery.Picture.Width = Me.Width
        'gallery.Picture.Height = Me.Height
    End If
End Sub

Private Sub Command4_Click()
Me.Hide
afterlog.Show

End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4. = "Back to previous menu"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set gallery.Picture = LoadPicture(App.Path & "\images\gallery 2.jpg")
    'gallery.AutoRedraw=True
    'me.PaintPicture me.Picture,0,0,me.Width,me.Height
    
    If Me.WindowState <> vbMinimized Then
        'gallery.Picture.Width = Me.Width
        'gallery.Picture.Height = Me.Height
    End If
End Sub
Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set gallery.Picture = LoadPicture(App.Path & "\images\galprint.jpg")
    
    If Me.WindowState <> vbMinimized Then
        'gallery.Picture.Width = Me.Width
        'gallery.Picture.Height = Me.Height
    End If
End Sub

Private Sub Image1_Click()
With cd
.FileName = ""
.Filter = "Image(*.jpg;*.bmp;*.png;*.gif)|*.jpg;*.bmp;*.png;*.gif"
.ShowOpen

If Len(.FileName) <> 0 Then
strpic = .FileName
Image1.Picture = LoadPictureGDIPlus(.FileName)
End If

End With
End Sub
