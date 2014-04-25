VERSION 5.00
Begin VB.Form gallery 
   Caption         =   "Form2"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15615
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   5625
   ScaleWidth      =   15615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Height          =   855
      Left            =   10800
      Picture         =   "Form2.frx":102E7
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Height          =   855
      Left            =   6240
      Picture         =   "Form2.frx":13370
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   2280
      Picture         =   "Form2.frx":16478
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   2055
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set gallery.Picture = LoadPicture(App.Path & "\images\PhotoGalleryHeader2.jpg")
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
