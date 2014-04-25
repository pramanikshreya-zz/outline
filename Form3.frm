VERSION 5.00
Begin VB.Form afterlog 
   Caption         =   "Form3"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15360
   LinkTopic       =   "Form3"
   ScaleHeight     =   8310
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   855
      Left            =   12240
      TabIndex        =   2
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   855
      Left            =   6480
      TabIndex        =   1
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1800
      TabIndex        =   0
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   8115
      Left            =   120
      Picture         =   "Form3.frx":0000
      Top             =   0
      Width           =   15600
   End
End
Attribute VB_Name = "afterlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_resize()

    Set Image1.Picture = LoadPicture("C:\Users\SHREYA-lappy\Desktop\outline\instagram41.jpg")
    
    If Me.WindowState <> vbMinimized Then
        Image1.Width = Me.Width
        Image1.Height = Me.Height
    End If

End Sub

