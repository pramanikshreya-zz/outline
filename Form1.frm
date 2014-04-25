VERSION 5.00
Begin VB.Form login 
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6630
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   5160
      MouseIcon       =   "Form1.frx":4B64
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":50EE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DisabledPicture =   "Form1.frx":7F9B
      Height          =   615
      Left            =   3000
      MouseIcon       =   "Form1.frx":8C4B
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":91D5
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   1935
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
