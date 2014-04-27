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
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   3720
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   1440
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
      BackColor       =   &H80000018&
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
Dim conn As ADODB.Connection
Dim login As ADODB.Recordset

Private Sub Command1_Click()
Timer1.Enabled = False

If Text1.Text = "" Then
    MessageBar.Caption = "ERROR: Username cannot be empty!"
    'MessageBar.Visible = True
    Text1.SetFocus
    Timer1.Enabled = True
    Exit Sub
ElseIf Text2.Text = "" Then
    MessageBar.Caption = "ERROR: Password cannot be empty!"
    'MessageBar.Visible = True
    Text2.SetFocus
    Timer1.Enabled = True
    Exit Sub
Else
StartConn:
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\database1.mdb;Persist Security Info=False"
    conn.CursorLocation = adUseClient
    conn.Open
    If Not conn.State = adStateOpen Then
        Select Case MsgBox("There was an error opening the databse! Please exit and restart the program. Alternately, you can try to connect again.", vbCritical + vbApplicationModal + vbRetryCancel + vbDefaultButton1, "Database Error")
        Case vbRetry
            GoTo StartConn
        Case vbCancel
            End
        End Select
    End If
    
    Set login = New ADODB.Recordset
    login.CursorType = adOpenDynamic
    login.CursorLocation = adUseClient
    login.LockType = adLockOptimistic
    login.Open "Select * from Users where UserName='" & Text1.Text & "'", conn, login.CursorType, login.LockType, adCmdUnknown
    
    If login.EOF Then
        MessageBar.Caption = "ERROR: No such user exists! Please check for spelling errors."
        'MessageBar.Visible = True
        txtUsername.SetFocus
  
        Exit Sub
    Else
        If login.Fields("Password") = Text2.Text Then
            Me.Hide
            afterlog.Show
        Else
            MessageBar.Caption = "ERROR: Wrong password! Please check for spelling/capitalization errors."
            'MessageBar.Visible = True
            Text2.SetFocus
         
            Exit Sub
        End If
    End If
End If
End Sub

Private Sub Command2_Click()
End
End Sub


