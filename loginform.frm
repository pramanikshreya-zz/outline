VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00400000&
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9525
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "loginform.frx":0000
   ScaleHeight     =   7290
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      MouseIcon       =   "loginform.frx":89AF
      MousePointer    =   99  'Custom
      Picture         =   "loginform.frx":8F39
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000016&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      MouseIcon       =   "loginform.frx":B240
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   2055
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
      Height          =   570
      IMEMode         =   3  'DISABLE
      Left            =   3000
      TabIndex        =   1
      Text            =   "password..."
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3000
      TabIndex        =   0
      Text            =   "Username..."
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Not registered? Register here!"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      MouseIcon       =   "loginform.frx":B7CA
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter username and password!"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   6855
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
'Timer1.Enabled = False

If Text1.Text = "" Then
Label1.Height = 375
    Label1.Caption = "ERROR: Username cannot be empty!"
    'MessageBar.Visible = True
    Text1.SetFocus
    'Timer1.Enabled = True
    Exit Sub
ElseIf Text2.Text = "" Then
Label1.Height = 375
    Label1.Caption = "ERROR: Password cannot be empty!"
    'MessageBar.Visible = True
    Text2.SetFocus
    'Timer1.Enabled = True
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
        Label1.Caption = "ERROR: No such user exists! Please check for spelling errors."
        Label1.Height = 735
        'MessageBar.Visible = True
        Text1.SetFocus
  
        Exit Sub
    Else
        If login.Fields("Password") = Text2.Text Then
            Me.Hide
            afterlog.Show
            Exit Sub
            
        Else
            Label1.Height = 735
            Label1.Caption = "ERROR: Wrong password! Please check for spelling/capitalization errors."
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

Private Sub Label2_Click()
Me.Hide
signup.Show
End Sub

'CHANGED!!!!!!!!! FROM HERE!!!!!

Private Sub Text2_KeyPress(a As Integer)
If a = 13 Then
Call Command1_Click
End If
End Sub




Private Sub Text1_Click()
Text1.Text = ""
Text1.SetFocus
End Sub



Private Sub Text2_Click()
Text2.Text = ""
Text2.PasswordChar = "*"
Text2.SetFocus
End Sub
'Private Sub Form_Resize()
 'Me.PaintPicture Me.Picture, 0, 0, Me.Width, Me.Height
'End Sub


