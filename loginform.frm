VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00400000&
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "loginform.frx":0000
   ScaleHeight     =   6630
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000018&
      Caption         =   "Cancel"
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
      Left            =   5160
      MouseIcon       =   "loginform.frx":4B64
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000018&
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
      Left            =   3240
      MouseIcon       =   "loginform.frx":50EE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1695
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
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   1560
      Width           =   2895
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
      Height          =   735
      Left            =   720
      TabIndex        =   4
      Top             =   5520
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
    Label1.Caption = "ERROR: Username cannot be empty!"
    'MessageBar.Visible = True
    Text1.SetFocus
    'Timer1.Enabled = True
    Exit Sub
ElseIf Text2.Text = "" Then
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
        'MessageBar.Visible = True
        Text1.SetFocus
  
        Exit Sub
    Else
        If login.Fields("Password") = Text2.Text Then
            Me.Hide
            afterlog.Show
            Exit Sub
            
        Else
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

'CHANGED!!!!!!!!! FROM HERE!!!!!

'Private Sub Form_KeyPress(keyascii As Integer)
'If keyascii = 13 Then
'If Text1.SetFocus = True Then
'Text2.SetFocus
'End If
'If Text2.SetFocus = True Then
'Text2.SetFocus = False
'call Command1_Click
'End If
'End If
'End Sub






