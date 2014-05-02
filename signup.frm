VERSION 5.00
Begin VB.Form signup 
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MouseIcon       =   "signup.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "signup.frx":058A
   ScaleHeight     =   7245
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   4920
      Width           =   975
   End
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
      Left            =   8280
      MouseIcon       =   "signup.frx":9E41
      MousePointer    =   99  'Custom
      Picture         =   "signup.frx":A3CB
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Text            =   "Mobile no....."
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Text            =   "Telephone No....."
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Text            =   "Pincode...."
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Text            =   "Address...."
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Text            =   "Confirm Email-id......"
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Text            =   "Email id...."
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Text            =   "Password...."
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000016&
      Caption         =   "Sign Up now!"
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
      Left            =   5880
      MouseIcon       =   "signup.frx":C6D2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Text            =   "Username...."
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   1680
      Width           =   5295
   End
End
Attribute VB_Name = "signup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
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
ElseIf Text9.Text <> "5" Then
Label1.Caption = "You sure know how to add right?"
Text9.SetFocus
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
Set rs = New ADODB.Recordset
    rs.CursorType = adOpenDynamic
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.Open "Select * from Users", conn, rs.CursorType, rs.LockType, adCmdUnknown
    'where username='" & login.Text1.Text & "'",
    'rs.execute "INSERT INTO Table1" VALUES (value1,value2);
   'conn.Execute ("insert into " & login.Text1.Text & "(PhotoName, Photoname) values('" & Text1.Text & "','" & File1.name & "')")
    
 rs.AddNew
 rs!UserName = Text1.Text
 rs!Password = Text2.Text
 
 rs.Update
 rs.Close
 Set rs = Nothing
 MsgBox "You're signed up!"
 Print "done"
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Text1_Click()
Text1.Text = ""
Text1.SetFocus
End Sub


Private Sub Text2_Click()
Text2.Text = ""
Text2.SetFocus
End Sub

Private Sub Text3_Click()
Text3.Text = ""
Text3.SetFocus
End Sub

Private Sub Text9_Change()
Label1.Caption = ""
End Sub
