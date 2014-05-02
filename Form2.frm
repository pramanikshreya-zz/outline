VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
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
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   660
      Left            =   5520
      TabIndex        =   5
      Top             =   3840
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1200
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
      Height          =   795
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SAVE AS"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   2535
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
      Height          =   795
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "gallery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strpic As String
Dim conn As New ADODB.Connection
'Dim rs As New ADODB.Recordset
'conn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;data source=" & App.name & "\db\Database1.mdb;"
'conn.Open

Private Sub Command1_Click()
With cd
.FileName = ""
.Filter = "Image(*.jpg;*.bmp;*.png;*.gif)|*.jpg;*.bmp;*.png;*.gif"
.ShowOpen

If Len(.FileName) <> 0 Then
strpic = .FileName
Image1.Picture = LoadPicture(.FileName)
'Image1.Height = 2415
'Image1.Width = 3255
End If
'---
cd.ShowSave
SavePicture Image1.Picture, cd.FileName
End With
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set gallery.Picture = LoadPicture(App.Path & "\images\galopen.jpg")
    
    If Me.WindowState <> vbMinimized Then
        'gallery.Picture.Width = Me.Width
        'gallery.Picture.Height = Me.Height
    End If
End Sub

Private Sub Command2_Click()
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
    rs.Open "Select * from Table1 where username='" & login.Text1.Text & "'", conn, rs.CursorType, rs.LockType, adCmdUnknown
    'where username='" & login.Text1.Text & "'",
    'rs.execute "INSERT INTO Table1" VALUES (value1,value2);
   'conn.Execute ("insert into " & login.Text1.Text & "(PhotoName, Photoname) values('" & Text1.Text & "','" & File1.name & "')")
    
 rs.AddNew
 rs!UserName = login.Text1.Text
 rs!Name = Text1.Text
 If strpic <> "" Then
 Set picstrm = New ADODB.Stream
 picstrm.Type = adTypeBinary
 picstrm.Open
 picstrm.LoadFromFile strpic
 rs!Picture = picstrm.Read
 picstrm.Close
 Set picstrm = Nothing
 End If
 rs.Update
 rs.Close
 Set rs = Nothing
 MsgBox "saved"
 Dim filen As String
 filen = Text1.Text
 cd.FileName = App.Path & "\picture\" & filen & ".jpg"
SavePicture Image1.Picture, cd.FileName
 
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set gallery.Picture = LoadPicture(App.Path & "\images\galsave.jpg")
    
    If Me.WindowState <> vbMinimized Then
        'gallery.Picture.Width = Me.Width
        'gallery.Picture.Height = Me.Height
    End If
End Sub

Private Sub Command3_Click()
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
rs.Open "select * from Table1 where name ='" & Text1.Text & "'", conn, 3, 2
If Not rs.EOF Then
Set DataReport1.DataSource = rs

 Dim filen As String
 filen = Text1.Text
Set DataReport1.Sections("section1").Controls.Item("pic").Picture = LoadPicture("" & App.Path & "\picture\" & filen & ".jpg")
DataReport1.Show


End If
End Sub

Private Sub Command4_Click()
Me.Hide
afterlog.Show

End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Command4.= "Back to previous menu"
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
Image1.Picture = LoadPicture(.FileName)
End If

End With
End Sub

Private Sub Picture1_Click()

End Sub

