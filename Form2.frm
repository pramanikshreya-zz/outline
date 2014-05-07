VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form gallery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12765
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   7155
   ScaleWidth      =   12765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Height          =   735
      Left            =   11520
      MouseIcon       =   "Form2.frx":13D0E
      MousePointer    =   99  'Custom
      Picture         =   "Form2.frx":14298
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7800
      MouseIcon       =   "Form2.frx":16E1B
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Print"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9840
      Top             =   2400
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   9340
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   2520
      Top             =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   540
      Left            =   3600
      TabIndex        =   5
      Top             =   5520
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
      Left            =   11280
      MouseIcon       =   "Form2.frx":173A5
      MousePointer    =   99  'Custom
      Picture         =   "Form2.frx":1792F
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Go to Music"
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Height          =   735
      Left            =   10440
      MouseIcon       =   "Form2.frx":1A1A8
      MousePointer    =   99  'Custom
      Picture         =   "Form2.frx":1A732
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Back"
      Top             =   6360
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
      Height          =   555
      Left            =   5160
      MouseIcon       =   "Form2.frx":1CEEA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print"
      Top             =   360
      Width           =   2055
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
      Height          =   555
      Left            =   6720
      MouseIcon       =   "Form2.frx":1D474
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Save "
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
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
      Height          =   555
      Left            =   2640
      MouseIcon       =   "Form2.frx":1D9FE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Open"
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   240
      TabIndex        =   7
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   3975
      Left            =   3600
      MouseIcon       =   "Form2.frx":1DF88
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      ToolTipText     =   "Click to zoom"
      Top             =   1320
      Width           =   5895
   End
End
Attribute VB_Name = "gallery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strpic As String
Dim conn As New ADODB.Connection
Dim i As Integer
Dim flag As Integer




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
'cd.ShowSave
'SavePicture Image1.Picture, cd.FileName
End With
Image1.Visible = True
Image1.Height = 3975
Image1.Width = 5895
Image1.Left = 3600
Image1.Top = 1320
Image1.ToolTipText = "Click to zoom"
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Set gallery.Picture = LoadPicture(App.Path & "\images\galopen.jpg")
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
 If Text1.Text = "" Then
 MsgBox "Name cannot be empty"
 Exit Sub
 Else
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
 Me.filllistview
 Dim filen As String
 filen = Text1.Text
 cd.FileName = App.Path & "\picture\" & filen & ".jpg"
SavePicture Image1.Picture, cd.FileName
 
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Set gallery.Picture = LoadPicture(App.Path & "\images\galsave.jpg")
    
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
    'rs.CursorType = adOpenDynamic
    'rs.CursorLocation = adUseClient
    'rs.LockType = adLockOptimistic
rs.Open "select * from Table1 where name ='" & Text1.Text & "'", conn, 3, 2
If Not rs.EOF Then
Set DataReport1.DataSource = rs

 Dim filen As String
 filen = Text1.Text
Set DataReport1.Sections("section1").Controls.Item("pic").Picture = LoadPicture("" & App.Path & "\picture\" & filen & ".jpg")
DataReport1.Show
'Set DataReport1.Sections("section1").Controls.Item("pic").Picture = LoadPicture("" & App.Path & "\picture\a.jpg")

End If
End Sub

Private Sub Command4_Click()
Me.Hide
afterlog.Show
Timer1.Enabled = False

End Sub
Private Sub Command5_Click()
Me.Hide
mp3.Show
Timer1.Enabled = False
End Sub

Private Sub Command6_Click()
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\database1.mdb;Persist Security Info=False"
conn.Open
conn.Execute "delete * from Table1 where name='" & Text2.Text & "' and username='" & login.Text1.Text & "'"
conn.Close
Me.filllistview
End Sub

Private Sub Command7_Click()
End
End Sub

Private Sub Form_Load()
Me.filllistview
Image1.Visible = False
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Set gallery.Picture = LoadPicture(App.Path & "\images\galprint.jpg")
    
    If Me.WindowState <> vbMinimized Then
        'gallery.Picture.Width = Me.Width
        'gallery.Picture.Height = Me.Height
    End If
End Sub

Private Sub Image1_Click()
'With cd
'.FileName = ""
'.Filter = "Image(*.jpg;*.bmp;*.png;*.gif)|*.jpg;*.bmp;*.png;*.gif"
'.ShowOpen

'If Len(.FileName) <> 0 Then
'strpic = .FileName
'Image1.Picture = LoadPicture(.FileName)
'End If

'End With
If flag = 0 Then
Image1.Height = 6615
Image1.Width = 9615
Image1.Left = 2520
Image1.Top = 360
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command6.Visible = False
Text1.Visible = False
Image1.ToolTipText = "click to zoom back"

flag = 1
Else
Image1.Height = 3975
Image1.Width = 5895
Image1.Left = 3600
Image1.Top = 1320
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command6.Visible = True
Text1.Visible = True
Image1.ToolTipText = "click to zoom back"

flag = 0
End If


End Sub
Sub filllistview()
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
rs.Open " select * from Table1 where username='" & login.Text1.Text & "'", conn, 3, 2
If Not rs.EOF Then
ListView1.ListItems.Clear
rs.MoveFirst
Do While Not rs.EOF
Set Item = ListView1.ListItems.Add(, , rs!Name)
rs.MoveNext
Loop
Else
ListView1.ListItems.Clear
End If
rs.Close
Set rs = Nothing
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
Text2.Text = .Text


End With
Me.Text2_Change

End Sub


Private Sub Picture1_Click()

End Sub

Private Sub Text1_Change()
Me.filllistview
End Sub

Private Sub Timer1_Timer()

Dim j As Integer

i = i + 1

j = i Mod 10
Select Case j
Case 1
Set gallery.Picture = LoadPicture(App.Path & "\images\galopen.jpg")
Case 4
Set gallery.Picture = LoadPicture(App.Path & "\images\galsave.jpg")

Case 7
Set gallery.Picture = LoadPicture(App.Path & "\images\galprint.jpg")

End Select

End Sub
Sub Text2_Change()
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
    'rs.CursorType = adOpenDynamic
    'rs.CursorLocation = adUseClient
    'rs.LockType = adLockOptimistic
rs.Open "select * from Table1 where name='" & Text2.Text & "'", conn, 3, 2
If Not rs.EOF Then
Set picstrm = New ADODB.Stream
picstrm.Type = adTypeBinary
picstrm.Open
If IsNull(rs!Picture) = False Then
picstrm.Write rs!Picture
picstrm.SaveToFile "" & App.Path & "\picture\a.jpg", adSaveCreateOverWrite
Image1.Picture = LoadPicture("" & App.Path & "\picture\a.jpg")
picstrm.Close
Set picstrm = Nothing


End If
End If
rs.Close
Set rs = Nothing
'Dim filen As String
 'filen = Text1.Text
 'cd.FileName = App.Path & "\picture\" & filen & ".jpg"
'SavePicture Image1.Picture, cd.FileName
 
End Sub
