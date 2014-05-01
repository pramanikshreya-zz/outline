VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Command2"
      Height          =   555
      Left            =   1320
      TabIndex        =   1
      Top             =   1920
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
    On Error GoTo cmdSave_Error
    With CommonDialog1
        .CancelError = True
        .Filter = "Image Files (*.gif; *.bmp; *.jpg)| *.gif;*.bmp;*.jpg"
        .ShowOpen
    End With
    
    AddImageToDB CommonDialog1.FileName, 1, "File added to database"
    
Exit Sub
cmdSave_Error:
End Sub

Private Sub cmdView_Click()
Dim strTempPath As String
Dim strTempName As String
Dim strTempFile As String
Dim blnShow As Boolean

    'Create a temp file name
    strTempPath = IIf(Right(AppPath, 1) = "\", App.Path, App.Path & "\")
    strTempName = Format(Now, "MMDDYYHHNNSS") & ".bmp"
    strTempFile = strTempPath & strTempName
    
    blnShow = ViewFromDB(1, strTempFile)
    
    If blnShow Then
        Picture1.Picture = LoadPicture(strTempFile)
        DoEvents
        Kill (strTempFile)
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    'Set the connectionstring to your database
    strConnString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ImageDatabase;Data Source=MARKPC"
End Sub

