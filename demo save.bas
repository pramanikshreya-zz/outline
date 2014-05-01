Attribute VB_Name = "Module2"
Option Explicit
Public strConnString As String

Public Sub AddImageToDB(ByVal strFile As String, ByVal ID As Integer, ByVal Description As String)
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strStream As ADODB.Stream

    Set cn = New ADODB.Connection
    cn.ConnectionString = strConnString
    cn.Open
    
   
    'Add the image to the database
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
    strStream.LoadFromFile strFile


    Set rs = New ADODB.Recordset
    With rs
        .ActiveConnection = cn
        .Source = "SELECT ID, Picture, Description FROM tblImages"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    
    rs.AddNew
    
    rs.Fields("ID").Value = ID
    rs.Fields("Description").Value = Description
    rs.Fields("Picture").Value = strStream.Read
    rs.Update

    rs.Close

    'Cleanup
    strStream.Close
    Set strStream = Nothing
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
End Sub

Public Function ViewFromDB(ByVal ID As String, ByVal TempPath As String) As Boolean
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strStream As ADODB.Stream
Dim strSQL As String

    Set cn = New ADODB.Connection
    cn.ConnectionString = strConnString
    cn.Open
    
    strSQL = "SELECT Picture, Description " & _
                "FROM tblImages " & _
                "WHERE ID = " & ID
                
    Set rs = New ADODB.Recordset
    With rs
        .ActiveConnection = cn
        .Source = strSQL
        .Open
    End With
    
    If Not (rs.BOF And rs.EOF) Then
        Set strStream = New ADODB.Stream
        strStream.Type = adTypeBinary
        strStream.Open
    
        strStream.Write rs!Picture
    
        strStream.SaveToFile TempPath, adSaveCreateOverWrite
        
        strStream.Close
        Set strStream = Nothing
        
        ViewFromDB = True
    End If
    
    rs.Close
    Set rs = Nothing
    
    cn.Close
    Set cn = Nothing
    
End Function

