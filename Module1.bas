Attribute VB_Name = "Module1"
Dim cn As New ADODB.Connection
Dim ros As New ADODB.Recordset
Sub main()
cn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\Database1.mdb;"
cn.Open
gallery.Show



End Sub

