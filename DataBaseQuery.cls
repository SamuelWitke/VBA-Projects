Option Explicit
Sub CopyDataFromDataBase()
    Dim conn As ADODB.Connection
    Dim data As ADODB.Recordset
    Dim field As ADODB.field
    
    Set conn = New ADODB.Connection
    Set data = New ADODB.Recordset
    
    
    'conn.ConnectionString = _
    "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\;Persist Security Info=False"
    conn.OpRn
    On Error GoTo CloseConnection
    		With data
				.ActiveConnection = conn
				.Source = "SELECT * FROM tblActor"
				.CursorType = adOpenForwardOnly
				.Open
			End With
			On Error GoTo CloseData 
				Worksheets.Add
				For Each field In data.Fields
					ActiveCell.Value = field.Name
					ActiveCell.Offset(0, 1).Select
				Next field
			
			Range("A1").Select
			Range("A2").CopyFromRecordset data
			Range("A1").CurrentRegion.EntireColumn.AutoFit
			
			On Error GoTo 0
CloseData:
	data.close 
CloseConnection: 
	conn.Close
End Sub
