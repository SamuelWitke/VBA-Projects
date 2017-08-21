Option Explicit
Sub CopyDataFromDataBase()
    Dim conn As ADODB.Connection
    Dim data As ADODB.Recordset
    Dim field As ADODB.field
    
    Set conn = New ADODB.Connection
    Set data = New ADODB.Recordset
    
    
    'conn.ConnectionString = _
    "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\;Persist Security Info=False"

    conn.Open
    With data
        .ActiveConnection = conn
        .Source = "tblActor"
        .LockType = adLockReadOnly
        .CursorType = adOpenForwardOnly
        .Open
    End With
    Worksheets.Add
    
    For Each field In data.Fields
        ActiveCell.Value = field.Name
        ActiveCell.Offset(0, 1).Select
    Next field
    
    Range("A1").Select
    Range("A2").CopyFromRecordset data
    Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    data.Close
    conn.Close
End Sub

