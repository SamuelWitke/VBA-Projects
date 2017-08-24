
Public Function CreateEmployee(name As String, age As Integer) As Employee
    Set CreateEmployee = New Employee
    CreateEmployee.InitiateProperties name:=name, age:=age
End Function


