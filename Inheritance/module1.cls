Option Explicit
Sub runIt()
'    Dim this_Employee As Employee
'    Set this_Employee = EmployeeFactory.CreateEmployee(name:="sam", age:="21")
'    Debug.Print this_Employee.name
    
    Dim clsChildRange As ARangeChild
    Set clsChildRange = RangeFactory.CreateRange(name:="Range1", typeName:=ARangeChild)
    Dim BChildRange As New BRangeChild
    
    'Set the RngName property of the child
    clsChildRange.RngName = "Range1"

    'Test that the parent has the property set
    Debug.Print clsChildRange.RngName
    
    clsChildRange.ParseData
    BChildRange.ParseData
    

End Sub
