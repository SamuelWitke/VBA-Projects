Option Explicit

Public Enum CustomerType
    Regular
    Premium
End Enum

Public Enum DateType
    WkDay
    WkEnd
End Enum

Public Property Get DateType() As DateType
End Property

Public Property Get CustomerType() As CustomerType
End Property

Public Function ToString() As String
End Function


