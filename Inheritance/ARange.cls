Option Explicit

Private mclsRange As Range

Public Property Set Range(ByVal clsRange As Range): Set mclsRange = clsRange: End Property
Public Property Get Range() As Range: Set Range = mclsRange: End Property

Public Property Get RngName() As String

    RngName = Me.Range.RngName & " A Range"

End Property

Public Property Let RngName(sName As String)

    Me.Range.RngName = sName

End Property

Private Sub Class_Initialize()

    Set mclsRange = New Range

End Sub

Public Sub ParseData()
    Debug.Print "Parsing Data the A range way"
End Sub

Private Sub Class_Terminate()

    Set mclsRange = Nothing

End Sub
