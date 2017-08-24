Option Explicit

Private msRngName As String

Public Property Let RngName(ByVal sRngName As String): msRngName = sRngName: End Property
Public Property Get RngName() As String: RngName = msRngName: End Property

