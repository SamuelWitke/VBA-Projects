Option Explicit

Public Property Get Name() As String
End Property

Public Property Get Rating() As Byte
End Property

Public Function CalculatePricing(ByVal info As IPricingRuleInfo) As Currency
End Function

Public Function GetDateType(ByVal value As Date) As DateType
End Function

