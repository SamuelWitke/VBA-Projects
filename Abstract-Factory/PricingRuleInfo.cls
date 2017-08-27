Option Explicit

Private Type TInfo
    DateType As DateType
    CustomerType As CustomerType
End Type
Private this As TInfo

Implements IPricingRuleInfo

Public Property Get CustomerType() As CustomerType
    CustomerType = this.CustomerType
End Property

Public Property Let CustomerType(ByVal value As CustomerType)
    this.CustomerType = value
End Property

Public Property Get DateType() As DateType
    DateType = this.DateType
End Property

Public Property Let DateType(ByVal value As DateType)
    this.DateType = value
End Property

Public Property Get Self() As IPricingRuleInfo
    Set Self = Me
End Property

Public Function Create(ByVal dtType As DateType, ByVal custType As CustomerType) As IPricingRuleInfo
    With New PricingRuleInfo
        .DateType = dtType
        .CustomerType = custType
        Set Create = .Self
    End With
End Function

Private Property Get IPricingRuleInfo_CustomerType() As CustomerType
    IPricingRuleInfo_CustomerType = this.CustomerType
End Property

Private Property Get IPricingRuleInfo_DateType() As DateType
    IPricingRuleInfo_DateType = this.DateType
End Property

Private Function IPricingRuleInfo_ToString() As String
    IPricingRuleInfo_ToString = CStr(this.CustomerType) & ";" & CStr(this.DateType)
End Function

