Private Type THotel
    PricingRules As New Scripting.Dictionary
    Name As String
    Rating As Byte
End Type
Private this As THotel
Implements IHotel

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(ByVal value As String)
    this.Name = value
End Property

Public Property Get Rating() As Byte
    Rating = this.Rating
End Property

Public Property Let Rating(ByVal value As Byte)
    this.Rating = value
End Property

Public Property Get Self() As IHotel
    Set Self = Me
End Property

Public Function Create(ByVal hotelName As String, ByVal stars As Byte, Optional ByVal rules As Collection = Nothing) As StandardHotel
 
    Dim rule As IPricingRule
    With New StandardHotel
 
        .Name = hotelName
        .Rating = stars
 
        If Not rules Is Nothing Then
            For Each rule In rules
                .AddPricingRule rule
            Next
        End If
 
        Set Create = .Self
 
    End With

End Function

Public Sub AddPricingRule(ByVal rule As IPricingRule)
    this.PricingRules.Add rule.RuleInfo.ToString, rule
End Sub

Private Function IHotel_CalculatePricing(ByVal info As IPricingRuleInfo) As Currency
    Dim rule As IPricingRule
    Set rule = this.PricingRules(info.ToString)
    IHotel_CalculatePricing = rule.Evaluate(info)
End Function

Private Function IHotel_GetDateType(ByVal value As Date) As DateType
    IHotel_GetDateType = IIf(Weekday(value, vbMonday) <= 5, WkDay, WkEnd)
End Function

Private Property Get IHotel_Name() As String
    IHotel_Name = this.Name
End Property

Private Property Get IHotel_Rating() As Byte
    IHotel_Rating = this.Rating
End Property
