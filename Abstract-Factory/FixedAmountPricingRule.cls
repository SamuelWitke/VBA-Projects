Option Explicit


Private Type TRule
    RuleInfo As IPricingRuleInfo
    Amount As Currency
End Type
Private this As TRule

Implements IPricingRule

Private Property Get IPricingRule_RuleInfo() As IPricingRuleInfo
    Set IPricingRule_RuleInfo = this.RuleInfo
End Property

Private Function IPricingRule_Evaluate(ByVal info As IPricingRuleInfo) As Currency
    IPricingRule_Evaluate = this.Amount
End Function

Public Property Get RuleInfo() As IPricingRuleInfo
    Set RuleInfo = this.RuleInfo
End Property

Public Property Set RuleInfo(ByVal value As IPricingRuleInfo)
    Set this.RuleInfo = value
End Property

Public Property Get Amount() As Currency
    Amount = this.Amount
End Property

Public Property Let Amount(ByVal value As Currency)
    this.Amount = value
End Property

Public Property Get Self() As IPricingRule
    Set Self = Me
End Property

Public Function Create(ByVal info As IPricingRuleInfo, ByVal value As Currency) As IPricingRule
    With New FixedAmountPricingRule
        Set .RuleInfo = info
        .Amount = value
        Set Create = .Self
    End With
End Function

