Option Explicit
Sub main()
    Test Now, Now + 3, Regular
End Sub
Sub Test(ByVal checkin As Date, ByVal checkout As Date, ByVal custType As CustomerType)
    Dim finder As New HotelFinder
    InitializeHotels finder
    Debug.Print finder.FindCheapestHotel(checkin, checkout, custType)
End Sub

Private Sub InitializeHotels(ByVal finder As HotelFinder)
    Dim StandardHotel As New StandardHotel
    Dim PricingRuleInfo As New PricingRuleInfo
    Dim FixedAmountPricingRule As New FixedAmountPricingRule
   With StandardHotel.Create("Green Valley", 3)
        .AddPricingRule FixedAmountPricingRule.Create(PricingRuleInfo.Create(WkDay, Premium), 800)
        .AddPricingRule FixedAmountPricingRule.Create(PricingRuleInfo.Create(WkEnd, Premium), 800)
        .AddPricingRule FixedAmountPricingRule.Create(PricingRuleInfo.Create(WkDay, Regular), 1100)
        .AddPricingRule FixedAmountPricingRule.Create(PricingRuleInfo.Create(WkEnd, Regular), 900)
        finder.Hotels.Add .Self
    End With
 
    With StandardHotel.Create("Red River", 4)
        .AddPricingRule FixedAmountPricingRule.Create(PricingRuleInfo.Create(WkDay, Premium), 1100)
        .AddPricingRule FixedAmountPricingRule.Create(PricingRuleInfo.Create(WkEnd, Premium), 500)
        .AddPricingRule FixedAmountPricingRule.Create(PricingRuleInfo.Create(WkDay, Regular), 1600)
        .AddPricingRule FixedAmountPricingRule.Create(PricingRuleInfo.Create(WkEnd, Regular), 600)
        finder.Hotels.Add .Self
    End With
 
    With StandardHotel.Create("Blue Hills", 5)
        .AddPricingRule FixedAmountPricingRule.Create(PricingRuleInfo.Create(WkDay, Premium), 1000)
        .AddPricingRule FixedAmountPricingRule.Create(PricingRuleInfo.Create(WkEnd, Premium), 400)
        .AddPricingRule FixedAmountPricingRule.Create(PricingRuleInfo.Create(WkDay, Regular), 2200)
        .AddPricingRule FixedAmountPricingRule.Create(PricingRuleInfo.Create(WkEnd, Regular), 1500)
        finder.Hotels.Add .Self
    End With
End Sub

