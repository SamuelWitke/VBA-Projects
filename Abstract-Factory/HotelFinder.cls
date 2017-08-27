Option Explicit

Private Type TFinder
    Hotels As Collection
End Type
Private this As TFinder

Public Property Get Hotels() As Collection
    Set Hotels = this.Hotels
End Property

Public Function FindCheapestHotel(ByVal fromDate As Date, ByVal toDate As Date, ByVal custType As CustomerType) As String

    Dim place As IHotel
    Dim checkedDate As Date

    Dim cheapestAmount As Currency
    Dim cheapestHotel As IHotel
 
    Dim hotelTotal As Currency
    For Each place In this.Hotels
 
        hotelTotal = 0
        For checkedDate = fromDate To toDate
            Dim info As New PricingRuleInfo
        
            Set info = info.Create(place.GetDateType(checkedDate), custType)
            hotelTotal = hotelTotal + place.CalculatePricing(info)
        Next
 
        If cheapestAmount = 0 Or hotelTotal < cheapestAmount Then
            cheapestAmount = hotelTotal
            Set cheapestHotel = place
        ElseIf hotelTotal = cheapestAmount And cheapestHotel.Rating > place.Rating Then
            'same price, but higher rating; higher rating gets precedence
            Set cheapestHotel = place
        End If
 
        Debug.Print place.Name, Format(hotelTotal, "$#,##0.00"), place.Rating, custType
        
    Next
 
    FindCheapestHotel = cheapestHotel.Name

End Function

Private Sub Class_Initialize()
    Set this.Hotels = New Collection
End Sub

Private Sub Class_Terminate()
    Set this.Hotels = Nothing
End Sub

