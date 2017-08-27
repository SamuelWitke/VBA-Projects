Public Sub ImportData()
	Dim ws As Worksheet
	Dim qt As QueryTable
	Dim URL As String
	URL = "http://online.wsj.com/mdc/public/page/2_3024-NYSE.html"
	Set ws = Worksheets.Add
	Set qt = ws.QueryTables.Add( _
    	Connection:="URL;" & URL, _
    	Destination:=Range("A1"))
	With qt
    	.RefreshOnFileOpen = True
    	.Name = "Import Data"
    	.WebFormatting = xlWebFormattingRTF
    	.WebSelectionType = xlAllTables
    	.Refresh
	End With	
End Sub


Public Sub UpdateXTr()
	Dim qt As QueryTable
	Dim URL As String
	If wsRates.Range("B1").Value = "" Then
    	MsgBox "You must choose a currency!", vbExclamation
    	Exit Sub
	End If
	
	
	URL = "http://www.x-rates.com/table/?from=" & _
    	wsRates.Range("B1").Value & _
    	"&amount=" & wsRates.Range("B2").Value
	
	Set qt = wsRates.QueryTables("XRates Data")
	
	With qt
    	.Name = "XRates Data"
    	.WebFormatting = xlWebFormattingRTF
    	.WebSelectionType = xlSpecifiedTables
    	.WebTables = "1"
    	.Connection = "URL;" & URL
    	.Refresh
	End With
	
End Sub

