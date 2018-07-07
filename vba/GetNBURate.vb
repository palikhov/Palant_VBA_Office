Function GetNBURate(ByVal CurrencyName As String, ByVal RateDate As Date) As Double
On Error Resume Next
Dim CurrencyRate As Double
CurrencyName = UCase(CurrencyName): If Len(CurrencyName) <> 3 Then Exit Function
Set xmldoc = CreateObject("Msxml.DOMDocument"): xmldoc.async = False
url_request = "http://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?date=" & Format(RateDate, "yyyymmdd")
If xmldoc.Load(url_request) <> True Then Exit Function ' Çàïðîñ ê ñåðâåðó
Set nodeList = xmldoc.SelectNodes("/exchange/currency") 
 For i = 0 To nodeList.Length - 1
Set xmlNode = nodeList.Item(i).CloneNode(True)
If xmlNode.ChildNodes(3).Text = CurrencyName Then
CurrencyRate = Val(Replace(xmlNode.ChildNodes(2).Text, ",", "."))
GetNBURate = CurrencyRate
Exit Function
End If
Next
End Function
