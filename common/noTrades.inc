	<%
	Dim objRSMe
	Set objRSMe   = Server.CreateObject("ADODB.RecordSet")
	objRSMe.Open "SELECT * FROM tblOwners where ownerID = "&ownerid&" ", objConn,3,3,1 
	myTradeInd = objRSMe.Fields("acceptTradeOffers").Value
	'Response.Write "MY TRADE IND = "&myTradeInd&".<br>"
	%>