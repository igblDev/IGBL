  <%
	Set objRSteamsSession  = Server.CreateObject("ADODB.RecordSet")
	
	objRSteamsSession.Open "SELECT * FROM qryTeams WHERE (((qryTeams.OwnerID)=" & ownerid & "))", objConn
	w_tm_rank = objRSteamsSession.Fields("rank").value
	%>