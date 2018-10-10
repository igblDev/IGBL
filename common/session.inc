	<%
	ownerid = session("ownerid")
	lastLogin = session("LastLogin")	
	
	if ownerid = "" then
    	GetAnyParameter "var_ownerid", ownerid
	end if
	
	if ownerid = "" then
		Response.Redirect("timeout.asp")
	else 
		Session.Timeout=240
	end if
	
	Set objRSteamsSession  = Server.CreateObject("ADODB.RecordSet")
	
	objRSteamsSession.Open "SELECT * FROM qryTeams WHERE (((qryTeams.OwnerID)=" & ownerid & "))", objConn
	
	'Response.Write "SESSION WAIVER BAL: " & objRSteamsSession.Fields("WaiverBal").Value		 & ".<br>"
	'Response.Write "SESSION TEAM NAME: " & objRSteamsSession("TeamName").value & ".<br>"
	w_tm_rank      = objRSteamsSession.Fields("rank").value	
	w_seasonOver   = objRSteamsSession.Fields("seasonOver").value	
	w_waiver_order = objRSteamsSession.Fields("WaiverPriority").value	
	w_WaiverBal    = objRSteamsSession.Fields("WaiverBal").Value	
	%>