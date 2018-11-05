<!-- #include file="adovbs.inc" -->

<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCID
	Session.CodePage = 1252
	Err.Clear

	strErrorUrl = ""

	Dim objConn, w_action, teamNum, objrsTest, objRSLineups, objRSStats, objRSStandings, objRSStreak, objRSHead2Head, objEmail
	Dim objRSteam, loopcnt,objParams

	Set objConn         = Server.CreateObject("ADODB.Connection")
	Set objRSLineups    = Server.CreateObject("ADODB.RecordSet")
	Set objRSPlayerNames= Server.CreateObject("ADODB.RecordSet")
	Set objRSStats      = Server.CreateObject("ADODB.RecordSet")
	set objRSStandings  = Server.CreateObject("ADODB.RecordSet")
	set objRSStreak     = Server.CreateObject("ADODB.RecordSet")
	set objRSHead2Head  = Server.CreateObject("ADODB.RecordSet")
	Set objParams       = Server.CreateObject("ADODB.RecordSet")
	Set objEmail        = Server.CreateObject("ADODB.RecordSet")
	Set objRSPower      = Server.CreateObject("ADODB.RecordSet")
	Set objRSPoints     = Server.CreateObject("ADODB.RecordSet")
	Set objrsWins       = Server.CreateObject("ADODB.RecordSet")
	Set objrsTies       = Server.CreateObject("ADODB.RecordSet")
	Set objrsWork       = Server.CreateObject("ADODB.RecordSet")

  objConn.Open Application("lineupstest_ConnectionString")

  objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source=lineupstest.mdb;" & _
	  						"Persist Security Info=False"

	%>
	<!--#include virtual="Common/session.inc"-->
	<%

	Dim objRSwaivers, w_priority, I
	Set objRSwaivers 	= Server.CreateObject("ADODB.RecordSet")

	w_action = Request.Form("action")
	'Response.Write "Value is: " & w_Action & ".<br>"

  If Request.Form("action") = "Save Form Data" Then

	 objRSLineups.Open "SELECT * FROM qry_matchupLineups_stagger", objConn,3,3,1
	 if objRSLineups.Recordcount > 0 then
		wGameDay = objRSLineups.Fields("gameday").Value
	 else
		wGameDay = date()
	 end if

	 objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'PLAYOFF_START_DATE' ",objConn
	 wPlayoffStart = objRSWork.Fields("param_date").Value
	 objRSWork.Close

	 '########### Delete all rows from newBox and Points Scored ###########
	 strSQL = "delete from newBox where gameDate = #"&wGameDay&"#"
	 objConn.Execute strSQL

	 strSQL = "delete from tbl_points_scored where gameDate = #"&wGameDay&"#"
	 objConn.Execute strSQL

	 Response.Write "Building Box for "&wGameDay&".<br>"

	 loopcnt = 0
	 wFinalBox = 1
	 wStatsNotFound = null
	 bExceptionsFound = FALSE
	 
     While Not objRSLineups.EOF
			whtot3s   = 0
			whtotRBS  = 0
			whtotASS  = 0
			whtotSTL  = 0
			whtotBLK  = 0
			whtotTOs  = 0
			whtotPTS  = 0
			whtotmp  = 0
			whtotBarps= 0

			watot3s   = 0
			watotRBS  = 0
			watotASS  = 0
			watotSTL  = 0
			watotBLK  = 0
			watotTOs  = 0
			watotPTS  = 0
			watotmp  = 0
			watotBarps= 0

			wHomeTeam     = objRSLineups.Fields("HomeTeam").Value
			wHomeTeamShort= objRSLineups.Fields("HomeTeamShort").Value
			wHomeTeamID   = objRSLineups.Fields("HomeOwner").Value
			wAwayTeam     = objRSLineups.Fields("VisitingTeam").Value
			wAwayTeamShort= objRSLineups.Fields("VisitingTeamShort").Value
			wAwayTeamID   = objRSLineups.Fields("VisitingOwner").Value
			wHomeTeamPen  = objRSLineups.Fields("HomeTeamPen").Value
			wAwayTeamPen  = objRSLineups.Fields("VisitingTeamPen").Value


		'Response.Write "******************* TOP OF LOOP   ************************** <br>"
		'Response.Write "Game Day = "&wGameDay&". Home Team = "&wHomeTeam&", HTeamPen = "&wHomeTeamPen&", AwayTeam = "&wAwayTeam&", ATeamPen = "&wAwayTeamPen&" <br>"

		wPid = objRSLineups.Fields("sCenter").Value
		wRetcd = GetStats(wPid,wFName,wLName,w3,wR,wA,wS,wB,wT,wP,wMP,wBarps,"A")

		'########### Insert to NewBox ###########
		loopcnt = loopcnt + 1
		strSQL = "insert into newbox " & _
				"(gameDate, game_Number, HomeTeam, HomeTeamOID, AwayTeam, AwayTeamOID, ap1PID, ap1First, ap1Last, " & _
				"httotals, attotals, ap13S, ap1RBS, ap1ASS, ap1STL, ap1BLK, ap1TOs, ap1PTS, ap1mp, ap1Barps,AwayTeamShort,HomeTeamShort " & _
				") " & _
				"values (#"&wGameDay&"#,"&loopcnt&",'"&wHomeTeam&"',"&wHomeTeamID&",'"&wAwayTeam&"',"&wAwayTeamID&","&wPid&",'"&wFName&"','"&wLName&"'," &  _
				"'Totals','Totals',"&w3&","&wR&","&wA&","&wS&","&wB&","&wT&","&wP&","&wMP&","&wBarps&",'"&wAwayTeamShort&"','"&wHomeTeamShort&"' " & _
				")"

	    'Response.Write "Sql FOR Insert = " & strSQL  & "<br>"
        objConn.Execute strSQL

		'***************
		'*************** 
		wPid = objRSLineups.Fields("sForward").Value
		wRetcd = GetStats(wPid,wFName,wLName,w3,wR,wA,wS,wB,wT,wP,wMP,wBarps,"A")
		strSQL = 	"update newbox " & _
					"set ap2PID = "&wPid&", ap2First = '"&wFName&"', ap2Last = '"&wLName& "', " & _
					"ap23S="&w3&",ap2RBS="&wR&",ap2ASS="&wA&",ap2STL="&wS&",ap2BLK="&wB&",ap2TOs="&wT&",ap2PTS="&wP&",ap2mp="&wMP&",ap2Barps="&wBarps&" " & _
					"where game_Number = "&loopcnt&" and gameDate = #"&wGameDay&"#"

		'Response.Write "Sql FOR sForward = " & strSQL  & "<br>"
        objConn.Execute strSQL

		'***************
		'***************
		wPid = objRSLineups.Fields("sForward2").Value
		wRetcd = GetStats(wPid,wFName,wLName,w3,wR,wA,wS,wB,wT,wP,wMP,wBarps,"A")
		strSQL = 	"update newbox " & _
					"set ap3PID = "&wPid&", ap3First = '"&wFName&"', ap3Last = '"&wLName& "', " & _
					"ap33S="&w3&",ap3RBS="&wR&",ap3ASS="&wA&",ap3STL="&wS&",ap3BLK="&wB&",ap3TOs="&wT&",ap3PTS="&wP&",ap3mp="&wMP&",ap3Barps="&wBarps&" " & _
					"where game_Number = "&loopcnt&" and gameDate = #"&wGameDay&"#"

        objConn.Execute strSQL

		'***************
		'***************
		wPid = objRSLineups.Fields("sGuard").Value
		wRetcd = GetStats(wPid,wFName,wLName,w3,wR,wA,wS,wB,wT,wP,wMP,wBarps,"A")
		strSQL = 	"update newbox " & _
					"set ap4PID = "&wPid&", ap4First = '"&wFName&"', ap4Last = '"&wLName& "', " & _
					"ap43S="&w3&",ap4RBS="&wR&",ap4ASS="&wA&",ap4STL="&wS&",ap4BLK="&wB&",ap4TOs="&wT&",ap4PTS="&wP&",ap4mp="&wMP&",ap4Barps="&wBarps&" " & _
					"where game_Number = "&loopcnt&" and gameDate = #"&wGameDay&"#"

        'Response.Write "strSQL = "&strSQL&" <br>"
        objConn.Execute strSQL

		'***************
		'***************
		wPid = objRSLineups.Fields("sGuard2").Value
        wRetcd = GetStats(wPid,wFName,wLName,w3,wR,wA,wS,wB,wT,wP,wMP,wBarps,"A")
		strSQL = 	"update newbox " & _
					"set ap5PID = "&wPid&", ap5First = '"&wFName&"', ap5Last = '"&wLName& "', " & _
					"ap53S="&w3&",ap5RBS="&wR&",ap5ASS="&wA&",ap5STL="&wS&",ap5BLK="&wB&",ap5TOs="&wT&",ap5PTS="&wP&",ap5mp="&wMP&",ap5Barps="&wBarps&" " & _
					"where game_Number = "&loopcnt&" and gameDate = #"&wGameDay&"#"

        'Response.Write "strSQL = "&strSQL&" <br>"
        objConn.Execute strSQL

		'***************
		'***************
		wPid = objRSLineups.Fields("HomeC").Value
		wRetcd = GetStats(wPid,wFName,wLName,w3,wR,wA,wS,wB,wT,wP,wMP,wBarps,"H")
		strSQL = 	"update newbox " & _
					"set hp1PID = "&wPid&", hp1First = '"&wFName&"', hp1Last = '"&wLName& "', " & _
					"hp13S="&w3&",hp1RBS="&wR&",hp1ASS="&wA&",hp1STL="&wS&",hp1BLK="&wB&",hp1TOs="&wT&",hp1PTS="&wP&",hp1mp="&wMP&",hp1Barps="&wBarps&" " & _
					"where game_Number = "&loopcnt&" and gameDate = #"&wGameDay&"#"

        'Response.Write "strSQL = "&strSQL&" <br>"
        objConn.Execute strSQL

		'***************
		'***************
		wPid = objRSLineups.Fields("HomeF").Value
		wRetcd = GetStats(wPid,wFName,wLName,w3,wR,wA,wS,wB,wT,wP,wMP,wBarps,"H")
		strSQL = 	"update newbox " & _
					"set hp2PID = "&wPid&", hp2First = '"&wFName&"', hp2Last = '"&wLName& "', " & _
					"hp23S="&w3&",hp2RBS="&wR&",hp2ASS="&wA&",hp2STL="&wS&",hp2BLK="&wB&",hp2TOs="&wT&",hp2PTS="&wP&",hp2mp="&wMP&",hp2Barps="&wBarps&" " & _
					"where game_Number = "&loopcnt&" and gameDate = #"&wGameDay&"#"

        'Response.Write "strSQL = "&strSQL&" <br>"
        objConn.Execute strSQL

		'***************
		'***************
		wPid = objRSLineups.Fields("HomeF2").Value
		wRetcd = GetStats(wPid,wFName,wLName,w3,wR,wA,wS,wB,wT,wP,wMP,wBarps,"H")
		strSQL = 	"update newbox " & _
					"set hp3PID = "&wPid&", hp3First = '"&wFName&"', hp3Last = '"&wLName& "', " & _
					"hp33S="&w3&",hp3RBS="&wR&",hp3ASS="&wA&",hp3STL="&wS&",hp3BLK="&wB&",hp3TOs="&wT&",hp3PTS="&wP&",hp3mp="&wMP&",hp3Barps="&wBarps&" " & _
					"where game_Number = "&loopcnt&" and gameDate = #"&wGameDay&"#"

        'Response.Write "strSQL = "&strSQL&" <br>"
        objConn.Execute strSQL

		'***************
		'***************
		wPid = objRSLineups.Fields("HomeG").Value
		wRetcd = GetStats(wPid,wFName,wLName,w3,wR,wA,wS,wB,wT,wP,wMP,wBarps,"H")
		strSQL = 	"update newbox " & _
					"set hp4PID = "&wPid&", hp4First = '"&wFName&"', hp4Last = '"&wLName& "', " & _
					"hp43S="&w3&",hp4RBS="&wR&",hp4ASS="&wA&",hp4STL="&wS&",hp4BLK="&wB&",hp4TOs="&wT&",hp4PTS="&wP&",hp4mp="&wMP&",hp4Barps="&wBarps&" " & _
					"where game_Number = "&loopcnt&" and gameDate = #"&wGameDay&"#"

        'Response.Write "strSQL = "&strSQL&" <br>"
        objConn.Execute strSQL

		'***************
		'***************
		wPid = objRSLineups.Fields("HomeG2").Value
		wRetcd = GetStats(wPid,wFName,wLName,w3,wR,wA,wS,wB,wT,wP,wMP,wBarps,"H")
		strSQL = "update newbox " & _
		         "set hp5PID = "&wPid&", hp5First = '"&wFName&"', hp5Last = '"&wLName& "', " & _
		         "hp53S="&w3&",hp5RBS="&wR&",hp5ASS="&wA&",hp5STL="&wS&",hp5BLK="&wB&",hp5TOs="&wT&",hp5PTS="&wP&",hp5mp="&wMP&",hp5Barps="&wBarps&"," & _
				 "htot3s="&whtot3s&",htotRBS="&whtotRBS&",htotASS="&whtotASS&",htotSTL="&whtotSTL&",htotBLK="&whtotBLK&",htotTOs="&whtotTOs&"," & _
		         "htotPTS="&whtotPTS&",htotmp="&whtotmp&",htotBarps="&whtotBarps&"," & _
				 "atot3s="&watot3s&",atotRBS="&watotRBS&",atotASS="&watotASS&",atotSTL="&watotSTL&",atotBLK="&watotBLK&",atotTOs="&watotTOs&"," & _
		         "atotPTS="&watotPTS&",atotmp="&watotmp&",atotBarps="&watotBarps&" " & _
		         "where game_Number = "&loopcnt&" and gameDate = #"&wGameDay&"#"

        'Response.Write "strSQL = "&strSQL&" <br>"
        objConn.Execute strSQL

		'***************
		'***************
		if whtotBarps > watotBarps then
			whWinLoss = "W"
			waWinLoss = "L"
		elseif watotBarps > whtotBarps then
			waWinLoss = "W"
			whWinLoss = "L"
	    elseif whtotTOs < watotTOs then
			whWinLoss = "W"
			waWinLoss = "L"
	    elseif watotTOs < whtotTOs then
			waWinLoss = "W"
			whWinLoss = "L"
		elseif whtotSTL > watotSTL then
			whWinLoss = "W"
			waWinLoss = "L"
		elseif watotSTL > whtotSTL then
			waWinLoss = "W"
			whWinLoss = "L"
		elseif whtotBLK > watotBLK then
			whWinLoss = "W"
			waWinLoss = "L"
		elseif watotBLK > whtotBLK then
			waWinLoss = "W"
			whWinLoss = "L"
		elseif whtot3s > watot3s then
			whWinLoss = "W"
			waWinLoss = "L"
		elseif watot3s > whtot3s then
			waWinLoss = "W"
			whWinLoss = "L"
		elseif whtotRBS > watotRBS then
			whWinLoss = "W"
			waWinLoss = "L"
		elseif watotRBS > whtotRBS then
			waWinLoss = "W"
			whWinLoss = "L"
		elseif whtotASS > watotASS then
			whWinLoss = "W"
			waWinLoss = "L"
		elseif watotASS > whtotASS then
			waWinLoss = "W"
			whWinLoss = "L"
		elseif whtotPTS > watotPTS then
			whWinLoss = "W"
			waWinLoss = "L"
		elseif watotPTS > whtotPTS then
			waWinLoss = "W"
			whWinLoss = "L"
		else
			whWinLoss = "W"
			waWinLoss = "L"
		end if

		'Response.Write "wGameDay = "&wGameDay&", wPlayoffStart = "&wPlayoffStart&"<br>"
		if whWinLoss = "W" and wHomeTeamPen = TRUE and wGameDay < wPlayoffStart then
			whPenalty = 1
			strSQL = "update newbox set PenOID = "&wHomeTeamID&" where game_Number = "&loopcnt&" and gameDate = #"&wGameDay&"#"
			objConn.Execute strSQL
		else
			whPenalty = 0
		end if

		if waWinLoss = "W" and wAwayTeamPen = TRUE and wGameDay < wPlayoffStart then
			waPenalty = 1
			strSQL = "update newbox set PenOID = "&wAwayTeamID&" where game_Number = "&loopcnt&" and gameDate = #"&wGameDay&"#"
			objConn.Execute strSQL
		else
			waPenalty = 0
		end if

        'strSQL = "update tbl_Points_Scored set TeamScore = "&whtotBarps&", OpponentScore = "&watotBarps&", Result = '"&whWinLoss&"', Penalty = "&whPenalty&" " & _
		'         "where HomeTeamID = " &wHomeTeamID& " and gameDate = #"&wGameDay&"#"

		strSQL = "insert into tbl_points_scored (GameDate,HomeTeamId,AwayTeamID,TeamScore,OpponentScore,Result,Penalty) " &_
		         "values (#"&wGameDay&"#,"&wHomeTeamID&","&wAwayTeamID&","&whtotBarps&","&watotBarps&",'"&whWinLoss&"',"&whPenalty&")"
		objConn.Execute strSQL

		'strSQL = "update tbl_Points_Scored set TeamScore = "&watotBarps&", OpponentScore = "&whtotBarps&", Result = '"&waWinLoss&"', Penalty = "&waPenalty&" " & _
		'         "where HomeTeamID = " &wAwayTeamID& " and gameDate = #"&wGameDay&"#"

		strSQL = "insert into tbl_points_scored (GameDate,HomeTeamId,AwayTeamID,TeamScore,OpponentScore,Result,Penalty) " &_
		         "values (#"&wGameDay&"#,"&wAwayTeamID&","&wHomeTeamID&","&watotBarps&","&whtotBarps&",'"&waWinLoss&"',"&waPenalty&")"
		objConn.Execute strSQL


		objRSLineups.MoveNext
	 Wend
	 objRSLineups.Close

     
	 strSQL = 	"update newbox set FinalBox = "&wFinalBox&" where gameDate = #"&wGameDay&"#"
     'Response.Write "strSQL = "&strSQL&" <br>"
     objConn.Execute strSQL

	 
	 if bExceptionsFound = TRUE then
		Response.Write "Statistics Not Found for the Following Players. <br> <br>"&wStatsNotFound
		'**************************************************************************
		'Send Email Notification of Players Not Matched on Line-ups table
		'**************************************************************************

		email_to     = "9729356748@mms.att.net, 4692358147@pm.sprint.com"  'Enter the email you want to send the form to
		'email_to     = "4692358147@pm.sprint.com"  'Enter the email you want to send the form to
		email_subject= "Statistics Not Found in Last5!"  'You can put whatever subject here
		host         = "mail.igbl.org"   'The mail server name. (Commonly mail.yourdomain.xyz if your mail is hosted with HostMySite)
		username     = "igbl_commish@igbl.org"  'A valid email address you have setup
		from_address ="igbl_commish@igbl.org" 'If your mail is hosted with HostMySite this has to match the email address above
		password     = "SpursRtheChamps2014" 'Password for the above email address

		reply_to     = "noReply@igbl.org"  'Enter the email you want customers to reply to
		port = "25" 'This is the default port. Try port 50 if this port gives you issues and your mail is hosted with HostMySite

		Set ObjSendMail = CreateObject("CDO.Message")
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = host
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = port
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = username
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = password
		ObjSendMail.Configuration.Fields.Update

		email_message       = "Statistics Not Found for the Following Players.  GameDay = "&wGameDay&"<br> <br>"&wStatsNotFound

	    'email_to = "fred_curry@igbl.org"  'Enter the email you want to send the form to
		ObjSendMail.To      = email_to
		ObjSendMail.Subject = email_subject
		ObjSendMail.From    = from_address
		ObjSendMail.HTMLBody= email_message
		ObjSendMail.Send
		set ObjSendMail     = Nothing

		'********************************************************************************
		'End of Send Email Notification of Players Not Matched on Line-ups table
		'********************************************************************************
	 end if
	
	'**************************************************************************
	'Send Email Notification thats Scores/Standings have been processed
	'**************************************************************************
	wEmailOwnerID = null
	wAlert        = "receiveBoxScoreAlerts"
	
	if wFinalBox = 1 then
	
	   Response.Write "<br>** FINAL BOX GENERATED ***<br><br>"
	   '*******************************************************************************
	   ' UPDATE STANDINGS FUNCTION CALLED DM-NOV-02-16
	   ' Do not update regular season standings once the playoffs start.   FC-MAR-96-17
	   '*******************************************************************************
       wRetcd = UpdateStandingsRD1()
	   wRetcd = UpdateStandingsRD2()
	   wRetcd = UpdateStandingsRD3()

       objrsWork.Open "SELECT * from Standings_RD1", objConn,3,3,1
       'if objrsWork.Recordcount > 0 then
	   '    Response.Write "Playoffs Started.  Regular Season Standings not updated. <br>"
	   'else
		    wRetcd = UpdateStandings()
	   'end if
	   
	   objrsWork.Close
	   
	   email_subject = "Final Scores/Standings Posted!"
	   email_message = "Login to Check Your Results!"
	else
	   Response.Write "<br>** PARTIAL BOX GENERATED ***<br><br>"
	   email_subject = "Partial Scores Posted!"
	   email_message = "Login to see how things are looking!"
	end if
%>
		<!--#include virtual="Common/email_league.inc"-->

<%
		'end if

  End if

  '######################################
  ' GetStats
  '   This function will be call for each player.
  '   Get First and Last Names.
  '   Get Box score information for Player.
  '   Pass information back to calling program.
  '######################################
  Function GetStats (pPID,pFName,pLName,p3s,preb,past,pstl,pblk,pto,ppoints,pMP,pbarps,phomeawayflag)

	'Response.Write "pPID = "&pPID&"<br>"

	objRSPlayerNames.Open "Select firstName,lastName,NBATeamId  from tblPlayers where PID = "& pPID , objConn,3,3,1
	pFName  = objRSPlayerNames.Fields("firstName").Value
	pLName  = objRSPlayerNames.Fields("lastName").Value
	pNbaTeamID = objRSPlayerNames.Fields("NBATeamId").Value
	objRSPlayerNames.Close

	objRSStats.Open	"SELECT * FROM tblLast5 " & _
	                "WHERE gamedate = #"&wGameDay&"# AND First = '"&pFName&"' AND Last = '"&pLName&"' order by barptot desc ", objConn,3,3,1

	'*******************************************************************
	'* New counter for determining if Bonus should be awarded to player
	'*******************************************************************
	bonusCnt = 0
	pbonus   = 0

	if objRSStats.Recordcount > 0 then
	    'Response.Write "** FOUND ** for "&pFName&" "&pLName&"<br>"
		p3s    = objRSStats.Fields("x3P").Value
		preb   = objRSStats.Fields("TRB").Value
		past   = objRSStats.Fields("AST").Value
		pstl   = objRSStats.Fields("STL").Value
		pblk   = objRSStats.Fields("BLK").Value
		pto    = objRSStats.Fields("TOV").Value
		ppoints= objRSStats.Fields("PTS").Value
		pMP    = objRSStats.Fields("MP").Value
		'Response.Write "Points From Last 5 = "&ppoints&"<br>"
		pbarps = objRSStats.Fields("BarpTot").Value

		'Response.Write pLName&", "&pbarps&"<br>"

		'pbarps = cint(pbarps)
		pbarps = cdbl(pbarps)
	else
		if pPID < 9000 then
			bExceptionsFound = TRUE
			wStatsNotFound = wStatsNotFound&pFName&" "&pLName&" - "
			
			objrsWork.Open "SELECT * FROM tblPlayers x WHERE nbaTeamId = "&pNbaTeamID&" " &_
			               "and exists " &_
			               "(select 1 from tblLast5 y " &_
			               "where x.firstName = y.First " &_
			               "and x.lastname = y.Last " &_
			               "and gamedate = #"&wGameDay&"#)", objConn,3,3,1
			
			if objrsWork.Recordcount > 0 then
			   wStatsNotFound = wStatsNotFound&" TM Found <br>" 
			else
			   wStatsNotFound = wStatsNotFound&" TM NOT Found <br>" 
			   wFinalBox = 0
			end if
			objrsWork.Close
			
		end if

		p3s    = 0
		preb   = 0
		past   = 0
		pstl   = 0
		pblk   = 0
		pto    = 0
		ppoints= 0
		pbonus = 0
		pbarps = 0
		pMP = 0


	end if
	objRSStats.Close

	if phomeawayflag = "H" then
		whtot3s   = whtot3s + p3s
		whtotRBS  = whtotRBS + preb
		whtotASS  = whtotASS + past
		whtotSTL  = whtotSTL + pstl
		whtotBLK  = whtotBLK + pblk
		whtotTOs  = whtotTOs + pto
		whtotPTS  = whtotPTS + ppoints
		whtotBon  = whtotBon + pbonus
		whtotBarps= whtotBarps + pbarps
	else
		watot3s   = watot3s + p3s
		watotRBS  = watotRBS + preb
		watotASS  = watotASS + past
		watotSTL  = watotSTL + pstl
		watotBLK  = watotBLK + pblk
		watotTOs  = watotTOs + pto
		watotPTS  = watotPTS + ppoints
		watotBon  = watotBon + pbonus
		watotBarps= watotBarps + pbarps
	end if

  End Function


  Function UpdateStandings ()

	'Response.Write "Top of updateStandings Function <br>"

	'Query has a condition to only consider regular season.  gamedate < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE')
    objRSStandings.Open "SELECT * FROM qry_update_standings", objConn

	wRank = 0
	While Not objRSStandings.EOF

		wRank = wRank + 1
		wPen    = objRSStandings.Fields("total_Pen").Value

		if wRank = 1 then
			wLeader = objRSStandings.Fields("adjusted_wins").Value
			wGB = 0
		else
			wGB = wLeader - objRSStandings.Fields("adjusted_wins").Value
		end if

		wTeamID  = objRSStandings.Fields("HomeTeamID").Value
		wWon     = objRSStandings.Fields("adjusted_wins").Value
		wLoss    = objRSStandings.Fields("adjusted_losses").Value
		wPPG     = objRSStandings.Fields("MyPPG").Value
		wOPPG    = objRSStandings.Fields("oppPPG").Value
		wDiff    = objRSStandings.Fields("MyPPG").Value - objRSStandings.Fields("oppPPG").Value

		wCurVal  = "L"
		lStreakCt = 1
		'#############
		'Get Streak
		'#############
		objRSStreak.Open "select * from tbl_points_scored " &_
		                 "where hometeamid="&wTeamID&" " &_
						 "and Result is not null " &_
		                 "and tbl_points_scored.gamedate < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') " &_
		                 "order by gamedate desc", objConn

		lLoopCt = 0
		last10Win  = 0
		last10Loss = 0
		bStreak = 1
		bLast10 = 1

		While Not objRSStreak.EOF and (bStreak = 1 or bLast10 = 1)
			lLoopCt = lLoopCt + 1

			'Streak Logic
			if bStreak = 1 then
			   if lLoopCt = 1 then
				  lStreakCt = 1
				  wCurVal = objRSStreak.Fields("Result").Value
			   elseif objRSStreak.Fields("Result").Value = wCurVal then
			      lStreakCt = lStreakCt + 1
			   else
				  bStreak = 0
			   end if
			end if

			'Last10 Logic
			if bLast10 = 1 then
			   if objRSStreak.Fields("Result").Value = "W" then
			      last10Win = last10Win + 1
			   else
			      last10Loss = last10Loss + 1
			   end if

			   if lLoopCt = 10 then
			      bLast10 = 0
			   end if
			end if

			objRSStreak.MoveNext
		Wend
		objRSStreak.Close
		'###################

		strStreak = wCurVal&lStreakCt
        strLast10 = last10Win&"-"&last10Loss
        strSQL = "update Standings " & _
		         "set Rank="&wRank&",Won="&wWon&",Loss="&wLoss&",PPG="&wPPG&",OPPG="&wOPPG&",DIFF="&wDiff&"," & _
				 "GB="&wGB&",ST='"&strStreak&"',cycle='"&strLast10&"',LP="&wPen&" "& _
		         "where ID = "&wTeamID

		objConn.Execute strSQL
		objRSStandings.MoveNext

	Wend
	objRSStandings.Close

	'####################
	'Manage Head to Head
	'####################
	'Response.Write "Checkpoint A <br>"
	strSQL = "update standings set GR=null,JP=null,DB=null,CJ=null,JW=null,MJ=null,AW=null,FC=null,TA=null,DM=null"
    objConn.Execute strSQL

	objRSHead2Head.Open	"SELECT p.hometeamid, o.OwnerInitials, count(*) as tot_wins " &_
	                    "FROM tbl_points_scored p, TBLOWNERS o " &_
                        "where p.AwayTeamID = o.ownerID and p.result = 'W' " &_
						"and p.gamedate < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') " &_
						"group by p.hometeamid, o.OwnerInitials ", objConn

	While Not objRSHead2Head.EOF
		wWorkID  = objRSHead2Head.Fields("hometeamid").Value
		wOwnerCol= objRSHead2Head.Fields("OwnerInitials").Value
		wTotWins = objRSHead2Head.Fields("tot_wins").Value

	    strSQL = "update Standings " & _
		         "set "&wOwnerCol&" = "&wTotWins&" where ID = "&wWorkID

		'Response.Write "strSQL = "&strSQL&"<br>"
		objConn.Execute strSQL

		objRSHead2Head.MoveNext
	Wend
	objRSHead2Head.Close

	'#############################
	'Calculate Power Ranking Score
	'#############################
	'Response.Write "Checkpoint B <br>"
	objRSPower.Open "SELECT * from tblowners where ownerid <> 99 order by ownerid ", objConn,1,1

	'Response.Write "xct = "&xct&"<br>"

	while not objRSPower.EOF
		wOwnerID = objRSPower.Fields("ownerID")
		wShort = objRSPower.Fields("ShortName")
		'Response.Write "**** OWNER ID = "&wOwnerID&"<br>"

		objRSPoints.Open "SELECT * from tbl_points_scored " &_
		                 "where homeTeamID = "&wOwnerID&" " &_
		                 "and tbl_points_scored.gamedate < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') " &_
		                 "and TeamScore is not null ", objConn,1,1
		wTotWins = 0
		while not objRSPoints.EOF
			wScore = objRSPoints.Fields("TeamScore")
			wDate = objRSPoints.Fields("GameDate")

		    objrsWins.Open "SELECT count(*) as wins from tbl_points_scored " &_
			               "where gamedate = #"&wDate&"# and TeamScore < "&wScore&" " &_
			               "and tbl_points_scored.gamedate < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') " &_
			               "and HomeTeamID <> "&wOwnerID, objConn,1,1
			wTotWins = wTotWins + objrsWins.Fields("Wins")
			objrsWins.Close

		    objrsTies.Open "SELECT count(*) as ties from tbl_points_scored " &_
			               "where gamedate = #"&wDate&"# " &_
						   "and TeamScore = "&wScore&" " &_
						   "and HomeTeamID <> "&wOwnerID&" " &_
			               "and tbl_points_scored.gamedate < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') " &_
			               "and HomeTeamID <> "&wOwnerID, objConn,1,1

			wTotWins = wTotWins + (objrsTies.Fields("ties") * .5)
			objrsTies.Close

			'Response.Write "   "&wDate&"  "&wScore&"<br>"
			objRSPoints.MoveNext
		wend
		objRSPoints.Close

		'Response.Write wShort&" - "&wTotWins&"<br>"
		strSQL = "update Standings set prs = "&wTotWins&" where id = "&wOwnerID
        objConn.Execute strSQL
        'Response.Write "Sql = " & strSQL  & ".<br>"

		objRSPower.MoveNext
	wend

	objRSPower.Close

	'Response.Write "Calculate Cycle Standings Function <br>"
	'#############################
	'Build Standings Cycle table
	'#############################
	strSQL = "delete from standings_cycle_work"
	objConn.Execute strSQL
	
	strSQL = "insert into standings_cycle_work select * from standings_cycle"
	objConn.Execute strSQL
	
	strSQL = "delete from standings_cycle"
    objConn.Execute strSQL

	
	wRank = 0
	prevCycle = 0
	
    objRSStandings.Open "SELECT * FROM qry_cycle_standings", objConn
	While Not objRSStandings.EOF
	    if objRSStandings.Fields("cycle").Value = prevCycle then
	       wRank = wRank + 1
		   'wLeader = objRSStandings.Fields("adjusted_wins").Value
		   'wGB = 0
		else
		   wRank = 1
           prevCycle = objRSStandings.Fields("cycle").Value
		end if

		wTeamID  = objRSStandings.Fields("HomeTeamID").Value
		wCycle   = objRSStandings.Fields("cycle").Value
		wTeam    = objRSStandings.Fields("shortname").Value
		wWon     = objRSStandings.Fields("adjusted_wins").Value
		wLoss    = objRSStandings.Fields("adjusted_losses").Value
		wPPG     = objRSStandings.Fields("MyPPG").Value
		wOPPG    = objRSStandings.Fields("oppPPG").Value
		'wDiff    = objRSStandings.Fields("MyPPG").Value - objRSStandings.Fields("oppPPG").Value

		strSQL = "insert into standings_cycle (ID,cycle,rank,team,won,loss,ppg,oppg) " &_
		         "values ("&wTeamID&","&wCycle&","&wRank&",'"&wTeam&"',"&wWon&","&wLoss&",'"&wPPG&"',"&wOPPG&")"
		  
		objConn.Execute strSQL

	    objRSStandings.MoveNext
	wend
	objRSStandings.Close
	
	objRSStandings.Open "SELECT max(cycle) as current_cycle  from standings_cycle", objConn
	wCurrentCycle = objRSStandings.Fields("current_cycle").Value
	objRSStandings.Close
	
	'Response.Write "wCurrentCycle = "&wCurrentCycle&"<br>"
	
	if wCurrentCycle > 1 then
	   objRSStandings.Open "SELECT * from standings_cycle_work where cycle < "&wCurrentCycle, objConn
	   	While Not objRSStandings.EOF
		   workCycle = objRSStandings.Fields("cycle").Value
		   workID = objRSStandings.Fields("ID").Value
		   workRank = objRSStandings.Fields("Rank").Value
		   
		   strSQL = "update standings_cycle set Rank = "&workRank&" where ID = "&workID&" and cycle = "&workCycle
		   'Response.Write "strSQL = "&strSQL&"<br>"
	       objConn.Execute strSQL
		   
		   objRSStandings.MoveNext
		wend
		objRSStandings.Close
		
	end if
	
	
	
	'Response.Write "Bottom of Standings Function<br>"

  End Function

  Function UpdateStandingsRD1 ()
	strSQL = "delete from Standings_RD1"
    objConn.Execute strSQL
    'Response.Write "Sql = " & strSQL  & "<br>"

	objRSStandings.Open "SELECT OwnerName, HomeTeamID, count(*) AS TotGames, avg(TeamScore) AS MyPPG, avg(OpponentScore) AS oppPPG, " &_
	                    "sum(switch(Result='W',1,Result='L',0)) AS adjusted_wins, " &_
						"sum(switch(Result='L',1,Result='W',0)) AS adjusted_losses " &_
                        "FROM tbl_points_scored, tblowners " &_
                        "WHERE TeamScore is not null " &_
						"and tbl_points_scored.gamedate >= (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') " &_
					    "and tbl_points_scored.gamedate < (select param_date from tblParameterCtl where param_name = 'PO_SEMIS') " &_
						"and tbl_points_scored.hometeamid = tblowners.ownerid  " &_
						"GROUP BY OwnerName, HomeTeamID " &_
                        "ORDER BY sum(switch(Result='W',1,Result='L',0)) DESC, " &_
						          "avg(TeamScore) DESC , avg(OpponentScore)", objConn,1,1

	wRank = 0
	While Not objRSStandings.EOF
	   'Response.Write "Checkpoint a <br>"
	   wRank = wRank + 1

	   if wRank = 1 then
	      wLeader = objRSStandings.Fields("adjusted_wins").Value
		  wGB = 0
	   else
	      wGB = wLeader - objRSStandings.Fields("adjusted_wins").Value
	   end if

	   wTeamNm  = objRSStandings.Fields("OwnerName").Value
	   wTeamID  = objRSStandings.Fields("HomeTeamID").Value
	   wWon     = objRSStandings.Fields("adjusted_wins").Value
	   wLoss    = objRSStandings.Fields("adjusted_losses").Value
	   wPPG     = objRSStandings.Fields("MyPPG").Value
	   wOPPG    = objRSStandings.Fields("oppPPG").Value
	   wDiff    = objRSStandings.Fields("MyPPG").Value - objRSStandings.Fields("oppPPG").Value

	   strSQL ="insert into Standings_RD1 (ID,Rank,Team,Won,Loss,PPG,OPPG,DIFF,GB) " & _
	           "values ("&wTeamID&","&wRank&",'"&wTeamNm&"',"&wWon&","&wLoss&","&wPPG&","&wOPPG&","&wDiff&","&wGB&")"

	   objConn.Execute strSQL

	   'Response.Write "Sql = " & strSQL  & "<br>"
	   objRSStandings.MoveNext

	Wend
	objRSStandings.Close

  End Function

  Function UpdateStandingsRD2 ()
	strSQL = "delete from Standings_RD2"
    objConn.Execute strSQL
    'Response.Write "Inside UpdateStandingsRD2 <br>"

	objRSStandings.Open "SELECT OwnerName, HomeTeamID, count(*) AS TotGames, avg(TeamScore) AS MyPPG, avg(OpponentScore) AS oppPPG, " &_
	                    "sum(switch(Result='W',1,Result='L',0)) AS adjusted_wins, " &_
						"sum(switch(Result='L',1,Result='W',0)) AS adjusted_losses " &_
                        "FROM tbl_points_scored, tblowners " &_
                        "WHERE TeamScore is not null " &_
						"and tbl_points_scored.gamedate >= (select param_date from tblParameterCtl where param_name = 'PO_SEMIS') " &_
					    "and tbl_points_scored.gamedate < (select param_date from tblParameterCtl where param_name = 'PO_FINALS') " &_
						"and tbl_points_scored.hometeamid = tblowners.ownerid  " &_
						"GROUP BY OwnerName, HomeTeamID " &_
                        "ORDER BY sum(switch(Result='W',1,Result='L',0)) DESC, " &_
						          "avg(TeamScore) DESC , avg(OpponentScore)", objConn,1,1

	wRank = 0
	While Not objRSStandings.EOF
	   'Response.Write "Checkpoint a <br>"
	   wRank = wRank + 1

	   if wRank = 1 then
	      wLeader = objRSStandings.Fields("adjusted_wins").Value
		  wGB = 0
	   else
	      wGB = wLeader - objRSStandings.Fields("adjusted_wins").Value
	   end if

	   wTeamNm  = objRSStandings.Fields("OwnerName").Value
	   wTeamID  = objRSStandings.Fields("HomeTeamID").Value
	   wWon     = objRSStandings.Fields("adjusted_wins").Value
	   wLoss    = objRSStandings.Fields("adjusted_losses").Value
	   wPPG     = objRSStandings.Fields("MyPPG").Value
	   wOPPG    = objRSStandings.Fields("oppPPG").Value
	   wDiff    = objRSStandings.Fields("MyPPG").Value - objRSStandings.Fields("oppPPG").Value

	   strSQL ="insert into Standings_RD2 (ID,Rank,Team,Won,Loss,PPG,OPPG,DIFF,GB) " & _
	           "values ("&wTeamID&","&wRank&",'"&wTeamNm&"',"&wWon&","&wLoss&","&wPPG&","&wOPPG&","&wDiff&","&wGB&")"

	   objConn.Execute strSQL

	   'Response.Write "Sql = " & strSQL  & "<br>"
	   objRSStandings.MoveNext

	Wend
	objRSStandings.Close

  End Function

  Function UpdateStandingsRD3 ()
	strSQL = "delete from Standings_RD3"
    objConn.Execute strSQL
	'Response.Write "Inside UpdateStandingsRD3 <br>"
    'Response.Write "Sql = " & strSQL  & "<br>"

	objRSStandings.Open "SELECT OwnerName, HomeTeamID, count(*) AS TotGames, avg(TeamScore) AS MyPPG, avg(OpponentScore) AS oppPPG, " &_
	                    "sum(switch(Result='W',1,Result='L',0)) AS adjusted_wins, " &_
						"sum(switch(Result='L',1,Result='W',0)) AS adjusted_losses " &_
                        "FROM tbl_points_scored, tblowners " &_
                        "WHERE TeamScore is not null " &_
                        "and tbl_points_scored.gamedate >= (select param_date from tblParameterCtl where param_name = 'PO_FINALS') " &_
						"and tbl_points_scored.hometeamid = tblowners.ownerid  " &_
						"GROUP BY OwnerName, HomeTeamID " &_
                        "ORDER BY sum(switch(Result='W',1,Result='L',0)) DESC, " &_
						          "avg(TeamScore) DESC , avg(OpponentScore)", objConn,1,1

	wRank = 0
	While Not objRSStandings.EOF
	   'Response.Write "Checkpoint a <br>"
	   wRank = wRank + 1
	   'wPen  = objRSStandings.Fields("total_Pen").Value

	   if wRank = 1 then
	      wLeader = objRSStandings.Fields("adjusted_wins").Value
		  wGB = 0
	   else
	      wGB = wLeader - objRSStandings.Fields("adjusted_wins").Value
	   end if

	   wTeamNm  = objRSStandings.Fields("OwnerName").Value
	   wTeamID  = objRSStandings.Fields("HomeTeamID").Value
	   wWon     = objRSStandings.Fields("adjusted_wins").Value
	   wLoss    = objRSStandings.Fields("adjusted_losses").Value
	   wPPG     = objRSStandings.Fields("MyPPG").Value
	   wOPPG    = objRSStandings.Fields("oppPPG").Value
	   wDiff    = objRSStandings.Fields("MyPPG").Value - objRSStandings.Fields("oppPPG").Value

	   strSQL ="insert into Standings_RD3 (ID,Rank,Team,Won,Loss,PPG,OPPG,DIFF,GB) " & _
	           "values ("&wTeamID&","&wRank&",'"&wTeamNm&"',"&wWon&","&wLoss&","&wPPG&","&wOPPG&","&wDiff&","&wGB&")"

	   objConn.Execute strSQL

	   'Response.Write "Sql = " & strSQL  & "<br>"
	   objRSStandings.MoveNext

	Wend
	objRSStandings.Close

  End Function
%>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=no">
<meta name="description" content="">
<meta name="Dee M. Myers" content="">
<title>Generate Results</title>
<!-- Bootstrap core CSS -->
<!--#include virtual="Common/bootstrap.inc"-->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
.bs-callout-success {
    border-left-color: #000000;
    padding: 10px;
    border-left-width: 4px;
    border-radius: 3px;
    background-color: white;
}
</style>
</head>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Generate Stats</strong>
			</div>
		</div>
	</div>
</div>
<div class="container">
  <div class="bs-callout bs-callout-success">
    <h4>Game Day Maintenance </h4>
    <ol>
      <li>Run After Daily Stats Posted</li>
			<li>Create Box Score</li>
      <li>Update Points Scored</li>
      <li>Build Standings</li>
      <li>Send Email on Name Mis-Matches</li>
			</ol>
  </div>
</div>
<br>
<body bgcolor="#FFFFF7">
<!--#include virtual="Common/headerMain.inc"-->
<form action="BuildBox.asp" method="POST">
  <input type="hidden" name="action" value="Save Form Data">
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<button class="btn btn-lg btn-default btn-block" value="Run BuildBox" name="Submit" type="submit">GENERATE NIGHTLY RESULTS</button>
			<input type="checkbox" name="email_league" value="Yes" >
			<span class="glyphicon glyphicon-envelope"></span>&nbsp;Email the League with this update!
		</div>
	</div>
</div>
</form>
<%
  Session.CodePage = Session("FP_OldCodePage")
  Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>