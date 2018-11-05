<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=no">
<meta name="description" content="">
<meta name="author" content="">
<title>IGBL 2018-2019</title>
<!--#include virtual="Common/bootstrap4.inc"-->
<!-- Custom styles for this template -->
<link href="css/stylesbs4.css" rel="stylesheet">
<style>
	body {
    font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif,"Apple Color Emoji","Segoe UI Emoji","Segoe UI Symbol";
    background-color: wheat;
    font-size: 12px;
}
.panel-primary {
    border-color: #ccc
}
.carousel-caption {
	position: relative;
	right: auto;
	bottom: 10px;
	left: auto;
	z-index: 10;
	padding-top: 20px;
	padding-bottom: 0px;
	color: black;
	text-align: center;
	text-shadow: 0 0px 0px rgba(0,0,0,.0);
}
.badgeBlue {
    display: inline-block;
    min-width: 10px;
    padding: 3px 7px;
    font-size: 25px;
    font-weight: 700;
    line-height: 1;
    color: black;
    text-align: center;
    white-space: nowrap;
    vertical-align: baseline;
    background-color: white;
    border-radius: 35px;
    border: double;
    border-color: darkorange;
}
.navbar-inverse {
    background-color: #ff8c00;
    border-color: #1e2d8c;
}
.btn-default {
    color: black;
    background-color: #f2f2f2;
    border-color: #ff8c00 !important;
    font-weight: bold !important;
}
.navbar-inverse .navbar-brand {
    color: #ffffff;
    font-weight: 600;
}
.btn-login {
	font-weight:600;
	color: white !important;
	background-color: #01579B !important;
	border-color: black !important;
}
#myCarousel {
  height: auto;
  width: auto;
  overflow: hidden;
}	
.h4, h4 {
    font-size: 20px;
}
.mark, mark {
    padding: .2em;
    background-color: #ffffff;
}
/* Make the image fully responsive 
.carousel-inner img {
	width: 100%;
	height: 100%;
}*/
</style>
</head>
<body> 

<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCID
	Session.CodePage = 1252
	Err.Clear

	Dim objRSlogin, objEmail, rstlogin, uid, pwd, hospid, cs, ownerid
	
	set objConn  = Server.CreateObject("ADODB.Connection")
	set rstlogin = Server.CreateObject("ADODB.Recordset")
	Set objEmail = Server.CreateObject("ADODB.RecordSet") 

	objConn.Open Application("lineupstest_ConnectionString")
	Dim strDatabaseType

	strDatabaseType = "Access"

	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
					"Data Source=lineuptest.mdb;" & _
					"Persist Security Info=False"
				  

  Session.Contents.RemoveAll()

  if (trim(request("txtUN")) <> "") then
     uid = request("txtUN")
     pwd = request("txtPW")

		rstlogin.Open "select userid, OwnerID, LastLogin from tblOwners where userid = '" & trim(uid) & "' and Password = '" & trim(pwd) & "'",objConn,3,3,1
		ownerid = rstlogin.Fields("OwnerID").Value

		if rstlogin.Recordcount > 0 then
			Session("LastLogin") = rstlogin.Fields("LastLogin").Value			
			strSQL = "update tblOwners set lastLogin = now() - 1/24 where ownerid = "& ownerid & ";"
			objConn.Execute strSQL
		else
			Response.Write "Your user id or password is invalid. <br>"
			Response.Write uid&"<br>"
			Response.Write pwd&"<br>"
			Response.End
		end if
		
		rstlogin.Close
		Session("ownerid") = ownerid
		Session.Timeout=240
		'Response.Redirect "dbdown.html"
		Response.Redirect "dashboard.asp"
  end if

%>
<%
 	Dim objConn,objRSwaivers,strSQL,iPlayerClaimed,objRSTxns,objRejectWaivers,iPlayerWaived,iOwner,w_action

	Set objRSwaivers       = Server.CreateObject("ADODB.RecordSet")
	Set objStagger         = Server.CreateObject("ADODB.RecordSet")
	Set objRSTxns 	       = Server.CreateObject("ADODB.RecordSet")
	Set objWork 	         = Server.CreateObject("ADODB.RecordSet")	
	Set objRejectWaivers   = Server.CreateObject("ADODB.RecordSet")
	Set objNextRun	       = Server.CreateObject("ADODB.RecordSet")
	Set objNoGame	         = Server.CreateObject("ADODB.RecordSet")	
	Set objrstrade         = Server.CreateObject("ADODB.RecordSet")
	Set objEmail           = Server.CreateObject("ADODB.RecordSet")	
	Set objParams          = Server.CreateObject("ADODB.RecordSet")
	Set objrsNames         = Server.CreateObject("ADODB.RecordSet")
	
	'#########################################################################################
	'### Check for Trade Deadline Date. When Passed Turn off Trade Indicator for all Owners
	'#########################################################################################
	objParams.Open "SELECT * FROM tblParameterCtl where param_name = 'TRADE_DEADLINE' ",objConn
	wTradeDeadLine = objParams.Fields("param_date").Value
	objParams.Close
	
	if (wTradeDeadLine + 1 + 1/24) < now() then 
		strSQL = "update tblOwners set acceptTradeOffers = false"  
		objConn.Execute strSQL
	end if 	
	
	'********************************************************
	'Remove expired Trade Offers
	'********************************************************
	objrstrade.Open "SELECT * FROM tblPendingTrades WHERE DecisionDate < (now() - 1/24)", objConn,3,3,1
	if objrstrade.Recordcount > 0 then
		While Not objrstrade.EOF		    
			tradeid     = objrstrade.Fields("TradeID").Value
			FromOwnerID = objrstrade.Fields("FromOid").Value
			ToOwnerID   = objrstrade.Fields("ToOid").Value
			FuncCall    = Get_Player_Team_Names(tradeid, msgtradeplayers, msgacquireplayers, FromOwnerName, ToOwnerName)			
			'DecisionDate = objrstrade.Fields("DecisionDate").Value
			'Response.Write "Remove trade.  TradeID = "&tradeid&", comparedate() = "&(now()-1/24)&", DecisionDate = "&DecisionDate&"<br>"				
			
			wEmailOwnerID  = FromOwnerID
			wAlert         = "receiveTradeAlerts"
			email_subject  = "Trade Offer to " & ToOwnerName & " Expired "
			email_message  = msgtradeplayers & "<br>"
			email_message  = email_message & "for <br><br>"
			email_message  = email_message & msgacquireplayers
%>		
		<!--#include virtual="Common/email_league.inc"-->
				   
<%		
	
			wEmailOwnerID  = ToOwnerID
			email_subject  = "Trade Offer from " & FromOwnerName & " Expired "
%>		
		<!--#include virtual="Common/email_league.inc"-->
				   
<%					
			
			strSQL = "DELETE FROM tblpendingtrades WHERE TradeID =   "& tradeid & ";"
			objConn.Execute strSQL
		
			objrstrade.MoveNext
		Wend
		
		FuncCall = Reset_PendTrade_Flags
	end if
	
	objrstrade.Close	
	'********************************************************
	'Run pendingwaiversall event if it has not been run today.
	'********************************************************
	objRSwaivers.Open "SELECT * FROM tbltimedEvents " & _
	                  "where event = 'pendingwaiversall' and nextrun_EST < now() ", objConn,3,3,1

	if  objRSwaivers.Recordcount > 0 then
		objRSTxns.Open		"SELECT * FROM qryUpdatewaiver ", objConn,3,3,1
    	w_action = objRSTxns.Recordcount
		
		objWork.Open "SELECT param_amount FROM tblParameterCtl where param_name = 'PICKUP' ",objConn
	    wPickupAmt = objWork.Fields("param_amount").Value
	    objWork.Close
		
		'Response.Write "Count = : " & w_action & "<br>"
    	while w_action > 0

			iPlayerClaimed= objRSTxns.Fields("PID_Claimed").Value
			iPlayerWaived = objRSTxns.Fields("PID_Waived").Value
			iOwner        = objRSTxns.Fields("OwnerID").Value
			iPriority     = objRSTxns.Fields("WaiverPriority").Value
			iActivePlayers= objRSTxns.Fields("ActivePlayerCnt").Value
			iwaiverBid    = objRSTxns.Fields("waiverBid").Value
			iwaiverID     = objRSTxns.Fields("waiverID").Value

			objWork.Open "SELECT WaiverBal FROM tblOwners where ownerid = "&iOwner&";",objConn
	        wWaiverBal = objWork.Fields("WaiverBal").Value
	        objWork.Close
		
			if iPlayerWaived = 0 AND iActivePlayers >= 14 then
				'***************************************************************
				'REJECT THIS TRANSACTION BECAUSE THE PLAYER LIMIT IS 14 PER TEAM
				'****************************************************************
				'TransType = "Waiver Rej. (Roster Full)"
				TransType = "Waiver Reject (Roster Full)"
				Cost = 0.00
				
				strSQL ="insert into tblTransactions(OwnerID,TransType,transAddPlayer1,TransCost,waiverbid,transAddPlayerCnt) " &_ 
				        "values ("&iOwner& ",'"&TransType&"',"&iPlayerClaimed&",'"&Cost&"',"&iwaiverBid&",1)"
				objConn.Execute strSQL

		 		'strSQL = "DELETE from tblWaivers where PID_Waived = 0 and PID_Claimed = "&iPlayerClaimed&" and OwnerID = "&iOwner&";"
				strSQL = "DELETE from tblWaivers where waiverID = "&iwaiverID&";"
				
				objConn.Execute strSQL
			elseif wWaiverBal < iwaiverBid then
				'TransType = "Waiver Rej. Bid > Bal."
				TransType = "Waiver Reject (Bid > Bal)"
				Cost = 0.00
				
				strSQL ="insert into tblTransactions(OwnerID,TransType,transAddPlayer1,TransCost,waiverbid,transAddPlayerCnt) " &_ 
				        "values ("&iOwner& ",'"&TransType&"',"&iPlayerClaimed&",'"&Cost&"',"&iwaiverBid&",1)"
				objConn.Execute strSQL

				strSQL = "DELETE from tblWaivers where waiverID = "&iwaiverID&";"
				objConn.Execute strSQL
			else
				'**********************************************************
				'UPDATE TO PLAYER TABLE for player being added.
				'**********************************************************
				strSQL = "update tblPlayers SET playerStatus = 'O', OwnerId = " & iOwner & ", " &_
				         "pendingwaiver = 0, clearwaiverdate = null, LastTeamInd = null " &_
				         "WHERE tblPlayers.PID = " & iPlayerClaimed & ";"
				objConn.Execute strSQL

				'******************************************************************
				'UPDATE TO OWNERS TABLE.  Update other owners waiver priorities first
				'then set the current owner's waiver priority to 10.
				'******************************************************************
				strSQL ="update tblowners SET waiverpriority = waiverpriority - 1 WHERE waiverpriority > " & iPriority & ";"
				objConn.Execute strSQL
				
				'TransType = "Signed off waivers"
				TransType = "Added"
				Cost = wPickupAmt
				if iPlayerWaived = 0 then
					strSQL = "update tblowners SET waiverpriority = 10, ActivePlayerCnt = ActivePlayerCnt + 1, " &_  
					         "Waiverbal = Waiverbal - "&iwaiverBid&" WHERE ownerid = "&iOwner&";"
					objConn.Execute strSQL
					
					'strSQL ="insert into tblTransactions(OwnerID,TransType,transAddPlayer1,TransCost,transAddPlayerCnt) values ('" &_
					'iOwner & "', '" & TransType & "','" &  iPlayerClaimed  & "', '" &  Cost & "', 1)"
					'objConn.Execute strSQL
					
					strSQL = "insert into tblTransactions(OwnerID,TransType,transAddPlayer1,TransCost,waiverbid,transAddPlayerCnt) " &_ 
				             "values ("&iOwner& ",'"&TransType&"',"&iPlayerClaimed&",'"&Cost&"',"&iwaiverBid&",1)"
				    objConn.Execute strSQL
					
				else
					strSQL ="update tblowners SET waiverpriority = 10, Waiverbal = Waiverbal - "&iwaiverBid&" WHERE ownerid = "&iOwner& ";"
					objConn.Execute strSQL

					'**********************************************************
					'Update Player Table for Player being Waived
					'**********************************************************
					strSQL = "update tblPlayers " & _
					         "SET playerStatus = 'W', OwnerId = 0, pendingwaiver = 0, Ontheblock = 0, ir = 0, " & _
					         "LastTeamInd = "&iOwner&", clearwaiverdate = date() + 1 " & _
					         "WHERE tblPlayers.PID = "&iPlayerWaived&";"

					objConn.Execute strSQL
					'**********************************************************
					'Player Released TRANSACTION
					'**********************************************************
					'TransType = "Released"
					'Cost = 0.00
					'strSQL ="insert into tblTransactions(OwnerID,TransType,PID,TransCost) values ('" &_
					'iOwner & "', '" &  TransType & "', '" & iPlayerWaived & "', '" &  Cost & "')"
					'objConn.Execute strSQL
					
					'**********************************************************
					'Cleanup Pending Trades
					'**********************************************************
					strSQL = "DELETE FROM tblpendingtrades " & _
                             "WHERE  tradedplayerid      = "& iPlayerWaived & "  "  & _
                             "or tradedplayerid2 = "& iPlayerWaived & "   " & _
                             "or tradedplayerid3 = "& iPlayerWaived & "   " & _
                             "or acquiredplayerid = "& iPlayerWaived & "  " & _
                             "or acquiredplayerid2 = "& iPlayerWaived & " " & _
                             "or acquiredplayerid3 = "& iPlayerWaived & "  "
		            objConn.Execute strSQL
		
					FuncCall = Remove_From_Lineups(iPlayerWaived,0,0,0,0,0,iOwner,0)
					
					strSQL ="insert into tblTransactions " & _
				"(OwnerID,TransType,TransCost,transAddPlayerCnt,transAddPlayer1,transReleasePlayerCnt,transReleasePlayer1,waiverbid) " & _
				"values ("&iOwner&",'"&TransType&"',"&Cost&",1,"&iPlayerClaimed&",1,"&iPlayerWaived&","&iwaiverBid&") "		
					objConn.Execute strSQL
					
				end if

				'**********************************************************
				'Player Signed TRANSACTION
				'**********************************************************
				'TransType = "Signed off waivers"
				'Cost = wPickupAmt
				'strSQL ="insert into tblTransactions(OwnerID,TransType,PID,TransCost) values ('" &_
				'iOwner & "', '" &  TransType & "', '" &  iPlayerClaimed  & "', '" &  Cost & "')"				
				'objConn.Execute strSQL

				objRejectWaivers.Open "SELECT * FROM tblWaivers " &_ 
				                      "where PID_Claimed = "&iPlayerClaimed&" and OwnerID <> "&iOwner&" order by waiverbid desc" , objConn
				'TransType = "Rejected"
				TransType = "Rejected"
				Cost = 0.00

				While Not objRejectWaivers.EOF

					iRejOwner = objRejectWaivers.Fields("OwnerID").Value
					iWaivedReject = objRejectWaivers.Fields("pid_waived").Value
					iClaimedReject = objRejectWaivers.Fields("pid_claimed").Value
					iBidReject = objRejectWaivers.Fields("waiverbid").Value
					
					'strSQL ="insert into tblTransactions(OwnerID,TransType,transAddPlayer1,TransCost,transAddPlayerCnt) values ('" &_
					'iRejOwner & "', '" & TransType & "','" &  iClaimedReject  & "', '" &  Cost & "', 1)"
					'objConn.Execute strSQL

					strSQL ="insert into tblTransactions(OwnerID,TransType,transAddPlayer1,TransCost,waiverbid,transAddPlayerCnt) " &_ 
				            "values ("&iRejOwner& ",'"&TransType&"',"&iClaimedReject&",'"&Cost&"',"&iBidReject&",1)"
				    objConn.Execute strSQL
					
					'Update PendingWaiver flag
					strSQL = "update tblPlayers SET pendingwaiver = 0 WHERE tblPlayers.PID = " & iWaivedReject & ";"
					objConn.Execute strSQL

					objRejectWaivers.MoveNext
				Wend

				objRejectWaivers.Close

				'*************************************************************************
				'Delete all entries from tblWaivers table where player_id = Player Claimed
				'*************************************************************************
				strSQL = "DELETE from tblWaivers where PID_Claimed = " & iPlayerClaimed & ";"
				objConn.Execute strSQL

				'*************************************************************************
				'Delete any additional rows from the tblWaivers table for the player that
				'was just waived.  This is necessary if the owner had the same player on
				'multiple waivers.
				'*************************************************************************
				if iPlayerWaived <> 0 then
					strSQL = "DELETE from tblWaivers where PID_Waived = " & iPlayerWaived & ";"
					objConn.Execute strSQL
				end if
		
			end if
		
			'*********************************************************
			'Close the Query and Open it again to see if any rows remain
			'**********************************************************
			ObjRsTxns.Close
			objRSTxns.Open	"SELECT * FROM qryUpdatewaiver ", objConn,3,3,1
			w_action = objRSTxns.Recordcount

		wend

		ObjRsTxns.Close
		
		'********************************************************************
		'Send Email Notification to the League that Waivers have run for today!
		'*********************************************************************
	
			Dim email_to, email_subject, host, username, password, reply_to, port, from_address
			Dim first_name, last_name, home_address, email_from, telephone, comments, error_message
			Dim ObjSendMail, email_message

			wEmailOwnerID  = null
			wAlert         = "receiveWaiverAlerts"
			'email_subject= "Waivers Processed: "&date()
			email_subject= "Waivers: "&date()
			email_message = null			
			objRSTxns.Open	"SELECT * FROM qWaiversReportemail ", objConn,3,3,1							
			'*********************************************************************
			'New Code to include Transactions in the email 09/16/2015
			'*********************************************************************	
            'Response.Write "Top of Loop <br>"			
			While Not ObjRsTxns.EOF						
			email_message = email_message & left(ObjRsTxns.Fields("Claim_Fname").Value,1) & ". " & ObjRsTxns.Fields("Claim_Lname").Value & " " & ObjRsTxns.Fields("TransType").Value & " by " & ObjRsTxns.Fields("ShortName").Value & "<br>" 			
			if ObjRsTxns.Fields("Waive_Fname").Value <> "" then
			   email_message = email_message & left(ObjRsTxns.Fields("Waive_Fname").Value,1) & ". " & ObjRsTxns.Fields("Waive_Lname").Value & " Dropped by " & ObjRsTxns.Fields("ShortName").Value & "<br>" 
			end if
			ObjRsTxns.MoveNext
			Wend
			ObjRsTxns.Close
			
			if ISNULL(email_message) then
			   email_message = "No waivers entered today. Players are now available for P/U or Rental!"
			end if
			
			'email_message = replace(replace(replace(email_message, "Added", "(A)"), "Dropped", "(D)"), "Rejected", "(R)")
			'email_message = replace(email_message, "<br>", "<vbcrlf>")
			'Response.Write email_message
			
%>		
		<!--#include virtual="Common/email_league.inc"-->
				   
<%				
			
		'*********************************************************************
		'Make players Free whose clearwaiver Date is less then Now()
		'and Set Rental Players back Free
		'*********************************************************************
		strSQL = "update tblPlayers " & _
		         "SET playerStatus = 'F', OwnerId = 0, clearwaiverdate = null, LastTeamInd = null " & _
		         "WHERE clearwaiverdate < now() and playerStatus = 'W'"

		objConn.Execute strSQL

        '**************************************
        'CLEANUP when waiver stacking is used
        '**************************************
		strSQL = "update tblPlayers SET pendingwaiver = 0 "
 		objConn.Execute strSQL

		'***************************************************************************
		'Set the time for the next pendingwaiversall run.  If tomorrow is a game
		'day, then set the pendingwaivers date to run 6 hours before cutofftime.
		'Note that that code subtracts 5 hours from the time because the times
		'in the database are CST but the server is hosted on EST.  If tomorrow is
		'not a IGBL game day then set the nextrun 6 hours before the first NBA game.
		'***************************************************************************
		objNextRun.Open	"SELECT * FROM qryGamedeadlines where gameday = (date() + 1) ", objConn,3,3,1

		if objNextRun.Recordcount > 0 then
			dnextrun = objNextRun.Fields("cutofftime").Value - 5/24
		else
		    objNoGame.Open "SELECT min(TipTimeEst) as EarlyTip FROM tblLeagueSetup where gameDate = date() + 1 ", objConn
			if ISNULL(objNoGame.Fields("EarlyTip").Value) then
			   dnextrun = date() + 1 + (13/24)
	        else
		       dnextrun = (date() + 1 + objNoGame.Fields("EarlyTip").Value) - 6/24
	        end if
			objNoGame.Close
		end if

		strSQL = "update tbltimedEvents " & _
				 "SET lastrun_EST = now(), nextrun_EST = '"&dnextrun&"' " & _
				 "WHERE event = 'pendingwaiversall' "
		objConn.Execute strSQL
		objNextRun.Close
		
		FuncCall = Reset_PendTrade_Flags
	end if

	objRSwaivers.Close	

	Set objRSLeaders   = Server.CreateObject("ADODB.RecordSet")
	objRSLeaders.Open  "SELECT MAX(gameDate) as nbaGameDate FROM tblLast5 ", objConn,3,3,1

	if IsNull(objRSLeaders.Fields("nbaGameDate")) then
		wGameDay = date()
	else
		wGameDay = objRSLeaders.Fields("nbaGameDate").Value
	end if

	objRSLeaders.Close
	objRSLeaders.Open  	"SELECT tblLast5.*, tblPlayers.pid, tblPlayers.ownerid, tblPlayers.pos, tblPlayers.image, TBLOWNERS.shortname, " & _
											"tblNBATeams.teamShortName " & _
											"FROM ((tblLast5 LEFT JOIN tblPlayers ON (tblLast5.Last = tblPlayers.lastName) AND (tblLast5.First = tblPlayers.firstName)) " & _
											"LEFT JOIN TBLOWNERS ON tblPlayers.ownerID = TBLOWNERS.ownerID) " & _
											"LEFT JOIN tblNBATeams ON tblPlayers.NBATeamID = tblNBATeams.NBATID " & _
											"WHERE tblLast5.gameDate= #"&wGameDay&"# and tblLast5.BarpTot >= 45 " &_
											"ORDER BY tblLast5.barptot DESC, tblLast5.last", objConn,3,3,1
									   
				

    '***********************************************************
	' Get_Player_Team_Names()
	'***********************************************************
   	Function Get_Player_Team_Names(p_tradeID, p_msgtradeplayers, p_msgacquireplayers, p_FromOwnerName, p_ToOwnerName)

		objrsNames.Open "SELECT * FROM qryPendingTrade Where TradeID = "&p_tradeID&" " , objConn

		'********************************************************
		'** Populate Traded Players String
		'********************************************************
		p_msgtradeplayers = objrsNames.Fields("t1first").Value & " " & objrsNames.Fields("t1last").Value & "<br />"

		if objrsNames.Fields("t2PID").Value > 0 then
			p_msgtradeplayers = p_msgtradeplayers & objrsNames.Fields("t2first").Value & " " & objrsNames.Fields("t2last").Value & "<br />"  
		end if

		if objrsNames.Fields("t3PID").Value > 0 then
			p_msgtradeplayers = p_msgtradeplayers & objrsNames.Fields("t3first").Value & " " & objrsNames.Fields("t3last").Value & "<br />"
		end if

		'********************************************************
		'** Populate Acquired Players String
		'********************************************************
		p_msgacquireplayers= objrsNames.Fields("a1first").Value & " " & objrsNames.Fields("a1last").Value & "<br />"

		if objrsNames.Fields("a2PID").Value > 0 then
			p_msgacquireplayers = p_msgacquireplayers & objrsNames.Fields("a2first").Value & " " & objrsNames.Fields("a2last").Value & "<br />"
		end if

		if objrsNames.Fields("a3PID").Value > 0 then
			p_msgacquireplayers = p_msgacquireplayers &  objrsNames.Fields("a3first").Value & " " & objrsNames.Fields("a3last").Value & "<br />"
		end if
		
		'********************************************************
		'** Populate Owner Short Names
		'********************************************************
        p_FromOwnerName = objrsNames.Fields("tteamnameshort").Value
		p_ToOwnerName   = objrsNames.Fields("ateamnameshort").Value
		
		objrsNames.Close
	End Function				
				
%>	
<!--#include virtual="Common/setStaggeredAll.inc"-->
<!--#include virtual="Common/setwaiversall.inc"-->
<!--#include virtual="Common/header.inc"-->  
<!-- Fixed navbar -->
</br></br>
<div class="container">
	<div class="row">
		<div class="col-sm-12">
			<div id="myCarousel" class="carousel slide" data-ride="carousel">				
				<div class="carousel-inner">
					<div class="carousel-item active">
						<img class="center-block" src="images/logomed.gif" style="margin:0px auto;display:block;">
					</div>
					<%
					While Not objRSLeaders.eof
					teamName = objRSLeaders.Fields("shortname").Value
					
					if isNull(teamName) then 
						teamName = "Free Agent"
					end if
				
					%>
					<div class="carousel-item">
						<img class="rounded-circle" src="<%= objRSLeaders.Fields ("image").Value %>" alt="Los Angeles" style="margin:0px auto;display:block;background-color:white">
						<div class="carousel-caption">
							<h4><span style="text-align:center;color:#01579B"><mark><%=objRSLeaders.Fields("first").Value%>&nbsp;<%=objRSLeaders.Fields("last").Value%></mark></h4><span"><h4 style="darkorange;text-shadow: #ff8c00 2px 0px 1px;-webkit-font-smoothing: antialiased;">
							<%=teamName%></h4></span></span>
							<span class="badgeBlue"><%= objRSLeaders.Fields ("barptot").Value %></span>
						</div>
					</div>	
					<%
					objRSLeaders.MoveNext
					Wend
					%>
				</div>
				<!-- Left and right controls -->
				<a class="carousel-control-prev" href="#myCarousel" data-slide="prev">
					<span style="background-color:darkorange;" class="carousel-control-prev-icon"></span>
				</a>
				<a class="carousel-control-next" href="#myCarousel" data-slide="next">
					<span style="background-color:darkorange;" class="carousel-control-next-icon"></span>
				</a>
			</div>				
		</div>
	</div>
</div>
</br></br>
<div class="container">	
	<div class="row">
		<div class="col-sm-12">	
		  <form class="form-inline" role="form">
			  <div class="input-group mb-2 col-sm-12">
					<div class="input-group-prepend">
						<span style="background-color:darkorange;color:white;" class="input-group-text"><i class="fa fa-user" aria-hidden="true"></i></span>
					</div>
					<input type="text" class="form-control" name="txtUN" placeholder="Enter UserID">
				</div>
				<span class="help-block"></span>
				<div class="input-group mb-2 col-sm-12">
					<div class="input-group-prepend">
						<span style="background-color:darkorange;color:white;" class="input-group-text"><i class="fa fa-lock" aria-hidden="true"></i></span>
					</div>
					<input type="password" class="form-control" name="txtPW" placeholder="Enter Password">
				</div>
				<div class="col-sm-12">
				<button type="submit" class="btn btn-default btn-block mb-2"" value="Log In"><i class="fas fa-sign-in-alt"></i> Sign In</button>
				</div>
			</form>
		</div>
	</div>
</div>
<%
  objConn.Close
	objRSLeaders.Close
  Set objConn = Nothing
%>
<!--#include virtual="Common/functions.inc"-->
<div class="container-fluid">
  <div class="row">
    <div class="col-sm-12">
			<!--#include virtual="Common/footer.inc"-->  
		</div>
	</div>
</div>	
</body>
</html>