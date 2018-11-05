<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCID
	Session.CodePage = 1252
	Err.Clear
	strErrorUrl = ""
	Session.Timeout=120
    
	Dim objConn, sAction, sURL,ownerid, pid
	
	
	Set objConn	= Server.CreateObject("ADODB.Connection")
	objConn.Open Application("lineupstest_ConnectionString")

	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"
	          
	Dim objRS,pStatus,teamName,objRSNext5,skedcnt,objRSPO,objRSLast5,objEmail
	Set objRS         = Server.CreateObject("ADODB.RecordSet")
	Set objRSNext5    = Server.CreateObject("ADODB.RecordSet")
	Set objRSReg      = Server.CreateObject("ADODB.RecordSet")
	Set objRSPO       = Server.CreateObject("ADODB.RecordSet")
	Set objRSLast5    = Server.CreateObject("ADODB.RecordSet")
	Set objRSAvgLast5 = Server.CreateObject("ADODB.RecordSet")
	Set objRSPlayers  = Server.CreateObject("ADODB.RecordSet")
	Set objRSNBASked  = Server.CreateObject("ADODB.RecordSet")
	Set objRSActive   = Server.CreateObject("ADODB.RecordSet")
	Set objEmail      = Server.CreateObject("ADODB.RecordSet")
	
	GetAnyParameter "Action", sAction	
	'Response.Write "ACTION   " &sAction& "<br> "
	if len(sAction) > 0 then 
		PID_Split = Split(Request.Form("Action"), ";")
		Response.Write "SPLIT ACTION   " &PID_Split(0)& "<br> "
		Response.Write "PID            " &PID_Split(1)& "<br> "
		Response.Write "PLAYER NAME    " &PID_Split(2)& "<br> "
		sAction = PID_Split(0)
	end if%>
	<!--#include virtual="Common/session.inc"-->
	<%		
		

	pid = Request.querystring("pid")
	if (CInt(pid)) > 0 then
	else
		pid = Request.Form("pid")
	end if
	
	if pid = "" then
		Response.Redirect("timeout.asp")
	end if
	
	wPlayerID = pid
	objRSActive.Open "SELECT * FROM tbl_lineups where gameday = date() AND ownerID = "&ownerid& " " & _
															"AND final_lineup = FALSE AND (" & _
															"(sCenter   = "&wPlayerID&"  AND sCenterTip   < time() - 1/24) OR " & _
															"(sForward  = "&wPlayerID&"  AND sForwardTip  < time() - 1/24) OR " & _
															"(sForward2 = "&wPlayerID&"  AND sForwardTip2 < time() - 1/24) OR " & _
															"(sGuard    = "&wPlayerID&"  AND sGuardTip    < time() - 1/24) OR " & _
															"(sGuard2   = "&wPlayerID&"  AND sGuardTip2   < time() - 1/24) " & _
															") ",objConn,3,3,1
					
																														
	if objRSActive.RecordCount > 0 then
		 wPlayerInPlay = True
	else
		 wPlayerInPlay = False
	end if
	
	objRSActive.Close												

	objRS.Open       	"SELECT * FROM qPlayerProfiles WHERE (((qPlayerProfiles.pid)=" & pid & "))",objConn,1,1
	objRSNext5.Open  	"SELECT * FROM qryAllPlayerGameDays WHERE pid = " & pid & " and gameday >= Date() order by gameday ", objConn,1,1
	objRSReg.Open  		"SELECT * FROM qryAllPlayerGameDays WHERE pid = " & pid & " and gameday >= Date() " & _
										"and gameday < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') order by gameday ", objConn,1,1					 
	objRSPO.Open     	"SELECT * FROM qryAllPlayerGameDays WHERE pid = " & pid & " " & _
										"and gameday >= (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,1,1
	
	wFirstName = objRS.Fields("firstName").Value
	wLastName  = objRS.Fields("lastName").Value

	objRSLast5.Open   "SELECT t.*, format(gameDate, 'mm/dd') as xDate " & _
										"FROM tblLast5 t " & _
	                  "WHERE t.first = '" & wFirstName & "'  and t.last ='"&wLastName & "' order by gameDate desc", objConn,1,1
	
	'Response.Write "LAST 5 GAME COUNT   " &objRSLast5.RecordCount& "<br> "
	
	if objRSLast5.RecordCount >= 5 then
		last5LoopCtrl = 5
	else 
		last5LoopCtrl = objRSLast5.RecordCount
	end if	

	espnLink       = objRS.Fields("espnLink").Value
	pStatus        = objRS.Fields("playerStatus").Value
	teamName       = objRS.Fields("teamName").Value
	pos            = objRS.Fields("pos").Value
	rentedPlayer	 = objRS.Fields("rentalPlayer").Value  
	pendingWaiver  = objRS.Fields("PendingWaiver").Value  
	pendingTrade	 = objRS.Fields("PendingTrade").Value
	
	

	playerGmCnt  = objRSReg.RecordCount
	playerPOGmCnt= objRSPO.RecordCount
	
	objRSAvgLast5.Open  "SELECT  avg(x3p) as avg3pt, avg(TRB) as avgReb, avg(AST) as avgAst, avg(STL) as avgStl, avg(BLK) as avgBlks, " &_
								      "avg(TOV) as avgTo, avg(PTS) as avgPts, avg(BARPTot) as avgBarps, avg(MP) as avgMP " &_
						          "FROM tblLast5 t " & _
						          "WHERE t.first = '"&wFirstName&"' and t.last = '"&wLastName&"' and L5Game = 1", objConn,1,1
	
	avgMP    = objRSAvgLast5.Fields("avgMP").Value
	avgBlks  = objRSAvgLast5.Fields("avgBlks").Value
	avgAst   = objRSAvgLast5.Fields("avgAst").Value
	avgReb   = objRSAvgLast5.Fields("avgReb").Value
	avgPts   = objRSAvgLast5.Fields("avgPts").Value
	avgStl   = objRSAvgLast5.Fields("avgStl").Value
	avg3pt   = objRSAvgLast5.Fields("avg3pt").Value
	avgTo    = objRSAvgLast5.Fields("avgTo").Value
	avgBarps = objRSAvgLast5.Fields("avgBarps").Value
	'Response.Write "AVERAGE BARP TOTAL  " &avgBarps& "<br> "
	'	Response.Write "RECORD COUNT  " &objRSAvgLast5.RecordCount& "<br> "
	objRSAvgLast5.Close
	
	skedcnt  = 0
	last5cnt = 1
		
	 objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRS.Fields("NBATeamID").Value, objConn,3,3,1
	 'Must assign it to a variable because the code give inconsitent values when you evaluate the Time field when no rows are returned.
	 if objRSNBASked.RecordCount > 0 then
			 wTipTime = objRSNBASked.Fields("GameTime").Value
	 else
			 wTipTime = "12:00:00 AM"
	 end if						   
	 if len(objRSNBASked.Fields("GameTime").Value) = 10 then
			wtime = Left(objRSNBASked.Fields("GameTime").Value,4) & Right(objRSNBASked.Fields("GameTime").Value,3)
	 else
			wtime = Left(objRSNBASked.Fields("GameTime").Value,5) & Right(objRSNBASked.Fields("GameTime").Value,3)
	 end if	

	 
	 'Response.Write "DO I PLAY TODAY  " &objRSNBASked.RecordCount& "<br> "
	
	
	select case pStatus
	case "F"
	 	pStatus = "Free Agent"
	case "S"
	 	pStatus = "Staggered"
	case "W"
		pStatus = "Waivers"
	case "O"
		pStatus = objRS.Fields("shortName").Value
	case "R"
		pStatus = "Rented"
	end select

	select case sAction

		
	case "Counter"
		sURL = "tradeanalyzer.asp"
		conAction = "Continue"
		
		var_tradepartner = Request.Form("var_tradepartner")
		var_ownerid      = Request.Form("var_ownerid")
		var_tradeid      = 0
		dim cmbTeam
		var_sPid         = Request.Form("pid")
		
		AddLinkParameter "var_ownerid", var_ownerid, sURL	
		AddLinkParameter "Action", conAction, sURL
		AddLinkParameter "cmbTeam", var_tradepartner, sURL
		AddLinkParameter "var_tradeid", var_tradeid, sURL
		AddLinkParameter "var_sPid", var_sPid, sURL
		Response.Redirect sURL
		
		'******************************
		'*** PROCESS WAIVERS ACTION ***
		'******************************
		case "Process Waivers"

		Response.Write "OWNERID    " &ownerID& "<br> "
		
		iRecordToUpdate = PID_Split(1)
		sPlayerName     = PID_Split(2)
		Response.Write "Player waived is : " & sPlayerName & "<br>"
	  
		objRSPlayers.Open  "SELECT * FROM qry_PlayerAll WHERE (((qry_PlayerAll.pid)=" & iRecordToUpdate & "))", objConn,3,3,1

		if objRSPlayers.Fields("ownerID").value = 0 then
			errorCode = True
			errorDesc = "Player Released Previously"

		else
			errorCode = False
			'UPDATE TO PLAYER TABLE
			Response.Write "UPDATE PLAYER TABLE = " & PID_Split(1)  & ".<br>"  
			strSQL = "update tblPlayers SET PlayerStatus = 'W', ontheblock = 0, OwnerID = 0, "&_
							 "LastTeamInd = "& ownerid &", clearwaiverDate = date() + 1 "&_
							 "WHERE tblPlayers.PID = " & iRecordToUpdate & ";"
			objConn.Execute strSQL
			Response.Write "Sql = " & strSQL  & ".<br>"  
			
			FuncCall = Remove_From_Lineups(iRecordToUpdate,0,0,0,0,0,ownerid,0)
			
			'UPDATE  TO OWNER TABLE
			strSQL = "update tblOwners SET ActivePlayerCnt = ActivePlayerCnt - 1 WHERE tblOwners.OwnerID = " & ownerid & ";"
			objConn.Execute strSQL
			Response.Write "Sql = " & strSQL  & ".<br>"  
			'INSERT TO TRANSACTION TABLE
			TransType = "Released"
			Cost = 0.00

			strSQL ="insert into tblTransactions(OwnerID,TransType,TransReleasePlayer1,TransCost,transReleasePlayerCnt) values ('" &_
			ownerid & "', '" &  TransType & "', '" & iRecordToUpdate & "', '" &  Cost & "', 1)"
			objConn.Execute strSQL

			Response.Write "Sql = " & strSQL  & ".<br>"  
			
			objRSPlayers.Close

			Set objRS1      = Server.CreateObject("ADODB.RecordSet")	

			objRS1.Open "SELECT ShortName FROM tblowners WHERE tblowners.OwnerID = " & ownerid & " ", objConn
			Response.Write "OWNERID    " &ownerid& "<br> "
			txnteamname = objRS1.Fields("ShortName").value
			Response.Write "TEXT NAME    " &txnteamname& "<br> "
			objRS1.Close
			
				'**************************************************************************
				'Send Email Notification to the League that the player has been waived.
				'**************************************************************************
				wEmailOwnerID  = null
				wAlert         = "receiveWaiverAlerts"
				email_subject  = "Player Waived by "& txnteamname
				email_message  = sPlayerName
%>		
		<!--#include virtual="Common/email_league.inc"-->
				   
<%			
					end if
					
					sURL = "dashboard.asp"
					AddLinkParameter "var_ownerid", ownerid, sURL
					Response.Redirect sURL

		'******************************
		'*** PROCESS INJURY ACTION  ***
		'******************************			
		
		case "Injury"
		'hurtPlayerCnt   = Request.Form("chkIRPID").count
		iRecordToUpdate = PID_Split(1)
		
		objRSPlayers.Open  "SELECT * FROM TBLPLAYERS WHERE PID = "& iRecordToUpdate &" ", objConn,3,3,1
		strSQL = "SELECT * TBLPLAYERS WHERE PID = "&iRecordToUpdate&""
	  
		if objRSPlayers.Fields("IR").value = true then	
		
			strSQL = "UPDATE TBLPLAYERS set IR = FALSE WHERE PID = "&iRecordToUpdate&""
	    objConn.Execute strSQL
			'Response.Write "SQL:   " & strSQL & "<br>"
		else
			strSQL = "UPDATE TBLPLAYERS set IR = TRUE WHERE PID = "&iRecordToUpdate&""
	    objConn.Execute strSQL		
			'Response.Write "SQL:   " & strSQL & "<br>"
		end if
		
		sURL = "dashboard.asp"
		AddLinkParameter "var_ownerid", ownerid, sURL
		Response.Redirect sURL


	case "Transaction"
		
		sURL = "transelect.asp"
	
		ownerid           = Request.Form("var_ownerid")
		sPlayer           = pid
		conAction         = "PPTransaction"
	
		AddLinkParameter "var_ownerid", ownerid, sURL	
		AddLinkParameter "Action", conAction, sURL
		AddLinkParameter "var_sPid", sPlayer, sURL
		Response.Redirect sURL
	
	case "Waive Player"

		sURL = "maintainRoster.asp"
		sPlayer           = pid
		conAction = "Waive Player"
		ownerid   = Request.Form("var_ownerid")
		
		Response.Write "var_ownerid = " & ownerid & ".<br>"
	
		AddLinkParameter "var_ownerid", ownerid, sURL	
		AddLinkParameter "Action", conAction, sURL
		AddLinkParameter "var_sPid", sPlayer, sURL
		Response.Redirect sURL
		
	end select

	
%>
<!--#include virtual="Common/noTrades.inc"-->
<!--#include virtual="Common/functions.inc"-->
<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!-- Bootstrap core CSS -->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Custom styles for this template -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
th {
	vertical-align: middle;
	text-align: center;
	color: #01579B;
	font-size:12px;
}
td {
	vertical-align: middle;
	text-align: center;
}

tr {
	vertical-align: middle;
	text-align: center;
	font-size:9px;
}
black {
	color:black;
}

orange {
	color: darkorange;
}

.panel-override {
	color: black;
	border-radius: 0;
}
.modal-header-success {
    color:#fff;
    padding:9px 15px;
    border-bottom:1px solid #eee;
    background-color: #01579B;
    -webkit-border-top-left-radius: 5px;
    -webkit-border-top-right-radius: 5px;
    -moz-border-radius-topleft: 5px;
    -moz-border-radius-topright: 5px;
     border-top-left-radius: 5px;
     border-top-right-radius: 5px;
}
.modal-header-modal {
    color: white;
    padding: 9px 15px;
    border-bottom: 1px solid #eee;
    background-color: #9a1400;
    -webkit-border-top-left-radius: 5px;
    -webkit-border-top-right-radius: 5px;
    -moz-border-radius-topleft: 5px;
    -moz-border-radius-topright: 5px;
    border-top-left-radius: 5px;
    border-top-right-radius: 5px;
    font-weight: bold;
}
blackText {
	color:black;
	font-weight: 500;
	text-transform: uppercase;
}

.badgeFlags {
    display: inline-block;
    min-width: 10px;
    padding: 3px 6px;
    line-height: 1;
    color: white;
    text-align: center;
    white-space: nowrap;
    vertical-align: baseline;
    background-color: #111;
    border-radius: 14px;
    border: #111;
    /* border-style: double; */
    color: yellow;
    color: yellow;
}
@media screen and (min-width: 550px) {
    .big{
		font-size:	16px;
	}
}
.checkbox, .radio {
    position: relative;
    display: unset;
}
</style>
</head>
<body>
<script>
$(document).ready(function(){
    $('[data-toggle="tooltip"]').tooltip(); 
});	
$('[data-toggle=confirmation]').confirmation({
  rootSelector: '[data-toggle=confirmation]',
  // other options
});
</script>

<script src="https://cdn.jsdelivr.net/npm/bootstrap-confirmation2/dist/bootstrap-confirmation.min.js"></script>
<!--#include virtual="Common/headerMain.inc"-->
	<% if sAction = "Confirm Waivers" then %>
	<div class="container">
		<div class="row">
			<div class="col-md-12 col-sm-12 col-xs-12">
				<table class="table table-custom-black table-bordered table-responsive table-condensed">
					<tr>
						<th>Name</th>
						<th>Action</th>
					</tr>
					<tr>
						<td><%=objRS.Fields("firstname").Value%>&nbsp;<%=objRS.Fields("lastname").Value %></td>
						<td>Request Waivers</td>
					</tr>										
				<table>
      </div>
      <div class="modal-footer">
        <button type="submit" value="Process Waivers;<%=objRS.Fields("PID").Value & ";" & objRS.Fields("firstName").Value & " " & objRS.Fields("lastName").Value%>" name="Action" class="btn btn-danger btn-block"><i class="fa fa-road" aria-hidden="true"></i>&nbsp;Process Waivers</button>
      </div>
				</div>
			</div>
		</div>
	<%end if%>

<form action="playerprofile.asp" name="frmPlayer" method="POST">
  <input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
  <input type="hidden" name="var_tradepartner" value="<%=objRS.Fields("ownerid").Value %>" />
  <input type="hidden" name="pid" value="<%=objRS.Fields("pid").Value %>" />
  <input type="hidden" name="var_pos" value="<%=objRS.Fields("pos").Value %>" />
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
				<table class="table table-no-pad  table-responsive table-bordered">
						<tr style="background-color:white;">
							<td style="vertical-align:middle;width:30%"><img class="img-responsive center-block big" src="<%=objRS.Fields("image").Value %>"></td>
							<td class="big" style="vertical-align:middle;width:40%;font-weight:bold;color: #01579B;"><%=objRS.Fields("firstname").Value%>&nbsp;<%=objRS.Fields("lastname").Value %><br>
							<span class="orange"><%=objRS.Fields("pos").Value %></span><br>
							<%if objRS.Fields("barps").Value > 0 then %>
								<span class="badgeBlue big" style="border-radius: 16px;"><%=round(objRS.Fields("barps").Value,2)%></span>
							<%else%>
								<span class="badgeBlue big" style="border-radius: 16px;">0</span>
							<%end if%>
							</td>
							<td style="vertical-align:middle;width:30%"><img class="img-responsive center-block big" style="height:75px" src="<%=objRS.Fields("teamlogo").Value %>"></td>
						</tr>
					</table>
					<table class="table table-condensed table-bordered">	
						<tr style="background-color:black;font-weight:bold;color:white;font-size:12px;">
						<% if pStatus = "Waivers" Then %>
						<td style="background-color:black;font-weight:bold;color:white;font-size:12px;" width="50%" align="center">STATUS</td>
						<td style="background-color:black;font-weight:bold;color:white;font-size:12px;"width="50%" align="center"><%=pStatus%>&nbsp;<%=objRS.Fields("clearwaiverdate").Value %></td>
						<% else %>
						<td style="background-color:black;font-weight:bold;color:white;font-size:12px;" width="50%" align="center">STATUS</td>
						<td style="background-color:black;font-weight:bold;color:white;font-size:12px;"><%=pStatus%></td>
						<% end if %>	
						</tr>
					
				
					<% if objRS.Fields("lastTeamInd").Value = ownerid then %>
						<tr>
							<td colspan="2" style="font-weight:bold;background-color: yellow;vertical-align: middle;"><i class="fas fa-user-lock red " aria-hidden="true" style="vertical-align: sub;"></i>&nbsp;<span style="font-weight:bold;background-color: yellow;vertical-align: middle;">PLAYER NOT ELIGIBLE TO BE ACQUIRED!&nbsp;<span class="badgeFlags big" style="border-radius: 16px;">LTOP</span></span></td>
						</tr>					
					<%else%>
						<!--SEASON OVER CHECK---->
						<% if w_seasonOver = true then %>
						<tr>
							<td colspan="2" style="font-weight:bold;background-color: yellow;vertical-align: middle;"><i class="fa fa-user-lock red " aria-hidden="true" style="vertical-align: sub;"></i>&nbsp;<span style="font-weight:bold;background-color: yellow;vertical-align: middle;">SEASON IS OVER!&nbsp;<span class="badgeFlags big" style="border-radius: 16px;">SOVR</span></span></td>
						</tr>
						<!--RENTED PLAYER CHECK-->
						<%elseif rentedPlayer = true then%>
						<tr>							
							<td colspan="2" style="font-weight:bold;background-color: yellow;vertical-align: middle;"><i class="fa fa-user-lock red " aria-hidden="true" style="vertical-align: sub;"></i>&nbsp;<span style="font-weight:bold;background-color: yellow;vertical-align: middle;">PLAYER NOT ELIGIBLE TO BE WAIVED!&nbsp;<span class="badgeFlags big" style="border-radius: 16px;">RENT</span></span></td>
						</tr>
						<!--PENDING WAIVER/TRADE CHECK-->			
						<%elseif pendingWaiver = true or pendingTrade = true and (objRS.Fields("ownerid").Value = ownerid) then%>
						<tr>
							<td colspan="2" style="font-weight:bold;background-color: yellow;vertical-align: middle;"><i class="fa fa-user-lock red " aria-hidden="true" style="vertical-align: sub;"></i>&nbsp;<span style="font-weight:bold;background-color: yellow;vertical-align: middle;">PLAYER NOT ELIGIBLE TO BE WAIVED!&nbsp;<span class="badgeFlags big" style="border-radius: 16px;">PWPT</span></span></td>
						</tr>	
						<!--PLAYER IN PLAY CHECK-->			
						<% elseif objRS.Fields("ownerid").Value = ownerid Then %>
							<%if wPlayerInPlay = True then%>
							<tr>
								<td colspan="2" style="font-weight:bold;background-color: yellow;vertical-align: middle;"><i class="fas fa-basketball-ball  fa-spin " data-toggle="tooltip" title="Playing!"></i>&nbsp;<span style="font-weight:bold;background-color: yellow;">PLAYER IN ACTION</span>&nbsp;<i class="fas fa-basketball-ball fa-spin " data-toggle="tooltip" title="Playing!"></i></td>
							<%else%>	
								<td colspan="2"><button type="button" class="btn btn-danger btn-block"  data-toggle="modal" data-target="#confirmWaivers"><i class="fas fa-truck-moving " style="color: white;"></i>&nbsp;WAIVE</button></td>
							<% end if %>
							</tr>
						<!--PLAYER STATUS CHECK-->				
						<% elseif pStatus = "Waivers" or pStatus = "Free Agent" or pStatus = " " or pStatus = "Staggered" Then %>
							<tr>
								<td colspan="2"><button type="submit" value="Transaction" name="Action" class="btn btn-trades  btn-block"><i class="fas fa-plus"></i>&nbsp;AQUIRE</button></td>	
							</tr>
						<!--PLAYER TRADE STATUS CHECK-->					
						<% elseif objRS.Fields("ownerid").Value > 0 and objRS.Fields("acceptTradeOffers").Value = true and myTradeInd = true Then %>
							<tr>						
								<td colspan="2"><button type="submit" value="Counter" name="Action" class="btn btn btn-trades btn-block "><i class="fas fa-exchange-alt "></i>&nbsp;SUBMIT TRADE OFFER</button></td> 
							</tr>
						<% elseif objRS.Fields("ownerid").Value > 0 and objRS.Fields("acceptTradeOffers").Value = false or myTradeInd = false Then %>
							<tr>
								<td colspan="2" style="font-weight:bold;background-color: yellow;vertical-align: middle;"><i class="fas fa-user-lock red " style="vertical-align: sub;"></i>&nbsp;<span style="font-weight:bold;background-color: yellow;vertical-align: middle;">TRADE DESK IS CLOSED!</span>&nbsp;<i class="fas fa-user-lock red " style="vertical-align: sub;"></i></td>
							</tr>
						<% end if %>					
					<%end if%>
						
						<!--IR INDICATOR CHECK---->
						<%if objRS.Fields("ownerid").Value = ownerid and rentedPlayer = false and  w_seasonOver = fale Then %>
							<tr>
								<td colspan="2">
									<%if objRS.Fields("IR").Value = true then %>
									<button type="submit" value="Injury;<%=objRS.Fields("PID").Value & ";" & objRS.Fields("firstName").Value & " " & objRS.Fields("lastName").Value%>" name="Action" class="btn btn-default-red btn-block btn-md"><i class="fas fa-briefcase-medical red fa-md"></i>&nbsp;TURN OFF INJURY INDICATOR</button></button>						
								<%else%>
									<button type="submit" value="Injury;<%=objRS.Fields("PID").Value & ";" & objRS.Fields("firstName").Value & " " & objRS.Fields("lastName").Value%>" name="Action" class="btn btn-default-red btn-block btn-md"><i class="fas fa-briefcase-medical red fa-md"></i>&nbsp;TURN ON	 INJURY INDICATOR</button></button>
								<%end if%>		
								</td>
							</tr>
						<%end if%>
					</table>
					</br>
					<table class="table table-custom-black table-responsive table-condensed table-bordered">
						<tr style="background-color:#ddd;" class="text-uppercase">
							<th align="center" width="12%">usg</th>
							<th align="center" width="12%">3D</th>
							<th align="center" width="9%"><i class="fas fa-clock"></i></th>		
							<th align="center" width="45%">Games Left</th>
							<th align="center" width="9%">GP</th>
							<th align="center" width="13%"><i class="fas fa-usd-circle"></i></th>								
						</tr>
						<tr style="background-color:white;color:black;">
							<td align="center"><span class="big"><%=objRS.Fields("usage").Value%></span></td>
							<td align="center"><span class="big"><%=objRS.Fields("numTdbls").Value %></span></td>										
							<td align="center"><span class="big"><%=objRS.Fields("min").Value %></span></td>	
							<td align="center"><span class="big"><span style="color:#468847;font-weight: bold;">REG:</span> <%=playerGmCnt %> - <span style="color:#468847;font-weight: bold;">PO:</span> <%=playerPOGmCnt%></span></td>
							<td align="center"><span class="big"><%=objRS.Fields("gp").Value %></span></td>	
							<td align="center"><span class="auctionText big"><strong><%=FormatCurrency(objRS.Fields("auctionPrice").Value)%></strong></span></td>
						</tr>
						<!--<% if espnLink <> "" then %>
						<tr style="background-color:white;">
							<td class="big" colspan="6" style="vertical-align:middle;text-align:center"><a class="blue" href="<%= espnLink %>"_self">View ESPN Player Card</bluePos></td>
						</tr>
						<%end if%>-->
					</table>
					<table class="table table-custom-black table-bordered table-responsive table-condensed">
						<thead>
						<tr>
							<th colspan="10">Last 5 Game Stats</th>
						</tr>
						</thead>
					<tr style="background-color:#ddd;font-weight:bold;color:grey;">
						<th class="big" style="width:12%"><span style="font-weight:bold;"<i class="fas fa-calendar-alt"></i></span></th>						
						<th class="big" style="width:12%"><span style="font-weight:bold;color:#a5a6a7;">VS</span></th>
						<th class="big" style="width:9%"><span style="font-weight:bold;color:#a5a6a7;">B</th>
						<th class="big" style="width:9%"><span style="font-weight:bold;color:#a5a6a7;">A</th>
						<th class="big" style="width:9%"><span style="font-weight:bold;color:#a5a6a7;">R</th>
						<th class="big" style="width:9%"><span style="font-weight:bold;color:#a5a6a7;">P</th>
						<th class="big" style="width:9%"><span style="font-weight:bold;color:#a5a6a7;">S</th>
						<th class="big" style="width:9%"><span style="font-weight:bold;color:#a5a6a7;">3</th>
						<th class="big" style="width:9%"><span style="font-weight:bold;color:#a5a6a7;">T</th>						
						<th class="big" style="width:13%"><span style="font-weight:bold;"><i class="fas fa-basketball-ball"></i></span></th>
					</tr>
				<% While last5cnt <= last5LoopCtrl
				
					tbonus = 0
				
					if Cint(objRSLast5.Fields("blk").Value) >= 10 then 
						tbonus = tbonus + 1
					end if

					if Cint(objRSLast5.Fields("ast").Value) >= 10 then 
						tbonus = tbonus + 1
					end if

					if Cint(objRSLast5.Fields("trb").Value) >= 10 then 
						tbonus = tbonus + 1
					end if

					if Cint(objRSLast5.Fields("pts").Value) >= 10 then 
						tbonus = tbonus + 1
					end if

					if Cint(objRSLast5.Fields("stl").Value) >= 10 then 
						tbonus = tbonus + 1
					end if

					if Cint(objRSLast5.Fields("x3p").Value) >= 10 then 
						tbonus = tbonus + 1
					end if
											
					%>						
					<tr style="color:black;background-color:white;text-align:center;">		
						<td class="big"><%=objRSLast5.Fields("xDate").Value %></td>
						<td  nowrap><%=objRSLast5.Fields("opp").Value %></td>
						<% if tbonus >= 3 and CInt(objRSLast5.Fields("blk").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;font-weight:bold;border: black;border-style:solid;border-width:thin;"><%=(objRSLast5.Fields("blk").Value)%></td>
						<%else%>
							<td class="big"><%=(objRSLast5.Fields("blk").Value)%></td>
						<%end if%>						
						
						<% if tbonus >= 3 and CInt(objRSLast5.Fields("ast").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;font-weight:bold;border: black;border-style:solid;border-width:thin;color:white;"><%=(objRSLast5.Fields("ast").Value)%></td>
						<%else%>
							<td class="big"><%=(objRSLast5.Fields("ast").Value)%></td>
						<%end if%>	
	
						<% if tbonus >= 3 and CInt(objRSLast5.Fields("trb").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;font-weight:bold;border: black;border-style:solid;border-width:thin;color:white;"><%=(objRSLast5.Fields("trb").Value)%></td>
						<%else%>
							<td class="big"><%=(objRSLast5.Fields("trb").Value)%></td>
						<%end if%>
						
						<% if tbonus >= 3 and CInt(objRSLast5.Fields("pts").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;font-weight:bold;border: black;border-style: solid;border-width:thin;color:white;"><%=(objRSLast5.Fields("pts").Value)%></td>
						<%else%>
							<td class="big"><%=(objRSLast5.Fields("pts").Value)%></td>
						<%end if%>

						<% if tbonus >= 3 and CInt(objRSLast5.Fields("stl").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;font-weight:bold;border: black;border-style: solid;border-width:thin;color:white;"><%=(objRSLast5.Fields("stl").Value)%></td>
						<%else%>
							<td class="big"><%=(objRSLast5.Fields("stl").Value)%></td>
						<%end if%>						

						<% if tbonus >= 3 and CInt(objRSLast5.Fields("x3p").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;font-weight:bold;border: black;border-style: solid;border-width:thin;color:white;"><%=(objRSLast5.Fields("x3p").Value)%></td>
						<%else%>
							<td class="big"><%=(objRSLast5.Fields("x3p").Value)%></td>
						<%end if%>	
							<td class="big"><%=objRSLast5.Fields("tov").Value %></td>
							<td class="big" style="font-weight:bold;background-color: black; color: darkorange;"><%=objRSLast5.Fields("barptot").Value %></td>
					</tr>
				<% 
				objRSLast5.MoveNext	
				last5cnt = last5cnt + 1				
				wend 
				%>
				<br>
				<thead>
					<tr>
						<th colspan="10">Averages</th>
					</tr>
				</thead>				
					<tr style="color: #a5a6a7;">
						<td class="big" style="background-color:black;color:white;font-weight:bold;" colspan="2">Stats for</td>
					  <th class="big"><span style="color:#a5a6a7;">B</th>
						<th class="big"><span style="color:#a5a6a7;">A</th>
						<th class="big"><span style="color:#a5a6a7;">R</th>
						<th class="big"><span style="color:#a5a6a7;">P</th>
						<th class="big"><span style="color:#a5a6a7;">S</th>
						<th class="big"><span style="color:#a5a6a7;">3</th>
						<th class="big"><span style="color:#a5a6a7;">T</th>
						<th class="big"><span><i class="fas fa-basketball-ball"></i></span></th>
					</tr>
					<tr style="color:black;background-color:white;text-align:center;">		
						<td colspan="2" class="big"><strong>Last 5</td>
						<td class="big"><%= round(avgBlks,0)%></td>	
						<td class="big"><%= round(avgAst,0) %></td>
						<td class="big"><%= round(avgReb,0) %></td>
						<td class="big"><%= round(avgPts,0) %></td>
						<td class="big"><%= round(avgStl,0) %></td>
						<td class="big"><%= round(avg3pt,0) %></td>
						<td class="big"><%= round(avgTo,0)%></td>
						<% if CDbl(objRS.Fields("l5barps").Value) > CDbl(objRS.Fields("barps").Value) then %>
							<td class="big">
							<span class="badgeUp big"><%= round(avgBarps,2) %></span>
						<% elseif CDbl(objRS.Fields("barps").Value) > CDbl(objRS.Fields("l5barps").Value) then%>
							<td class="big">
							<span class="badgeDown big"><%= round(avgBarps,2) %></span>
						<%else%>
							<td class="big">
							<span class="badgeEven big"><%= round(avgBarps,2) %></span>
						<%end if %>							
						</td>						
					</tr>
					<tr style="color:black;background-color:white;text-align:center;">		
						<td class="big" colspan="2"><strong>Season</td>
						<td class="big"><%=objRS.Fields("blk").Value%></td>
						<td class="big"><%=objRS.Fields("ast").Value%></td>
						<td class="big"><%=objRS.Fields("reb").Value%></td>
						<td class="big"><%=objRS.Fields("ppg").Value%></td>
						<td class="big"><%=objRS.Fields("stl").Value%></td>
						<td class="big"><%=objRS.Fields("three").Value%></td>
						<td class="big"><%=objRS.Fields("to").Value%></td>
						<td class="big"><span class="badgeBlue big"><%=round(objRS.Fields("barps").Value,2)%></span></td>
					</tr>				
			</table>
			</br>
			</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<table class="table table-custom-black 	ztable-condensed table-bordered">
				<tr>
					<th colspan="5" align="center"	>Next 5 IGBL Games</th>
				</tr>
				<tr style="background-color:white;">
				<%While skedcnt <= 4 %>
					<td class="big" align="center" style="width: 20%;"><%=objRSNext5.Fields("opponent").Value%><br><%=objRSNext5.Fields("gameday").Value%></td>
				<%
				objRSNext5.MoveNext
				skedcnt = skedcnt + 1
				Wend
				%>
				</tr>
			</table>	
			</br>		
		</div>
	</div>
</div>

</form>
<!--MODAL WAIVERS-->
<div class="container">
	<div class="row">
	<div class="col-md-12 col-sm-12 col-xs-12">
		<div id="confirmWaivers" class="modal fade" role="dialog">
			<div class="modal-dialog" role="document">
							
				<div class="modal-content">
					<div class="modal-header modal-header-modal">
						<button type="button" class="close" data-dismiss="modal">&times;</button>
						<h3 class="modal-title">Confirm Waivers</h3>
					</div>
					<div class="modal-body">
					<form action="playerprofile.asp" name="frmPlayer" method="POST">
						<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
						<input type="hidden" name="pid" value="<%=objRS.Fields("pid").Value %>" />
						<table class="table table-custom-black table-bordered table-responsive table-condensed">
							<tr>
								<th>Name</th>
								<th>Action</th>
							</tr>
							<tr>
								<td class="red"><%=objRS.Fields("firstname").Value%>&nbsp;<%=objRS.Fields("lastname").Value %></td>
								<td class="red">Request Waivers</td>
							</tr>										
						</table>
					</div>		
					<div class="modal-footer">
						<button type="submit" value="Process Waivers;<%=objRS.Fields("PID").Value & ";" & objRS.Fields("firstName").Value & " " & objRS.Fields("lastName").Value%>" name="Action" class="btn btn-danger btn-block"><i class="fa fa-road" aria-hidden="true"></i>&nbsp;Process Waivers</button>
					</div>
					</div>
					</div>
					</form>
				</div>
			</div>
		</div>
	</div>

</body>

<% 
objRSNBASked.Close
objRS.close
objRSNext5.close
objRSReg.close
objRSPO.close
objRSLast5.close
set objRS = Nothing
set objRSNext5 = Nothing
set objRSReg = Nothing
set objRSPO = Nothing
 %>
</html>