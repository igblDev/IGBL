<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""
	
  dim sTeam, sAction, sURL,tradePartner,ownerid,playerCount, playerCount2,objConn,objRSTeam1, objRSTeam2,playercnt
	dim objrstrade, PID_Split,PID_Split2,iRecordToUpdate,iRecordToUpdate2,strSQL,objRStrades,ownername,tradepartneremail,teamname
	dim errorcode,errordesc,owneremail,objRSCenters,objRSForwards,objRSGuards,objParams
	
	'EMAIL VARIABLES
	Dim email_to, email_subject, host, username, password, reply_to, port, from_address
	Dim first_name, last_name, home_address, email_from, telephone, comments, error_message,ppPID
	Dim ObjSendMail, email_message

	Set objConn        = Server.CreateObject("ADODB.Connection")
	Set objRSTeam1     = Server.CreateObject("ADODB.RecordSet")
	Set objRSTeam2     = Server.CreateObject("ADODB.RecordSet")
	Set objRS1         = Server.CreateObject("ADODB.RecordSet")
	Set objRStrades    = Server.CreateObject("ADODB.RecordSet")
	Set objRSteamSelect= Server.CreateObject("ADODB.RecordSet")
	Set objrstrade     = Server.CreateObject("ADODB.RecordSet")
	Set objRSCenters   = Server.CreateObject("ADODB.RecordSet")
	Set objRSForwards  = Server.CreateObject("ADODB.RecordSet")
	Set objRSGuards    = Server.CreateObject("ADODB.RecordSet") 
	Set objParams      = Server.CreateObject("ADODB.RecordSet")
	

	
	objConn.Open Application("lineupstest_ConnectionString")
	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"
				  
	%>
	<!--#include virtual="Common/session.inc"-->
	<%
	
	objParams.Open "SELECT * FROM tblParameterCtl where param_name = 'TRADE_DEADLINE' ",objConn
	wTradeDeadLine = objParams.Fields("param_date").Value
	objParams.Close
		
	GetAnyParameter "Action", sAction
	GetAnyParameter "cmbTeam", sTeam
	tradepartner = sTeam
	GetAnyParameter "var_sPid", ppPID
	
	'***THE REVISE ACTION IS USED FOR THE REVISE FUNCTION IN PENDING ANALYZED TRADE***'
	if sAction = "Continue" or Action = "Revise" then
		GetAnyParameter "cmbTeam", sTeam
		tradepartner = sTeam
		GetAnyParameter "var_sPid", ppPID	
		'Response.Write "PID SELECTED FROM PROFILE " & ppPID
	else 
		TEAM_Split        = Split(Request.Form("Action"), ";")
		sAction           = TEAM_Split(0)
		tradepartner      = TEAM_Split(1) 'Team ID
		sTeam             = TEAM_Split(1) 'Team
	end if
	
	objRSCenters.Open   "SELECT * FROM tblPlayers WHERE (tblPlayers.POS  = 'CEN' or tblPlayers.POS  = 'F-C') and tblPlayers.OwnerID=" & tradepartner & "",objConn,3,3,1
	objRSForwards.Open  "SELECT * FROM tblPlayers WHERE (tblPlayers.POS  = 'FOR' or tblPlayers.POS  = 'F-C' or tblPlayers.POS   = 'G-F') and tblPlayers.OwnerID=" & tradepartner & "",objConn,3,3,1
	objRSGuards.Open  	"SELECT * FROM tblPlayers WHERE (tblPlayers.POS  = 'GUA' or tblPlayers.POS  = 'G-F') and tblPlayers.OwnerID=" & tradepartner & "",objConn,3,3,1
	
	centerCnt= objRSCenters.RecordCount
	'Response.Write "CENTER COUNT " & centerCnt
	forwardCnt = objRSForwards.RecordCount
	guardCnt = objRSGuards.RecordCount
	
	objRSCenters.Close  	
	objRSForwards.Close  
	objRSGuards.Close 

	objRSCenters.Open   "SELECT * FROM tblPlayers WHERE (tblPlayers.POS  = 'CEN' or tblPlayers.POS  = 'F-C') and tblPlayers.OwnerID=" & ownerid & "",objConn,3,3,1
	objRSForwards.Open  "SELECT * FROM tblPlayers WHERE (tblPlayers.POS  = 'FOR' or tblPlayers.POS  = 'F-C' or tblPlayers.POS   = 'G-F') and tblPlayers.OwnerID=" & ownerid & "",objConn,3,3,1
	objRSGuards.Open  	"SELECT * FROM tblPlayers WHERE (tblPlayers.POS  = 'GUA' or tblPlayers.POS  = 'G-F') and tblPlayers.OwnerID=" & ownerid & "",objConn,3,3,1
	
	centerCntMe = objRSCenters.RecordCount
	forwardCntMe = objRSForwards.RecordCount
	guardCntME = objRSGuards.RecordCount
	
	objRSCenters.Close  	
	objRSForwards.Close  
	objRSGuards.Close  
	
	if ppPID = "" then
		ppPID = 0
	end if 
	
	select case sAction
	  	case "Submit Trade Offer"

	    ownerid           = Request.Form("var_ownerid")
	    tradepartner      = Request.Form("cmbteam2")
			tradepartneremail = Request.Form("var_tradepartneremail")
			tradepartnername  = Request.Form("var_tradepartnername")
			owneremail        = Request.Form("var_owneremail")
			teamname          = Request.Form("var_teamname")
			playercnt         = Request.Form("var_playercnt")
			playerCount       = Request.Form("cmbTrade1").Count
			playerCount2      = Request.Form("cmbTrade2").Count

			dim w_loop_count
			w_loop_count = 0

			if playerCount > w_loop_count then
				w_loop_count = playerCount
			end if

			if playerCount2 > w_loop_count then
				w_loop_count = playerCount2
			end if
			dim  I

			dim trade1,trade2,trade3,acquired1,acquired2,acquired3
			dim tplayername1, tplayername2, tplayername3, aplayername1, aplayername2, aplayername3

			trade2 = 0
			trade3 = 0
			acquired2 = 0
			acquired3  = 0

			For I = 1 To w_loop_count
				PID_Split = Split(Request.Form("cmbTrade1")(I), ";")
				PID_Split(0)
				PID_Split(1)

				PID_Split2 = Split(Request.Form("cmbTrade2") (I), ";")
				PID_Split2(0)
				PID_Split2(1)

			if I = 1 then
				trade1 = PID_Split(0)
				tplayername1 = PID_Split(1)
				acquired1 = PID_Split2(0)
				aplayername1 = PID_Split2(1)
				msgtradeplayers = tplayername1 & "<br />"
				msgaquireplayers = aplayername1 & "<br />"

			elseif I = 2 then
				if playerCount >= I then
					trade2 = PID_Split(0)
					tplayername2 = PID_Split(1)
					msgtradeplayers = msgtradeplayers  & tplayername2 & "<br />"
				end if

				if PlayerCount2 >= I then
					acquired2 = PID_Split2(0)
					aplayername2 = PID_Split2(1)
					msgaquireplayers = msgaquireplayers  & aplayername2 & "<br />"
				end if
			else
				if playerCount >= I then
					trade3 = PID_Split(0)
					tplayername3 = PID_Split(1)
					msgtradeplayers = msgtradeplayers  & tplayername3 & "<br />"
				end if

				if PlayerCount2 >= I then
					acquired3 = PID_Split2(0)
					aplayername3 = PID_Split2(1)
					msgaquireplayers = msgaquireplayers  & aplayername3 & "<br />"
										
				end if
			end if

			next

			errorcode = False

			if ((playercnt - playerCount) + playerCount2 ) > 14 then
				errorcode = true
			else

				objrstrade.Open	"SELECT * " & _
		        "FROM tblTradeAnalysis " & _
        		 "WHERE  FromOid = "& ownerid & "   " & _
        		    "and ToOid  = "& tradepartner & "   " & _
								"and tradedplayerid      = "& trade1  & "   " & _
								"and tradedplayerid2   = "& trade2  & "   " & _
								"and tradedplayerid3   = "& trade3  & "   " & _
								"and acquiredplayerid  = "& acquired1  & "  " & _
								"and acquiredplayerid2 = "& acquired2 & " " & _
								"and acquiredplayerid3 = "& acquired3  & " ", objConn,3,3,1


					if objrstrade.RecordCount > 0 then
						errorcode = "Invalid Offer"
					end if

					objrstrade.close

					if errorcode = False then
			    	w_errorCt = 0
					'Response.Write "Error Count = "&w_errorCt&" <br> "

					'** CHECK MY PLAYERS
					w_RetCd = Player_Owner_Relationship(trade1, ownerid, w_errorCt)
					w_RetCd = Player_Owner_Relationship(trade2, ownerid, w_errorCt)
					w_RetCd = Player_Owner_Relationship(trade3, ownerid, w_errorCt)

					'** CHECK TRADING PARTNER PLAYERS
					w_RetCd = Player_Owner_Relationship(acquired1, tradepartner, w_errorCt)
					w_RetCd = Player_Owner_Relationship(acquired2, tradepartner, w_errorCt)
					w_RetCd = Player_Owner_Relationship(acquired31, tradepartner, w_errorCt)
					'	Response.Write "Error Count = "&w_errorCt&" <br> "


 					if w_errorCt > 0 then
	 					errorcode = "Invalid Offer"
					end if
			end if

			if errorcode = False then
				objrstrade.close
				errorcode = False
          strSQL ="insert into tblTradeAnalysis(FromOid,TradedPlayerID,TradedPlayerID2,TradedPlayerID3,ToOid,AcquiredPlayerID,AcquiredPlayerID2,AcquiredPlayerID3) values ('" &_
					ownerid & "', '" &  trade1 & "', '" & trade2 & "', '" &  trade3 & "','" &  tradepartner & "','" &  acquired1 & "', '" & acquired2 & "', '" &  acquired3 & "')"					
					objConn.Execute strSQL

	  			'**********************************************************
					Response.Redirect "pendingAnalyzedTrades.asp?ownerid=" & ownerid

				end if

			end if


    	case "Continue"
		
				GetAnyParameter "var_tradeid", tradeID
				Response.Write var_tradeid
				'Response.Write tradePartner
    	
			case "Revise"
		
				GetAnyParameter "var_tradeid", tradeID
				Response.Write var_tradeid
				'Response.Write tradePartner		

   		case "Cancel"
			sURL = "lineups.asp"
			AddLinkParameter "var_ownerid", ownerid, sURL
    		Response.Redirect sURL

  		case ""
	    	'First time to the screen
			ownerid = session("ownerid")	
			
			if ownerid = "" then
				GetAnyParameter "var_ownerid", ownerid
			end if
	end select

	'***********************************************************
	' Player_Owner_Relationship()
	' This function searchs the tblPlyaers table to ensure the that
	' player is on this owners team.  if the player is on this owners
	' team then 0 is returned, else 1 is returned.  if playerID 0
	' is passed, then 0 is returned
	'***********************************************************
	Function Player_Owner_Relationship (p_PlayerID, p_OwnerID, p_errorCt)

		dim l_returnValue

		'Response.Write "Player ID = "&p_PlayerID&"  OwnerID = "&p_OwnerID&" <br> "

		if p_PlayerID = 0 then
			l_returnValue = 0
		else
			objrstrade.Open	"SELECT * " & _
	        "FROM tblplayers " & _
					"WHERE PID = "& p_PlayerID & "   " & _
					"AND OWNERID = "& p_OwnerID & " ", objConn,3,3,1

			if objrstrade.RecordCount = 0 then
				l_returnValue = 1
			else
				l_returnValue = 0
			end if

			objrstrade.close
		end if

		'Response.Write "l_returnValue = "&l_returnValue&" <br> "
		p_errorCt = p_errorCt + l_returnValue

	end Function


   	objRSTeam1.Open  "SELECT * FROM qrytradeplayers WHERE (((qrytradeplayers.OwnerID)=" & ownerid & "))", objConn,3,3,1
    objRSTeam2.Open  "SELECT * FROM qrytradeplayers WHERE (((qrytradeplayers.OwnerID)=" & tradePartner & "))", objConn,3,3,1
 		objRSteamSelect.Open  "SELECT * from qryTeamTrades WHERE (((qryTeamTrades.OwnerID)<>  " & ownerid & "))", objConn,3,3,1
		myPlayerCnt  = objRSTeam1.recordcount
		hisPlayerCnt = objRSTeam2.recordcount

%>

<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!-- Bootstrap core CSS -->
<!--#include virtual="Common/bootstrap.inc"-->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style type="text/css">
th {
    vertical-align:middle;
  	text-align:center
}
td {
    vertical-align:middle;
}

black {
	color:black;
}
white {
	color:white;
	font-weight: bold;
}
.alert-danger {
    color: #ffffff;
    background-color: red;
    border-color: #111;
}

.panel-title {
    color: yellowgreen;
    text-transform: none;
    font-size: 14px !important;
}
.panel-heading {
    background-image: none;
    background-color: black !important;
    color: white;
    height: 30px;
    padding: 5px 5px;
		border-radius: unset;
}
.alert-warning {
    color: #f2f2f2;
    background-color: #9a1400;
    border-color: black;
}
.btn-default-list {
    color: #fff;
    background-color: #01579B;
    border-color: #777 !important;
    font-weight: bold !important;
    border-style: outset !important;
    border-width: medium !important;
}
.btn-default-list {
    color: #000;
    background-color: yellowgreen;
    border-color: #777 !important;
    font-weight: bold !important;
    border-style: outset !important;
    border-width: medium !important;
}
.btn-default {
    color: black;
    background-color: #f2f2f2;
    font-weight: bold !important;
	border-style: outset !important;
    border-width: thin !important;
}
.h5, .h6, h5, h6 {
    font-family: inherit;
    font-weight: 500;
    line-height: 1.1;
    color: inherit;
    text-align: center;
}
.h4, h4 {
    font-size: 16px;
    color: #212121;
    font-weight: 600;
		text-transform: uppercase;
}

.panel-override {
    color: black;
    background-color: #dddddd;
    border-color: black;
    border-radius: 0px;
}
.table-bordered>tbody>tr>td, .table-bordered>tbody>tr>th, .table-bordered>tfoot>tr>td, .table-bordered>tfoot>tr>th, .table-bordered>thead>tr>td, .table-bordered>thead>tr>th {
    border: 1px solid #000;
}
.btn-warning {
    color: #111;
    background-color: yellow;
    border-color: black;
}
</style>
</head>
<body>
<script language="JavaScript" type="text/javascript"><!--
function processTrades(theForm) {
		
		/*Edits for first team */	
		var tradingPlayers = 0;
		for(var i=0; i < theForm.cmbTrade1.length; i++){
			if(theForm.cmbTrade1[i].checked) {
			tradingPlayers +=1;
			}
		}

		if(tradingPlayers > 3) {
			alert("Three Player Trade Limit Exceeded!");
		return (false);
		}

		if(tradingPlayers == 0) {
			alert("Select Min 1 : Max 3 Player(s)");
		return (false);
		}


		/*Edits for second team */

		var tradingPlayers2 = 0;
		for(var i=0; i < theForm.cmbTrade2.length; i++){
			if(theForm.cmbTrade2[i].checked) {
			tradingPlayers2 +=1;
			}
		}

		if(tradingPlayers2 == 0) {
			alert("Select Min 1 : Max 3 Player(s)");
		return (false);
		}

		if(tradingPlayers2 > 3) {
			alert("Three Player Trade Limit Exceeded!");
		return (false);
		}
}

$(document).ready(function(){
    $('[data-toggle="tooltip"]').tooltip(); 
});

//--></script>
<% if sAction = "" then %>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.cmbTeam.selectedIndex == 0)
  {
    alert("The first \"Team Name\" option is not a valid selection.  Please choose one of the other options.");
    theForm.cmbTeam.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan -->
<form action="tradeanalyzer.asp" method="POST" language="JavaScript" name="FrontPage_Form1" onSubmit="return FrontPage_Form1_Validator(this)">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<!--#include virtual="Common/headerMain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong><i class="far fa-exchange"></i>&nbsp;Select Trade Partner</strong>
			</div>
		</div>
	</div>
</div>
<% if (wTradeDeadLine + 1 + 1/24) < now() then %>  
<div class="container">
	<div class="row">		
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-warning">	
				<strong><i class="fa fa-exclamation-triangle" aria-hidden="true"></i>&nbsp;TRADE-DEADLINE HAS PASSED </strong>
			</div>
		</div>
	</div>
</div>
<% else %>  
<div class="container">
	<div class="row">		
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-warning ">	
				<strong>TRADE-DEADLINE: </strong> <%=(FormatDateTime(wTradeDeadLine,1))%>
			</div>
		</div>
	</div>
</div>
<% end if %>  
<div class="container">
	<div class="row">		
		<div class="col-md-12 col-sm-12 col-xs-12">			
			<%
			While Not objRSteamSelect.EOF
			%>				
				<%if  objRSteamSelect.Fields("acceptTradeOffers").value = true then  %>
						<button class="btn  btn-default btn-block" value="Continue;<%= Trim(objRSteamSelect.Fields("OwnerID").value) %>;<%= objRSteamSelect.Fields("TeamName").value%>" name="Action" type="submit"><%= objRSteamSelect.Fields("TeamName").value%></button>
				<% end if%>					
			<%
			objRSteamSelect.MoveNext
			Wend
			%>			
		</div>
	</div>
</div>
</form>
<%
end if
if sAction = "Continue" or sAction = "Revise" then %>
<!--#include virtual="Common/headerMain.inc"-->
<form action="tradeanalyzer.asp" method="POST" onSubmit="return processTrades(this)" language="JavaScript">
  <input type="hidden" name="var_playercnt" value="<%=objRSTeam1.Fields("ActivePlayerCnt").Value%>" />
  <input type="hidden" name="var_teamname" value="<%=objRSTeam1.Fields("TeamName").Value%>" />
  <input type="hidden" name="var_owneremail" value="<%=objRSTeam1.Fields("HomeEmail").Value%>" />
  <input type="hidden" name="var_tradepartneremail" value="<%=objRSTeam2.Fields("HomeEmail").Value%>" />
  <input type="hidden" name="var_tradepartnername" value="<%=objRSTeam2.Fields("teamname").Value%>" />
  <input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
  <input type="hidden" name="cmbTeam2" value="<%= sTeam %>" />
  <% 


		tradeCnt  = 0
		
		'***********************************************************************
		'** The tblPendingTrades is used for the COUNTER trade function
		'** The tblTradeAnalysis is used for the REVISE trade function
		'***********************************************************************
	 	
		if sAction = "Continue" then
			objrstrade.Open	"SELECT * FROM tblPendingTrades WHERE tblPendingTrades.tradeid = "&tradeID&"", objConn,1,1
			tradeCnt  = objrstrade.RecordCount		
		elseif sAction = "Revise" then
			objrstrade.Open	"SELECT * FROM tblTradeAnalysis WHERE tblTradeAnalysis.tradeid = "&tradeID&"", objConn,1,1
			tradeCnt  = objrstrade.RecordCount
		end if
		
		if tradeCnt > 0 then		
			AcquiredPlayerID1 = objrstrade.Fields("AcquiredPlayerID").Value
			AcquiredPlayerID2 = objrstrade.Fields("AcquiredPlayerID2").Value
			AcquiredPlayerID3 = objrstrade.Fields("AcquiredPlayerID3").Value	
			TradedPlayerID1   = objrstrade.Fields("TradedPlayerID").Value
			TradedPlayerID2   = objrstrade.Fields("TradedPlayerID2").Value
			TradedPlayerID3   = objrstrade.Fields("TradedPlayerID3").Value
		else
			AcquiredPlayerID1 = 0
			AcquiredPlayerID2 = 0
			AcquiredPlayerID3 = 0	
			TradedPlayerID1   = 0
			TradedPlayerID2   = 0
			TradedPlayerID3   = 0 
		end if
		
		'Response.Write "Trade ID           = : " & tradeID  & "  <br>"
		'Response.Write "Trade CNT          = : " & tradeCnt  & "  <br>"
		'Response.Write "Acquired PLayer 1  = : " & AcquiredPlayerID1  & "  <br>"		
		'Response.Write "Acquired PLayer 3  = : " & AcquiredPlayerID3  & "  <br>"		
		'Response.Write "Acquired PLayer 2  = : " & AcquiredPlayerID2  & "  <br>"		
		'Response.Write "Traded Player 1    = : " & TradedPlayerID1  & "  <br>"  
		'Response.Write "Traded Player 2    = : " & TradedPlayerID2  & "  <br>" 
		'Response.Write "Traded Player 3    = : " & TradedPlayerID3  & "  <br>" 
		'Response.Write "sAction            = : " & sAction  & "  <br>"

		%>
<% if (wTradeDeadLine + 1 + 1/24) < now() then %>  
<div class="container">
	<div class="row">		
    <div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-danger">
				<strong><i class="fa fa-exclamation-triangle" aria-hidden="true"></i>TRADE-DEADLINE HAS PASSED </strong>
			</div>
		</div>
	</div>
</div>
<% end if %>  
  
<div class="container ">
	<div class="row">
		<div class="col-md-6">
				<div class="panel panel-override">
				<div class="panel-heading clearfix">
					<h5 class="panel-title"><%=objRSTeam2.Fields("TeamName").Value %></h5>
				</div>
					<table class="table table-custom-black table-responsive table-bordered table-condensed">
						<tr style="background-color:yellowgreen;">
							<td colspan="5">
								<table class="table table-custom table-responsive table-bordered table-condensed">
									<tr style="text-align:center";>
										<th style="color:black" colspan="6">Roster Configuration <small class="pull-right">Player Count: <%=hisPlayerCnt%></th>
									</tr>
									<tr style="background-color:black;text-align:center;">
										<td class ="big" style="color:#ddd;font-weight:bold;width:12%">CEN</td>
										<td class ="big" style="background-color:white" width="12%"><%=centerCnt%></td>
										<td class ="big" style="color:#ddd;font-weight:bold;width:12%">FOR</td>
										<td class ="big" style="background-color:white" width="12%"><%=forwardCnt%></td>
										<td class ="big" style="color:#ddd;font-weight:bold;width:12%">GUA</td>
										<td class ="big" style="background-color:white" width="12%"><%=guardCnt%></td>
										</tr>
								</table>
							</td>
						</tr>
						<tr>
							<th style="border-radius:0px;text-align:center;white-space:nowrap;width:10%;"><i class="fas fa-basketball-ball"></i></th>
							<th style="border-radius:0px;text-align:left;width:40%;">Player - TM | Pos</th>
							<th style="border-radius:0px;text-align:center;width:20%;">AVG</th>
							<th style="border-radius:0px;text-align:center;width:20%;">L/5</th>
							<th style="border-radius:0px;text-align:center;width:10%;">OTB</th>
							</tr>
						<tbody>
						<%  While Not objRSTeam2.EOF %>
						<% if cint(ppPID) = cint(objRSTeam2.Fields("PID").Value)  or cint(objRSTeam2.Fields("PID").Value) = cint(TradedPlayerID1)or cint(objRSTeam2.Fields("PID").Value) = cint(TradedPlayerID2) or cint(objRSTeam2.Fields("PID").Value) = cint(TradedPlayerID3) or cint(objRSTeam2.Fields("PID").Value) = cint(AcquiredPlayerID1) or cint(objRSTeam2.Fields("PID").Value) = cint(AcquiredPlayerID2) or cint(objRSTeam2.Fields("PID").Value) = cint(AcquiredPlayerID3) then %>
							<tr style="background-color:#ddd;">
							<td class ="big" style="vertical-align:middle;text-align:center"><input type="checkbox" name="cmbTrade2" checked value="<%=objRSTeam2.Fields("PID").Value & ";" & objRSTeam2.Fields("firstName").Value & " " & objRSTeam2.Fields("lastName").Value%>"></td>
						<%else%>
							<tr style="background-color:white;">
							<td class ="big" style="vertical-align:middle;text-align:center"><input type="checkbox" name="cmbTrade2" value="<%=objRSTeam2.Fields("PID").Value & ";" & objRSTeam2.Fields("firstName").Value & " " & objRSTeam2.Fields("lastName").Value%>"></td>
						<%end if%>
							<td class ="big" style="vertical-align:middle;text-align:left">
						<%if (len(objRSTeam2.Fields("firstName").Value) + len(objRSTeam2.Fields("lastName").Value)) > 17 then %>
							<a class="blue" href="playerprofile.asp?pid=<%=objRSTeam2.Fields("PID").Value %>">
								<%=left(objRSTeam2.Fields("firstName").Value,1)%>.&nbsp;<%=left(objRSTeam2.Fields("lastName").Value,14)%>
							</a>
						<%else%>
							<a class="blue" href="playerprofile.asp?pid=<%=objRSTeam2.Fields("PID").Value %>">
								<%=objRSTeam2.Fields("firstName").Value%>&nbsp;<%=objRSTeam2.Fields("lastName").Value%>
							</a>
						<%end if%>
							<!--<a class="blue" href="playerprofile.asp?pid=<%=objRSTeam2.Fields("PID").Value %>" target="_self">
							<%=left(objRSTeam2.Fields("firstName").Value,1) + ". " + objRSTeam2.Fields("lastName").Value %>
							</a>-->
							<br><span class="greenTrade"><%=objRSTeam2.Fields("teamShortname").Value%>&nbsp;</span><span class="orange"><%=objRSTeam2.Fields("pos").Value%></small></span></td>
							<td style="vertical-align:middle;" class="text-center big">	
							<span class="badgeBlue"><%= round(objRSTeam2.Fields ("barps").Value,2) %></span>	
							</td>
							<% if CDbl(objRSTeam2.Fields("l5barps").Value) > CDbl(objRSTeam2.Fields("barps").Value) then %>
								<td class ="big" style="vertical-align:middle;text-align:center"><span class="badgeUp"><%= round(objRSTeam2.Fields ("l5barps").Value,2) %></span></td>
							<% elseif CDbl(objRSTeam2.Fields("barps").Value) > CDbl(objRSTeam2.Fields("l5barps").Value) then%>
								<td class ="big" style="vertical-align:middle;text-align:center"><span class="badgeDown"><%= round(objRSTeam2.Fields ("l5barps").Value,2) %></span></td>
							<%else%>
							<td class ="big" style="vertical-align:middle;text-align:center"><span class="badgeEven"><%= round(objRSTeam2.Fields ("l5barps").Value,2) %></span></td>
							<%end if %>	
							<% if objRSTeam2.Fields("ontheblock").Value = true then %>
								<td style="vertical-align: middle;text-align:center;"><span><i class="far fa-check"></i></span></td>
							<%else%>
								<td></td>
							<%end if%>
						</tr>
						<% objRSTeam2.MoveNext 
						Wend
						%>
						</tbody>
					</table>
				</div>
			</div>
			<div class="col-md-6">
				<div class="panel panel-override">
				<div class="panel-heading clearfix">
					<h5 class="panel-title"><%=objRSTeam1.Fields("TeamName").Value %></h5>
				</div>
				<table class="table table-custom-black table-responsive table-bordered table-condensed">
						<tr style="background-color:yellowgreen;">
							<td colspan="5">
								<table class="table table-custom table-responsive table-bordered table-condensed">
									<tr style="text-align:center";>
										<th style="color:black" colspan="8">Roster Configuration <small class="pull-right">Player Count: <%=myPlayerCnt%></small></th>
									</tr>
									<tr style="background-color:black;text-align:center;">
										<td class ="big" style="color:#ddd;font-weight:bold;width:12%">CEN</td>
										<td class ="big" style="background-color:white" width="12%"><%=centerCntMe%></td>
										<td class ="big" style="color:#ddd;font-weight:bold;width:12%">FOR</td>
										<td class ="big" style="background-color:white" width="12%"><%=forwardCntME%></td>
										<td class ="big" style="color:#ddd;font-weight:bold;width:12%">GUA</td>
										<td class ="big" style="background-color:white" width="12%"><%=guardCntME%></td>
										<td class ="big" style="color:#ddd;font-weight:bold;width:12%">TOT</td>
										</tr>
								</table>
							</td>
						</tr>
						<tr>
							<th style="border-radius:0px;text-align:center;white-space:nowrap;width:10%;"><i class="fas fa-basketball-ball"></i></th>
							<th style="border-radius:0px;text-align:left;width:40%;">Player - TM | Pos</th>
							<th style="border-radius:0px;text-align:center;width:20%;">AVG</th>
							<th style="border-radius:0px;text-align:center;width:20%;">L/5</th>
							<th style="border-radius:0px;text-align:center;width:10%;">OTB</th>
						</tr>
            <%  While Not objRSTeam1.EOF %>
						<% if cint(objRSTeam1.Fields("PID").Value) = cint(AcquiredPlayerID1)or cint(objRSTeam1.Fields("PID").Value) = cint(AcquiredPlayerID2) or cint(objRSTeam1.Fields("PID").Value) = cint(AcquiredPlayerID3) or cint(objRSTeam1.Fields("PID").Value) = cint(TradedPlayerID1) or cint(objRSTeam1.Fields("PID").Value) = cint(TradedPlayerID2) or cint(objRSTeam1.Fields("PID").Value) = cint(TradedPlayerID3) then %>
						<tr style="background-color:#ddd;">
							<td  class ="big" style="vertical-align:middle;text-align:center"><input type="checkbox" checked name="cmbTrade1" value="<%=objRSTeam1.Fields("PID").Value & ";" & objRSTeam1.Fields("firstName").Value & " " & objRSTeam1.Fields("lastName").Value%>"></td>
						<%else %>
						<tr style="background-color:white;">
							<td class ="big" style="vertical-align:middle;text-align:center"><input type="checkbox" name="cmbTrade1" value="<%=objRSTeam1.Fields("PID").Value & ";" & objRSTeam1.Fields("firstName").Value & " " & objRSTeam1.Fields("lastName").Value%>"></td>
						<%end if%>
							<td  class ="big" style="vertical-align:middle">
							<%if (len(objRSTeam1.Fields("firstName").Value) + len(objRSTeam1.Fields("lastName").Value)) > 17 then %>
								<a class="blue" href="playerprofile.asp?pid=<%=objRSTeam1.Fields("PID").Value %>">
									<%=left(objRSTeam1.Fields("firstName").Value,1)%>.&nbsp;<%=left(objRSTeam1.Fields("lastName").Value,14)%>
								</a>
							<%else%>
								<a class="blue" href="playerprofile.asp?pid=<%=objRSTeam1.Fields("PID").Value %>">
									<%=objRSTeam1.Fields("firstName").Value%>&nbsp;<%=objRSTeam1.Fields("lastName").Value%>
								</a>
							<%end if%>
							<!--<a class="blue" href="playerprofile.asp?pid=<%=objRSTeam1.Fields("PID").Value %>" target="_self">
							<%=left(objRSTeam1.Fields("firstName").Value,1) + ". " + objRSTeam1.Fields("lastName").Value %>
							</a>-->
							<br><span class="greenTrade"><%=objRSTeam1.Fields("teamShortname").Value%>&nbsp;</span><span class="orange"><%=objRSTeam1.Fields("pos").Value%></span>
							</td>
							<td style="vertical-align:middle;" class="text-center big">	
							<span class="badgeBlue"><%= round(objRSTeam1.Fields ("barps").Value,2) %></span>							
							</td>
							<% if CDbl(objRSTeam1.Fields("l5barps").Value) > CDbl(objRSTeam1.Fields("barps").Value) then %>
							<td class ="big" style="vertical-align:middle;text-align:center;"><span class="badgeUp"><%= round(objRSTeam1.Fields ("l5barps").Value,2) %></span></td>
							<% elseif CDbl(objRSTeam1.Fields("barps").Value) > CDbl(objRSTeam1.Fields("l5barps").Value) then%>
							<td class ="big" style="vertical-align:middle;text-align:center;"><span class="badgeDown"><%= round(objRSTeam1.Fields ("l5barps").Value,2) %></span></td>
							<%else%>
							<td class ="big" style="vertical-align:middle;text-align:center;"><span class="badgeEven"><%= round(objRSTeam1.Fields ("l5barps").Value,2) %></span></td>
							<%end if %>	
							<% if objRSTeam1.Fields("ontheblock").Value = true then %>
								<td style="vertical-align: middle;text-align:center;"><span><i class="far fa-check"></i></span></td>
							<%else%>
								<td></td>
							<%end if%>
					</tr>
            <% objRSTeam1.MoveNext 
            Wend
						%>
				</table> 
		
			</div>
    </div>
  </div>
</div>
<% if (wTradeDeadLine + 1 + 1/24) > now() then %>
   <div class="container">
	   <div class="row">
	   <div class="col-sm-12 col-md-12" align="right">	  
   		   <button type="submit" value="Submit Trade Offer" name="Action" class="btn btn-default  "><i class="fa fa-balance-scale" aria-hidden="true"></i>&nbsp;Analyze Trade</button>
		   <button type="reset"  value="Reset" name="Reset" class="btn btn-default  "><i class="fa fa-repeat" aria-hidden="true"></i>&nbsp;Reset</button>
		   </div>
	   </div>
   </div>
<%end if %>
<br>
</form>
<%
end if
If sAction = "Submit Trade Offer" and errorcode = True then %>
<!--#include virtual="Common/headerMain.inc"-->
<form action="tradeanalyzer.asp" method="POST" name="frmrejecttrade" language="JavaScript">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="cmbTeam" value="<%= sTeam %>" />
<input type="hidden" name="var_tradepartneremail" value="<%=objRSTeam2.Fields("HomeEmail").Value%>" />
<input type="hidden" name="var_owneremail" value="<%=objRSTeam1.Fields("HomeEmail").Value%>" />
<input type="hidden" name="var_teamname" value="<%=objRSTeam1.Fields("TeamName").Value%>" />
<input type="hidden" name="var_playercnt" value="<%=objRSTeam1.Fields("ActivePlayerCnt").Value%>" />
<input type="hidden" name="var_tradepartnername" value="<%=objRSTeam2.Fields("teamname").Value%>" />
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-danger">
			 <a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
				<strong>Error!</strong> Trade Request Error<br>
				Processing this trade would violate the roster limit of 14. Your previous trade has been discarded.			</div>
		</div>
	</div>
</div>
</form>
<%
end if
If sAction = "Submit Trade Offer" and errorcode = "Invalid Offer" then %>
<!--#include virtual="Common/headerMain.inc"-->
<form action="tradeanalyzer.asp" method="POST" name="frminvalidtrade" language="JavaScript">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="cmbTeam" value="<%= sTeam %>" />
<input type="hidden" name="var_tradepartneremail" value="<%=objRSTeam2.Fields("HomeEmail").Value%>" />
<input type="hidden" name="var_owneremail" value="<%=objRSTeam1.Fields("HomeEmail").Value%>" />
<input type="hidden" name="var_teamname" value="<%=objRSTeam1.Fields("TeamName").Value%>" />
<input type="hidden" name="var_playercnt" value="<%=objRSTeam1.Fields("ActivePlayerCnt").Value%>" />
<input type="hidden" name="var_tradepartnername" value="<%=objRSTeam2.Fields("teamname").Value%>" />
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-danger">
			 <a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
				<strong>Error!</strong> Trade Request Error<br>
				The following trade offer you specified cannot be processed due to an invalid browser state. Reasons this may happen are:<br>
				<ul type="bullet">
					<li>
					You hit the browser&#39;s back button 
					</li>
					<li>
					You are trying to offer an exact trade that is currently pending.
					</li>
					<li>
					You are trying to offer a trade involving players no longer on your team or no longer on your trading partner&#39;s team.
					</li>
				</ul>				
			</div>
		</div>
	</div>
</div>
</form>
<%
  end if
  objRSTeam1.Close
  objRSTeam2.Close
  objrstrade.Close
  ObjConn.Close
  Set objRSTeam1 = Nothing
  Set objRSTeam2 = Nothing
  Set objrstrade = Nothing
  Set objConn = Nothing

  Session.CodePage = Session("FP_OldCodePage")
  Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>
