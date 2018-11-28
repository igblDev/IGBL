
	<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""

	Dim objConn,objRS,ownerid,sAction,tradeid,objrstrade,objRSteam, objrsNames,txtNotes,towner
	
	'ENAIL VARIABLES
	Dim email_to, email_subject, host, username, password, reply_to, port, from_address
	Dim first_name, last_name, home_address, email_from, telephone, comments, error_message
	Dim ObjSendMail, email_message, objEmail


	GetAnyParameter "Action", sAction
	GetAnyParameter "var_tradeid", stradeid
	txtNotes = Request.Form("txtNotes")

	Set objConn   = Server.CreateObject("ADODB.Connection")
	Set objRS     = Server.CreateObject("ADODB.RecordSet")
	Set objrstrade= Server.CreateObject("ADODB.RecordSet")
	Set objrsNames= Server.CreateObject("ADODB.RecordSet")
	Set objRSteam = Server.CreateObject("ADODB.RecordSet")
	Set objEmail  = Server.CreateObject("ADODB.RecordSet")

	objConn.Open Application("lineupstest_ConnectionString")

	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"

	%>
	<!--#include virtual="Common/session.inc"-->
	<%
	tradeid   = stradeid

	objrstrade.Open	"SELECT t.*, to1.HomeEmail As TraderEmail, to1.TeamName AS TraderName, " & _
									"to2.HomeEmail, to2.TeamName, to2.ShortName " & _
									"FROM tblpendingtrades t, tblowners to1, tblowners to2 " & _
									"WHERE t.tradeid = " & tradeid & "  "  & _
									"and  t.tooid = to1.ownerid " & _
									"and  t.Fromoid = to2.ownerid ", objConn,3,3,1

	w_trade_count = objrstrade.Recordcount

	tplayer1   = objrstrade.Fields("TradedPlayerID").Value
	tplayer2   = objrstrade.Fields("TradedPlayerID2").Value
	tplayer3   = objrstrade.Fields("TradedPlayerID3").Value
	aplayer1   = objrstrade.Fields("AcquiredPlayerID").Value
	aplayer2   = objrstrade.Fields("AcquiredPlayerID2").Value
	aplayer3   = objrstrade.Fields("AcquiredPlayerID3").Value
	traderEmail= objrstrade.Fields("TraderEmail").Value
	traderName = objrstrade.Fields("TraderName").Value
	myname     = objrstrade.Fields("TeamName").Value
	myShortnme = objrstrade.Fields("ShortName").Value
	myemail    = objrstrade.Fields("HomeEmail").Value
	towner     = objrstrade.Fields("toOID").Value
	tcoments   = objrstrade.Fields("tradeComments").Value
	descDate   = objrstrade.Fields("DecisionDate").Value
	origDate   = objrstrade.Fields("TransDate").Value	
	
	objrstrade.close

	select case sAction
	  case "Withdraw Trade Offer"


		if w_trade_count > 0 then

			FuncCall = Get_Player_Names(tradeid, msgtradeplayers, msgacquireplayers)

			'*************************************************************
			'DELETION OF TRADE RECORD
			'**************************************************************
			strSQL = "DELETE FROM tblpendingtrades WHERE TradeID =   "& tradeid & ";"
			objConn.Execute strSQL

			FuncCall = Reset_PendTrade_Flags
	
			wEmailOwnerID  = towner
			wAlert         = "receiveTradeAlerts"
			email_subject  = "Trade Withdrawn by " & myShortnme
			email_message  = msgtradeplayers & "<br>"
			email_message  = email_message & "for <br><br>"
			email_message  = email_message & msgacquireplayers
			if len(txtNotes) > 0 then
			   email_message = email_message&"<br>"&txtNotes
			end if
%>		
		<!--#include virtual="Common/email_league.inc"-->				   
<%			
		
			dim TradeCnt,objRSTrades
			Set objRSTrades  = Server.CreateObject("ADODB.RecordSet")	
			
			objRSTrades.Open "SELECT * FROM qryPendingTrade WHERE (((qryPendingTrade.towner)=" & ownerid & ")) order by DecisionDate", objConn,3,3,1
			TradeCnt   = objRSTrades.RecordCount
			Response.Write "Record Count = : " & analysisCnt  & "  <br>"
			
			if TradeCnt > 0 then
				sURL = "pendingTrades.asp"
			else
				sURL = "dashboard.asp"
			end if	
			
			AddLinkParameter "var_ownerid", ownerid, sURL
			Response.Redirect sURL

		else
			errorcode = "Trade Deleted"
		end if

		case ""
		ownerid = session("ownerid")
		objRS.Open "SELECT * FROM qryPendingTrade WHERE (((qryPendingTrade.towner)=" & ownerid & ")) order by DecisionDate", objConn,3,3,1
		if objRS.RecordCount = 0 then
			errorcode = "Trade Deleted"	
			sAction = "Withdraw Trade Offer"
			errormessage = "Trade No longer exists in the database."
		end if
			
		case "Return"
		ownerid = session("ownerid")
    case "My-IGBL"
		sURL = "lineups.asp"
		AddLinkParameter "var_ownerid", ownerid, sURL
		Response.Redirect sURL

    case "Delete Invalid Trade"
		Response.Write "Trade ID  = : " & stradeid  & "  <br>"
    strSQL = "DELETE FROM tblpendingtrades where tradeid =" & stradeid & " ;"
		objConn.Execute strSQL

		sURL = "lineups.asp"
		AddLinkParameter "var_ownerid", ownerid, sURL
		Response.Redirect sURL

	end select

	'***********************************************************
	' Upd_PendingTrade_FLAG()
	' This function will search the pending Trades table to see if
	' the Player is involved in any more trades.  IF there are no more
	' trades involving this player then the Pending Trade flag is set to
	' False.
	'***********************************************************
	Function Upd_PendingTrade_Flag (p_PlayerID)
		Response.Write "PlayerID = " & p_PlayerID & " <br> "

		objrstrade.Open	"SELECT * " & _
        "FROM tblpendingtrades " & _
				"WHERE  tradedplayerid      = "& p_PlayerID & "   " & _
								"or tradedplayerid2 = "& p_PlayerID & "   " & _
								"or tradedplayerid3 = "& p_PlayerID & "   " & _
								"or acquiredplayerid = "& p_PlayerID & "  " & _
								"or acquiredplayerid2 = "& p_PlayerID& " " & _
								"or acquiredplayerid3 = "& p_PlayerID & " ", objConn,3,3,1

		if objrstrade.RecordCount = 0 then
			strSQL = "UPDATE tblPlayers SET tblPlayers.PendingTrade = 0 WHERE PID=" & p_PlayerID & ";"
			objConn.Execute strSQL
		end if

		objrstrade.close

	End Function

	'***********************************************************
	' Get_Player_Names()
	'***********************************************************
   	Function Get_Player_Names(p_tradeID, msgtradeplayers, msgacquireplayers)

		objrsNames.Open "SELECT * FROM qryPendingTrade Where TradeID = "&p_tradeID&" " , objConn

		'********************************************************
		'** Populate Traded Players String
		'********************************************************

		msgtradeplayers = objrsNames.Fields("t1first").Value & " " & objrsNames.Fields("t1last").Value & "<br />"

		if objrsNames.Fields("t2PID").Value > 0 then
			msgtradeplayers = msgtradeplayers & objrsNames.Fields("t2first").Value & " " & objrsNames.Fields("t2last").Value & "<br />"  
		end if

		if objrsNames.Fields("t3PID").Value > 0 then
			msgtradeplayers = msgtradeplayers & objrsNames.Fields("t3first").Value & " " & objrsNames.Fields("t3last").Value & "<br />"
		end if

		'********************************************************
		'** Populate Acquired Players String
		'********************************************************

		msgacquireplayers= objrsNames.Fields("a1first").Value & " " & objrsNames.Fields("a1last").Value & "<br />"

		if objrsNames.Fields("a2PID").Value > 0 then
			msgacquireplayers = msgacquireplayers & objrsNames.Fields("a2first").Value & " " & objrsNames.Fields("a2last").Value & "<br />"
		end if

		if objrsNames.Fields("a3PID").Value > 0 then
			msgacquireplayers = msgacquireplayers &  objrsNames.Fields("a3first").Value & " " & objrsNames.Fields("a3last").Value & "<br />"
		end if

		objrsNames.Close
	End Function

	t1first   = objRS.Fields("t1first").Value
	t1last    = objRS.Fields("t1last").Value


	if objRS.Fields("t2PID").Value > 0 then
		t2first = objRS.Fields("t2first").Value
		t2last  = objRS.Fields("t2last").Value
	end if

	if objRS.Fields("t3PID").Value > 0 then
		t3first = objRS.Fields("t3first").Value
			t3last= objRS.Fields("t3last").Value
	end if

	a1first   = objRS.Fields("a1first").Value
	a1last    = objRS.Fields("a1last").Value

	if objRS.Fields("a2PID").Value > 0 then
		a2first = objRS.Fields("a2first").Value
		a2last  = objRS.Fields("a2last").Value
	end if

	if objRS.Fields("a3PID").Value > 0 then
		a3first = objRS.Fields("a3first").Value
		a3last  = objRS.Fields("a3last").Value
	end if

%>
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
red {
	color:red;
}
black {
	color:black;
}
white {
	color:white;
}
yellow {
	color:yellow;
}
green {
	color:#468847;
	font-weight: bold;
	font-size:10px;	
  text-transform: uppercase;
}

blackMsg {
	color:black;
	font-weight: bold;
	font-size:10px;	
  text-transform: uppercase;
}

.panel-title {
    color: #FFEB3B;
    text-transform: none;
    font-size: 14px !important;
}
.panel-heading {
    background-image: none;
    background-color: #01579B !important;
    color: white;
    height: 30px;
    padding: 5px 5px;
		border-radius: unset;
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
</style>
</head>
<body>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Pending Trade Offers</strong>
			</div>
		</div>
	</div>
</div>
<% if sAction = "" then %>
<%
 While Not objRS.EOF
	wTeamName = replace((objRS.Fields("ateamname").Value), "THE ", "")
	if len(wTeamName) >19 then 
	 wTeamName = objRS.Fields("tteamnameshort").Value
	end if
%>
<form action="pendingtrades.asp" name="frmMain" method="POST">
  <input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
  <input type="hidden" name="var_tradeid" value="<%=objRS.Fields("tradeid").Value %>" />
  <input type="hidden" name="var_invtradeind" value="<%=objRS.Fields("InvalidTradeInd").Value %>" />
  <!--#include virtual="Common/headerMain.inc"-->
  <div class="container">
    <div class="row">
      <div class="col-md-12 col-sm-12 col-xs-12">	
					<div class="panel panel-override">
            <table class="table table-custom-black table-bordered table-condensed">
							<tr>
								<th style="width:50%;">My Players</th>
								<th style="width:50%;"><%=wTeamName%></th>
							</tr>
							<%
							objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&t1first&"' and lastName = '"&t1last&"' "
							set playerBarps = 0	 
							if objrstrade.RecordCount >= 0 then
								playerBarps = objrstrade.Fields("barps").Value
							end if
							%>
							<tr>
								<td style="width:50%;">
									<table class="table table-bordered table-condensed">
										<tr style="background-color:white">
											<td>
											<%if (len(objRS.Fields("t1first").Value) + len(objRS.Fields("t1last").Value)) >= 14 then %>
												<a class="blue big" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>">
													<%=left(objRS.Fields("t1first").Value,1)%>.&nbsp;<%=left(objRS.Fields("t1last").Value,14)%></a>
											<%else%>
												<a class="blue big" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>">
													<%=left(objRS.Fields("t1first").Value,8)%>&nbsp;<%=left(objRS.Fields("t1last").Value,10)%></a>
											<%end if%>
												<small><span class="greenTrade big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span></small>
											</td>
										</tr>
										<%
										objrstrade.Close
										if objRS.Fields("t2PID").Value > 0 then
											objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&t2first&"' and lastName = '"&t2last&"' "
											set playerBarps = 0	 
											if objrstrade.RecordCount >= 0 then
												playerBarps = objrstrade.Fields("barps").Value
											end if
										%>
										<tr style="background-color:white">
											<td>
											<%if (len(objRS.Fields("t2first").Value) + len(objRS.Fields("t2last").Value)) >= 14 then %>
												<a class="blue big" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("t2first").Value,1)%>.&nbsp;<%=left(objRS.Fields("t2last").Value,14)%></a>
											<%else%>
												<a class="blue big" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("t2first").Value,8)%>&nbsp;<%=left(objRS.Fields("t2last").Value,10)%></a>
											<%end if%>
												<small><span class="greenTrade big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span></small>
											</td>
										</tr>
										<%end if%>
										<%
										objrstrade.Close
										if (objRS.Fields("t3PID").Value) > 0 then
											objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&t3first&"' and lastName = '"&t3last&"' "
											set playerBarps = 0	 
											if objrstrade.RecordCount >= 0 then
											 playerBarps = objrstrade.Fields("barps").Value
											end if
										%>
										<tr style="background-color:white">
											<td>
											<%if (len(objRS.Fields("t3first").Value) + len(objRS.Fields("t3last").Value)) >= 14 then %>						
												<a class="blue big" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("t3first").Value,1)%>.&nbsp;<%=left(objRS.Fields("t3last").Value,14)%></a>
											<%else%>
												<a class="blue big" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("t3first").Value,8)%>&nbsp;<%=left(objRS.Fields("t3last").Value,10)%></a>
											<%end if%>
												<small><span class="greenTrade big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span></small>
											</td>
										</tr>
										<% end if %>
									</table>
								</td>
								<%
								objrstrade.Close
								objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&a1first&"' and lastName = '"&a1last&"' " 
								set playerBarps = 0	 
								 if objrstrade.RecordCount >= 0 then
									playerBarps = objrstrade.Fields("barps").Value
								 end if
								%>
								<td style="width:50%;">
										<table class="table table-bordered table-condensed">
										<tr style="background-color:white">
											<td>
											<%if (len(objRS.Fields("a1first").Value) + len(objRS.Fields("a1last").Value)) >= 14 then %>
												<a class="blue big" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("a1first").Value,1)%>.&nbsp;<%=left(objRS.Fields("a1last").Value,14)%></a>
											<%else%>
												<a class="blue big" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("a1first").Value,8)%>&nbsp;<%=left(objRS.Fields("a1last").Value,10)%></a>
											<%end if%>
												<small><span class="greenTrade big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span></small>
											</td>
										</tr>
									<%
									objrstrade.Close
									if (objRS.Fields("a2PID").Value) > 0 then
										objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&a2first&"' and lastName = '"&a2last&"' " 
										set playerBarps = 0	 
										if objrstrade.RecordCount >= 0 then
											playerBarps = objrstrade.Fields("barps").Value
										end if
									%>
									<tr style="background-color:white">
										<td>
										<%if (len(objRS.Fields("a2first").Value) + len(objRS.Fields("a2last").Value)) >= 14 then %>
											<a class="blue big" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("a2first").Value,1)%>.&nbsp;<%=left(objRS.Fields("a2last").Value,14)%></a>
										<%else%>
											<a class="blue big" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("a2first").Value,8)%>&nbsp;<%=left(objRS.Fields("a2last").Value,10)%></a>
										<%end if%>
											<small><span class="greenTrade big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span></small>
										</td>
									</tr>
									<% end if %>
									<% 
									objrstrade.Close
									if (objRS.Fields("a3PID").Value) > 0 then
										objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&a3first&"' and lastName = '"&a3last&"' "  		
										set playerBarps = 0	 
										if objrstrade.RecordCount >= 0 then
											playerBarps = objrstrade.Fields("barps").Value
										end if
									 %>
									<tr style="background-color:white">
										<td>
										<%if (len(objRS.Fields("a3first").Value) + len(objRS.Fields("a3last").Value)) >= 14 then %>
											<a class="blue big" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("a3first").Value,1)%>.&nbsp;<%=left(objRS.Fields("a3last").Value,14)%></a>
										<%else%>
											<a class="blue big" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("a3first").Value,8)%>&nbsp;<%=left(objRS.Fields("a3last").Value,10)%></a>
										<%end if%>
											<small><span class="greenTrade big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span></small>
										</td>
									</tr>
									<% end if %>									
									</table>
								</td>
							</tr>
              <%
							objrstrade.Close 
							%>
						<tr style="background-color:black;font-weight:bold;color:yellowgreen;">
							<td style="text-align:center;" colspan="6">This Trade offer expires around <%=objRS.Fields("DecisionDate").Value%></td>
						</tr>
							<tr>
								<td  colspan="6" valign="top"><textarea name="txtNotes" class="form-control" rows="3" value="<%=tcoments%>" placeholder="Enter Withdraw Trade Comments" id="txtNotes"></textarea></td>
							</tr>
 							<tr>
								<td colspan="6">
									<div>
										<button type="submit" value="Withdraw Trade Offer" name="Action" class="btn btn-default-red btn-block  btn-sm"><i class="fas fa-trash-alt"></i>&nbsp;Withdraw Offer</button>            </A> 
									</div>
								</td>
							</tr>
				</table>
			</div>
    </div>
  </div>
</div>	
</form>
<%
 	objRS.MoveNext

 	t1first  = objRS.Fields("t1first").Value
	t1last   = objRS.Fields("t1last").Value

	if objRS.Fields("t2PID").Value > 0 then
		t2first   = objRS.Fields("t2first").Value
		t2last    = objRS.Fields("t2last").Value
	end if

	if objRS.Fields("t3PID").Value > 0 then
		t3first  = objRS.Fields("t3first").Value
  		t3last   = objRS.Fields("t3last").Value
	end if

	a1first  = objRS.Fields("a1first").Value
	a1last   = objRS.Fields("a1last").Value

	if objRS.Fields("a2PID").Value > 0 then
		a2first   = objRS.Fields("a2first").Value
		a2last   = objRS.Fields("a2last").Value
	end if

	if objRS.Fields("a3PID").Value > 0 then
		a3first  = objRS.Fields("a3first").Value
		a3last   = objRS.Fields("a3last").Value
	end if
	Wend

%>
<%end if %>
<%if sAction = "Withdraw Trade Offer" and errorcode = "Trade Deleted"  then %>
<!--#include virtual="Common/headerMain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-danger">
				<strong><i class="fa fa-exclamation-triangle fa-lg" aria-hidden="true"></i></strong> Pending Trade Error!<br>
				<%=errormessage %>			
			</div>
		</div>
	</div>
</div>
<%end if %>
<%
  objConn.Close
  Set objConn = Nothing
%>
</body>
</html>