<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

 	Dim objRSgames,objRS,objConn, objRSwaivers, ownerid,objRSplayers, objsrCycle, objsrBonus
	Dim strSQL, iPlayerClaimed,objRSTxns, objRSOwners, objRejectWaivers, iPlayerWaived, iOwner, w_action
	
	Set objConn       = Server.CreateObject("ADODB.Connection")
	Set objRSgames    = Server.CreateObject("ADODB.RecordSet")
	Set objRSplayers  = Server.CreateObject("ADODB.RecordSet")
	Set objsrCycle    = Server.CreateObject("ADODB.RecordSet")
	Set objsrBonus    = Server.CreateObject("ADODB.RecordSet")
	Set objRSCenters  = Server.CreateObject("ADODB.RecordSet")
	Set objRSForwards = Server.CreateObject("ADODB.RecordSet")
	Set objRSGuards   = Server.CreateObject("ADODB.RecordSet") 
	Set objRSRosterCnt= Server.CreateObject("ADODB.RecordSet") 
	Set objRSteams= Server.CreateObject("ADODB.RecordSet")
	
	objConn.Open Application("lineupstest_ConnectionString")
	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"

	teamloop = 10							
%>

	<!--#include virtual="Common/SESSION.inc"-->
<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!-- Bootstrap core CSS -->
<!--#include virtual="Common/bootstrap.inc"-->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<!--#include virtual="Common/headerMain.inc"-->
<style>
gray {
	
}
red {
	color: #9a1400;
}
black {
	color: black;
}
white {
	color: white;
}
green {
	color: #468847;
		
}
.alert-warning {
    color: #8a6d3b;
    background-color: #fcf8e3;
    border-color: #354478;
}
.panel-override {
    color: black;
    background-color:WHITE;
    border-color: black;
    border-radius: 0px;
}
panel-title {
    color: yellowgreen;
    text-transform: none;
    font-size: 20px !important;
    font-weight: 700;
}
tr.green td {
  background-color: #468847;
  color: black;
}
tr.d1 td {
  background-color: #9999CC;
  color: black;
}

</style>
</head>
<body>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong><i class="fas fa-users"></i>&nbsp;TEAM ANALYSIS BREAKDOWN</strong>
			</div>
		</div>		
	</div>
</div>
<div class="container">
  <div class="row">
    <div class="col-md-12 col-sm-12 col-xs-12">     
			<table class="table table-responsive table-bordered  table-condensed">
				<%
				While teamloop > 0
				objRSteams.Open "SELECT * FROM tblOwners ", objConn,3,3,1
				%>
				<tr style="background-color:white;">
					<td style="width:50%;text-align:left;"><a class="blue" href="#<%=objRSteams.Fields("teamName").Value%>"><%=objRSteams.Fields("teamName").Value%></a></td>
					<%
					objRSteams.MoveNext
					teamloop  = teamloop - 1
					teamcnt = teamcnt + 1
					%>
					<td style="width:50%;text-align:left;"><a class="blue" href="#<%=objRSteams.Fields("teamName").Value%>"><%=objRSteams.Fields("teamName").Value%></a></td>
					<%
					objRSteams.MoveNext
					teamloop  = teamloop - 1
					teamcnt = teamcnt + 1
					Wend
					objRSteams.CLose
					%>
				</tr>
			</table>
    </div>
  </div>
</div>
</br>
<div class="container">
	<div class="row">
		<div class="col-xs-12">
				<div style="text-align: center;font-size:13px;"><strong><mark>Barp Legend:</mark></strong>&nbsp;&nbsp;
				<span style="color:darkorange;"><i class="fas fa-basketball-ball fa-lg"></i></span>&nbsp;<strong>45+</strong>&nbsp;
				<i class="fab fa-superpowers big fa-lg" style="color:#468847;font-weight:bold;font-size:13px;"></i>&nbsp;<strong>35+</strong>&nbsp;
				<i class="fas fa-battery-bolt fa-lg " style="color:purple;font-weight:bold;font-size:13px;"></i>&nbsp;<strong>25+</strong>&nbsp;
				<i class="fas fa-poo fa-lg" style="color:#6d4c41;font-weight:bold;font-size:13px;"></i>&nbsp;<strong>24-</strong>
				</div>
			</div>
	</div>
</div>
</br>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div>		

				<%
					objRSgames.Open   "SELECT * FROM tblowners o, standings s where o.ownerId = s.id order by s.rank asc ", objConn,3,3,1
					objsrCycle.Open   "SELECT max (cycle) as currentCycle from Standings_Cycle", objConn,3,3,1
					w_current_cycle = objsrCycle.Fields("currentCycle").Value
					objsrCycle.Close
				%>
				<%
					While Not objRSgames.EOF
					processingOwner = objRSgames.Fields("ownerId").Value
					
					objsrBonus.open  "SELECT * From Standings_Cycle WHERE cycle = "& w_current_cycle &" and ID = "& objRSgames.Fields("ownerId").Value &" ",objConn,3,3,1 
					cycleWins = objsrBonus.Fields("won").Value
					cycleLoss = objsrBonus.Fields("loss").Value					
					objsrBonus.Close
					
					objRSCenters.Open   "SELECT * FROM tblPlayers WHERE (tblPlayers.POS  = 'CEN' or tblPlayers.POS  = 'F-C') and tblPlayers.OwnerID=" & processingOwner & "",objConn,3,3,1
					objRSForwards.Open  "SELECT * FROM tblPlayers WHERE (tblPlayers.POS  = 'FOR' or tblPlayers.POS  = 'F-C' or tblPlayers.POS   = 'G-F') and tblPlayers.OwnerID=" & processingOwner & "",objConn,3,3,1
					objRSGuards.Open  	"SELECT * FROM tblPlayers WHERE (tblPlayers.POS  = 'GUA' or tblPlayers.POS  = 'G-F') and tblPlayers.OwnerID=" & processingOwner & "",objConn,3,3,1
					objRSRosterCnt.Open  "SELECT * FROM tblPlayers WHERE  tblPlayers.OwnerID=" & processingOwner & "",objConn,3,3,1
					
					centerCntMe  = objRSCenters.RecordCount
					forwardCntMe = objRSForwards.RecordCount
					guardCntME   = objRSGuards.RecordCount
					rosterCnt    = objRSRosterCnt.RecordCount

					objRSCenters.Close  	
					objRSForwards.Close  
					objRSGuards.Close 
					objRSRosterCnt.Close
					
				%>
		<div id="<%= objRSgames.Fields("teamName").Value%>"></div>
		<table class="table table-custom-black table-responsive table-bordered table-condensed">
			<tr>
				<th style="text-align:center;">Team Analysis</th>
			</tr>
			<tr>
				<td class="big" style="text-align: center;background-color: black;color: white;font-size:14px;font-weight">
					<span style="font-weight:bold;font-size:20px;"><%=objRSgames.Fields("teamname").Value %></span>&nbsp;<span style="color:darkorange;font-size:16px;font-weight:bold;">[<%=objRSgames.Fields("shortName").Value %>]</span>
					</br><span style="color:white;">REG:</span> <span style="color:gold;font-weight:bold;">(<%=objRSgames.Fields("won").Value %>-<%=objRSgames.Fields("loss").Value %>)</span>&nbsp;<span style="color:white;">CYCLE:</span> <span style="color:gold;font-weight:bold;">(<%=cycleWins%>-<%=cycleLoss %>)</span>
					</br> RANK:&nbsp;<span style="color: yellowgreen;font-weight:bold;"><%=objRSgames.Fields("rank").Value %></span>
					</br> WAIVER BALANCE:&nbsp;<span style="color: springgreen;font-weight:bold;"><%= FormatCurrency(objRSgames.Fields("WaiverBal").Value)%></span>
					</br>PPG:&nbsp;<span style="color: #2196F3;font-weight:bold;"><%=objRSgames.Fields("ppg").Value %></span>&nbsp;|&nbsp;OPPG:&nbsp;<span style="color: darkorange;font-weight:bold;"><%=objRSgames.Fields("oppg").Value %></span>&nbsp;|&nbsp;PPG Diff:
					<%if CInt(objRSgames.Fields("ppg").Value) >= CInt(objRSgames.Fields("oppg").Value) then %>
						<span style="color:cyan;font-weight:bold;">+<%=objRSgames.Fields("diff").Value %></span>
					<%else%>
						<span style="color:red;font-weight:bold;"><%=objRSgames.Fields("diff").Value %></span>
					<%end if%>
					<!--</br>Roster Player Count:&nbsp;<span style="color: yellow;font-weight:bold;"><%= rosterCnt%></span>-->
				</td>
			</tr>
			<!--<tr style="text-align:center";>
				<th style="color:black" colspan="6">Roster Configuration</th>
			</tr>-->
			<!--<tr style="text-align:center;">
				<td colspan="6">
					<table class="table table-striped table-bordered table-condensed table-responsive">
						<tr class="big">
							<td class ="big" style="background: yellowgreen;color:black;font-weight:bold;width:18%">CEN</td>
							<td class ="big" style="background-color:white" width="15%"><%=centerCntMe%></td>
							<td class ="big" style="background: yellowgreen;color:black;font-weight:bold;width:18%">FOR</td>
							<td class ="big" style="background-color:white" width="15%"><%=forwardCntME%></td>
							<td class ="big" style="background: yellowgreen;color:black;font-weight:bold;width:18%">GUA</td>
							<td class ="big" style="background-color:white" width="15%"><%=guardCntME%></td>
							</td>
						</tr>
					</table>
				</td>
			</tr>-->
			<tr>
			<%
			objRSplayers.Open "SELECT * FROM qry_PlayerAll where ownerId = " &objRSgames.Fields("ownerId").Value& " " & _
			                  "AND barps >= 35 order by barps desc ", objConn,3,3,1
			%>
			<td class="big" style="background-color:white;"><i class="fab fa-superpowers fa-lg big"></i> <span style="font-weight:bold;color:#468847;text-decoration:underline;font-size:13px;">STRENGTHS</span><br>
				<table class="table table-responsive table-striped table-bordered table-condensed">
					<tr>
						<th class="big" style="width:48%;">Player</th>	
						<th class="big" style="width:8%;text-align:center;">OTB</th>
						<th class="big" style="width:18%;text-align:center;">AVG</th>
						<th class="big" style="width:8%;text-align:center;">T</th>
						<th class="big" style="width:18%;text-align:center;">L/5</th>
					</tr>
					<%
						While Not objRSplayers.EOF
					%>
					<tr style="color:#468847;">
						<td class="big" style="vertical-align: middle;"><a class="blue" href="playerprofile.asp?pid=<%=objRSplayers.Fields("PID").Value %>"><%=left(objRSplayers.Fields("firstName").Value,1)%>.&nbsp;<%=left(objRSplayers.Fields("lastName").Value,14)%></a>&nbsp;<span class="greenTrade"><%=objRSplayers.Fields("TeamshortName").Value%></span>&nbsp;<span class="orange"><%=objRSplayers.Fields("POS").Value%></span>
						<% if objRSplayers.Fields("rentalPlayer").Value = true then %>
							 <strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong>
						<%end if%>
						</td>	
						<% if objRSplayers.Fields("ontheblock").Value = true then %>
							<td class="big" style="vertical-align: middle;text-align:center;"><span><i class="far fa-check"></i></span></td>
						<%else%>
							<td></td>
						<%end if%>					
						<%if round(objRSplayers.Fields("barps").Value,2) >= 45 then%>
							<td class="big" style="vertical-align: middle;text-align:center;"><span style="color:darkorange;"><i class="fas fa-basketball-ball"></i></span>&nbsp;<%=round(objRSplayers.Fields("barps").Value,2) %></span></td>
						<%else%>
							<td class="big"style="vertical-align: middle;text-align:center;"><%=round(objRSplayers.Fields("barps").Value,2) %></td>		
						<%end if%>

						<% if CDbl(objRSplayers.Fields("l5barps").Value) > CDbl(objRSplayers.Fields("barps").Value) then %>
							<td class="big" style="vertical-align:middle;text-align:center;"><i class="fal fa-long-arrow-up" style="color:#468847;vertical-align:middle;text-align:center;font-weight:bold;"></i></td>
						<% elseif CDbl(objRSplayers.Fields("barps").Value) > CDbl(objRSplayers.Fields("l5barps").Value) then%>
							<td class="big" style="vertical-align:middle;text-align:center;"><i class="fal fa-long-arrow-down" style="color:#9a1400;vertical-align:middle;text-align:center;font-weight:bold;"></i></td>
						<%else%>
							<td class="big" style="vertical-align:middle;text-align:center;"><i class="fal fa-arrows-h" style="color:gold;vertical-align:middle;text-align:center;font-weight:bold;"></i></td>
						<%end if %>	

						<%if round(objRSplayers.Fields("l5barps").Value,2) >= 45 then%>
							<td class="big" style="vertical-align: middle;text-align:center;"><span style="color:darkorange;"><i class="fas fa-basketball-ball"></i></span>&nbsp;<%=round(objRSplayers.Fields("l5barps").Value,2) %></td>
						<%else%>
							<td class="big" style="vertical-align: middle;text-align:center;"><%=round(objRSplayers.Fields("l5barps").Value,2) %></td>		
						<%end if%>										
					</tr>
					<%
						objRSplayers.MoveNext
						Wend
					%>
				</table>
			</td>
          </tr>
			<tr>
				<%
				objRSplayers.Close
				objRSplayers.Open "SELECT * FROM qry_PlayerAll where ownerId = " &objRSgames.Fields("ownerId").Value& " " &_
			                      "AND barps >24 and barps < 35 order by barps desc ", objConn,3,3,1
				%>
				<td class="big" style="background-color:white;"><i class="fas fa-battery-bolt fa-lg"></i> <span style="font-weight:bold;color:purple; text-decoration:underline;font-size:13px;">SUPPORTING CAST</span><br>
					<table class="table table-responsive table-striped table-bordered table-condensed">
					<tr>
						<th class="big" style="width:48%;">Player</th>	
						<th class="big" style="width:8%;text-align:center;">OTB</th>
						<th class="big" style="width:18%;text-align:center;">AVG</th>
						<th class="big" style="width:8%;text-align:center;">T</th>
						<th class="big" style="width:18%;text-align:center;">L/5</th>
					</tr>
						<%
							While Not objRSplayers.EOF
						%>
						<tr style="color:#5b33b7">
							<td class="big"><a class="blue" href="playerprofile.asp?pid=<%=objRSplayers.Fields("PID").Value %>"><%=left(objRSplayers.Fields("firstName").Value,1)%>.&nbsp;<%=left(objRSplayers.Fields("lastName").Value,14)%></a>&nbsp;<span class="greenTrade"><%=objRSplayers.Fields("TeamshortName").Value%></span>&nbsp;<span class="orange"><%=objRSplayers.Fields("POS").Value%></span>
							<% if objRSplayers.Fields("rentalPlayer").Value = true then %>
							 <strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong>
							<%end if%>
							</td>	
							<% if objRSplayers.Fields("ontheblock").Value = true then %>
								<td style="vertical-align: middle;text-align:center;"><span><i class="far fa-check"></i></span></td>
							<%else%>
								<td></td>
							<%end if%>	
							<td class="big" style="text-align:center;color:purple;"><%=round(objRSplayers.Fields("barps").Value,2) %></td>
							
							<% if CDbl(objRSplayers.Fields("l5barps").Value) > CDbl(objRSplayers.Fields("barps").Value) then %>
							<td style="vertical-align:middle;text-align:center;"><i class="fal fa-long-arrow-up" style="color:#468847;vertical-align:middle;text-align:center;font-weight:bold;"></i></td>
							<% elseif CDbl(objRSplayers.Fields("barps").Value) > CDbl(objRSplayers.Fields("l5barps").Value) then%>
							<td style="vertical-align:middle;text-align:center;"><i class="fal fa-long-arrow-down" style="color:#9a1400;vertical-align:middle;text-align:center;font-weight:bold;"></i></td>
							<%else%>
							<td style="vertical-align:middle;text-align:center;"><i class="fal fa-arrows-h" style="color:gold;vertical-align:middle;text-align:center;font-weight:bold;"></i></td>
							<%end if %>	
							
							<%if round(objRSplayers.Fields("l5barps").Value,2) >= 45 then%>
							<td class="big" style="text-align:center;color:#9a1400;"><span style="color:darkorange;"><i class="fas fa-basketball-ball"></i></span>&nbsp;<%=round(objRSplayers.Fields("l5barps").Value,2) %></td>
							<%else%>
								<td class="big" style="text-align:center;color:purple"><%=round(objRSplayers.Fields("l5barps").Value,2) %></td>		
							<%end if%>		
						</tr>
						<%
							objRSplayers.MoveNext
							Wend
						%>
					</table>
				</td>
			</tr>
			<%
			objRSplayers.Close
			objRSplayers.Open "SELECT * FROM qry_PlayerAll where ownerId = " &objRSgames.Fields("ownerId").Value& " " &_
			                  "AND barps <= 24 order by barps desc ", objConn,3,3,1
			%>
			<% if objRSplayers.RecordCount > 0 then %>
			<tr>
				<td class="big" style="background-color:white;"><i class="fas fa-poo fa-lg"></i> <span style="font-weight:bold;color:#6d4c41;text-decoration: underline;font-size:13px;">WEAKNESSES</span><br>
					<table class="table table-responsive table-striped table-bordered table-condensed">
					<tr>
						<th class="big" style="width:48%;">Player</th>	
						<th class="big" style="width:8%;text-align:center;">OTB</th>
						<th class="big" style="width:18%;text-align:center;">AVG</th>
						<th class="big" style="width:8%;text-align:center;">T</th>
						<th class="big" style="width:18%;text-align:center;">L/5</th>
					</tr>
						<%
							While Not objRSplayers.EOF
						%>
						<tr style="color:#6d4c41;">
							<td class="big"><a class="blue" href="playerprofile.asp?pid=<%=objRSplayers.Fields("PID").Value %>"><%=left(objRSplayers.Fields("firstName").Value,1)%>.&nbsp;<%=left(objRSplayers.Fields("lastName").Value,14)%></a>&nbsp;<span class="greenTrade"><%=objRSplayers.Fields("TeamshortName").Value%></span>&nbsp;<span class="orange"><%=objRSplayers.Fields("POS").Value%></span>
							<% if objRSplayers.Fields("rentalPlayer").Value = true then %>
							 <strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong>
							<%end if%>
							</td>	
							<% if objRSplayers.Fields("ontheblock").Value = true then %>
								<td style="vertical-align: middle;text-align:center;"><span><i class="far fa-check"></i></span></td>
							<%else%>
								<td></td>
							<%end if%>							
							<td class="big" style="text-align:center;color:#6d4c41;"><%=round(objRSplayers.Fields("barps").Value,2) %></td>							
							
							<% if CDbl(objRSplayers.Fields("l5barps").Value) > CDbl(objRSplayers.Fields("barps").Value) then %>
							<td style="vertical-align:middle;text-align:center;"><i class="fal fa-long-arrow-up" style="color:#468847;vertical-align:middle;text-align:center;font-weight:bold;"></i></td>
							<% elseif CDbl(objRSplayers.Fields("barps").Value) > CDbl(objRSplayers.Fields("l5barps").Value) then%>
							<td style="vertical-align:middle;text-align:center;"><i class="fal fa-long-arrow-down" style="color:#9a1400;vertical-align:middle;text-align:center;font-weight:bold;"></i></td>
							<%else%>
							<td style="vertical-align:middle;text-align:center;"><i class="fal fa-arrows-h" style="color:gold;vertical-align:middle;text-align:center;font-weight:bold;"></i></td>
							<%end if %>	
							
							<%if round(objRSplayers.Fields("l5barps").Value,2) >= 45 then%>
							<td class="big" style="text-align:center;color:#6d4c41;"><span style="color:darkorange;"><i class="fas fa-basketball-ball"></i></span>&nbsp;<%=round(objRSplayers.Fields("l5barps").Value,2) %></td>
							<%else%>
								<td class="big" style="text-align:center;color:#6d4c41;"><%=round(objRSplayers.Fields("l5barps").Value,2) %></td>		
							<%end if%>		
						</tr>
						<%
							objRSplayers.MoveNext
							Wend
						%>
					</table>
				</td>
			</tr>	
			<%end if%>	
			</table>
			</br>	
			<center>
				<a class="blue" href="#top">Return to top</a>
			</center>	
			</br>			
			<%
				objRSgames.MoveNext
				objRSplayers.Close
				count = count + 1
				Wend
			%>
		
		</div>
	</div>
</div>
	<!--<div class="row">		
<div class="fa-3x">
  <i class="fasner"></i>
  <i class="fas fa-circle-notch"></i>
  <i class="fas fa-sync"></i>
  <i class="fas fa-cog"></i>
  <i class="fasner fa-pulse"></i>
  <i class="fas fa-basketball-ball "></i>
</div>
	</div>-->
</div>

	<%
objRSgames.Close
objRSplayers.Close
'objRSteams.Close
objconn.close
Set objrsgames = Nothing
Set objRSplayers = Nothing
Set objConn = Nothing
%>
</body>
</html>
