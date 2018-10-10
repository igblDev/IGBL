<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""

    Dim objConn,omitPID,sAction
		Set objConn           = Server.CreateObject("ADODB.Connection")
		Set objRSNBASked      = Server.CreateObject("ADODB.RecordSet")	
		Set objRSPlayers      = Server.CreateObject("ADODB.RecordSet")	
		Set objParams         = Server.CreateObject("ADODB.RecordSet")
		
    objConn.Open Application("lineupstest_ConnectionString")
    objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=lineupstest.mdb;" & _
                  "Persist Security Info=False"
									
		GetAnyParameter "Action", sAction
		
		select case sAction
			case "Reforecast"
			
				selectedPID = Request.Form("omitPlayerPID")
				if selectedPID = "" then 
					omitPID = 0 
				else 
					omitPID = selectedPID
				end if
			
			case "ReforecastAdd"
			
				selectedPID = Request.Form("addPlayerPID")
				if selectedPID = "" then 
					omitPID = 0 
				else 
					omitPID = selectedPID
				end if	
				
			
			case "Submit Lineup"
				sDate    	  	    = Now()
				ownerid     	    = Request.Form("var_ownerid")
				gamedate          = Request.Form("GameDays")				
				c_Time_1159       = "11:59:59 PM"											
												
				wCenterBarps			= Request.Form("wCBarps")   
				wForwardBarps			= Request.Form("wF1Barps")
				wForward2Barps		= Request.Form("wF2Barps")  
				wGuardBarps				= Request.Form("wG1Barps")
				wGuard2Barps			= Request.Form("wG2Barps") 
				
				wCenterPID				= Request.Form("wCPID")
				wForwardPID			  = Request.Form("wF1Pid")
				wForward2PID			= Request.Form("wF2Pid")
				wGuardPID				  = Request.Form("wG1Pid")
				wGuard2PID				= Request.Form("wG2Pid")
		
				'Response.Write "CENTER PID      = "&wCenterPID&".<br>"
				'Response.Write "FORWARD PID     = "&wForwardPID&".<br>"
				'Response.Write "FORWARD2 PID    = "&wForward2PID&".<br>"
				'Response.Write "GUARD PID       = "&wGuardPID&".<br>"
				'Response.Write "GUARD2 PID      = "&wGuard2PID&".<br>"
				'Response.Write "GAME DATE       = "&gamedate&".<br>"
				'*************************************************************************************************
				'** Retrieving the NBA Team ID for each player to Retrieve the players tip time for todays game
				'*************************************************************************************************
				
				objRSPlayers.Open "SELECT * FROM qLineupDaily WHERE PID = "&wCenterPID&"   and qLineupDaily.GameDay = CDATE('"&gamedate&"') ", objConn,3,3,1
				wCenterTip = objRSPlayers.Fields("GameTime").Value
				objRSPlayers.Close

				objRSPlayers.Open "SELECT * FROM qLineupDaily WHERE PID = "&wForwardPID&"  and qLineupDaily.GameDay = CDATE('"&gamedate&"') ", objConn,3,3,1
				wForwardTip = objRSPlayers.Fields("GameTime").Value				
				objRSPlayers.Close

				objRSPlayers.Open "SELECT * FROM qLineupDaily WHERE PID = "&wForward2PID&" and qLineupDaily.GameDay = CDATE('"&gamedate&"') ", objConn,3,3,1
				wForward2Tip = objRSPlayers.Fields("GameTime").Value
				objRSPlayers.Close

				objRSPlayers.Open "SELECT * FROM qLineupDaily WHERE PID = "&wGuardPID&"    and qLineupDaily.GameDay = CDATE('"&gamedate&"') ", objConn,3,3,1
				wGuardTip = objRSPlayers.Fields("GameTime").Value
				objRSPlayers.Close

				objRSPlayers.Open "SELECT * FROM qLineupDaily WHERE PID = "&wGuard2PID&"    and qLineupDaily.GameDay = CDATE('"&gamedate&"') ", objConn,3,3,1
				wGuard2Tip = objRSPlayers.Fields("GameTime").Value
				objRSPlayers.Close

				'Response.Write "GAME DAY TIP    = "&gamedate&".<br>"
				'Response.Write "CENTER TIP      = "&wCenterTip&".<br>"
				'Response.Write "FORWARD TIP     = "&wForwardTip&".<br>"
				'Response.Write "FORWARD2 TIP    = "&wForward2Tip&".<br>"
				'Response.Write "GUARD TIP       = "&wGuardTip&".<br>"
				'Response.Write "GUARD2 TIP      = "&wGuard2Tip&".<br>"
				
				'************************************************************************************
				'**** INSERT FORECASTED LINEUP INTO LINEUP TABLE
				'************************************************************************************					
				
				strSQL = "DELETE FROM tbl_Lineups WHERE tbl_Lineups.GameDay =  #" & gamedate & "#  AND tbl_Lineups.OwnerID = "& ownerid & ";"
				objConn.Execute strSQL

				strSQL ="insert into tbl_lineups(OwnerID,GameDay,sCenter,sCenterBarps,sForward,sForwardBarps,sForward2,sForward2Barps,sGuard,sGuardBarps,sGuard2,sGuard2Barps,sCenterTip,sForwardTip,sForwardTip2,sGuardTip,sGuardTip2,foreCastedLineup) values ('" &_
				ownerid & "', '" &  gamedate & "', '" & wCenterPID & "', '" & wCenterBarps & "', '" & wForwardPID & "', '" & wForwardBarps & "', '" & wForward2PID & "', '" &_
				wForward2Barps & "', '" & wGuardPID & "', '" & wGuardBarps & "', '" & wGuard2PID & "', '" & wGuard2Barps & "', '" &	wCenterTip & "', '" & wForwardTip & "', '" & wForward2Tip & "', '" & wGuardTip & "', '" & wGuard2Tip & "', true)"
				objConn.Execute strSQL		  
				'Response.Write "FORECASTED LINEUP = "&strSQL&".<br>"
				
				strSQL ="insert into tbl_lineups_history(OwnerID,GameDay,sCenter,sCenterBarps,sForward,sForwardBarps,sForward2,sForward2Barps,sGuard,sGuardBarps,sGuard2,sGuard2Barps,sCenterTip,sForwardTip,sForwardTip2,sGuardTip,sGuardTip2,foreCastedLineup) values ('" &_
				ownerid & "', '" &  gamedate & "', '" & wCenterPID & "', '" & wCenterBarps & "', '" & wForwardPID & "', '" & wForwardBarps & "', '" & wForward2PID & "', '" &_
				wForward2Barps & "', '" & wGuardPID & "', '" & wGuardBarps & "', '" & wGuard2PID & "', '" & wGuard2Barps & "', '" &	wCenterTip & "', '" & wForwardTip & "', '" & wForward2Tip & "', '" & wGuardTip & "', '" & wGuard2Tip & "',true)"
				objConn.Execute strSQL
		
		end select
		
		
		
		'********************************************
		'***  CODE MOVED TO SELECT STATEMENT ABOVE
		'********************************************		
		'if sAction = "Reforecast" then	
		'	selectedPID = Request.Form("omitPlayerPID")
		'	if selectedPID = "" then 
		'		omitPID = 0 
		'	else 
		'		omitPID = selectedPID
		'	end if
		'elseif sAction = "ReforecastAdd" then
		'	selectedPID = Request.Form("addPlayerPID")
		'	if selectedPID = "" then 
		'		omitPID = 0 
		'	else 
		'		omitPID = selectedPID
		'	end if	
		'end if-->


									
%>
<!--#include virtual="Common/session.inc"-->
<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
.panel-matchups {
  background-color:white;
  border-color:black;
	color:#01579B;	
}
.panel-nogames {
  background-color:black;
  border-color:black;
	color:yellowgreen;	
}
.alert-info {
    color: #01579B;
    background-color: white;
    border-color: white;
}
red {
	color:#9a1400;
	font-weight: bold;
}

white {
	color:white;
	font-weight: normal;
}
td {
    vertical-align: middle;		
}
.panel-forecast {
  background-color:white;
  border-color:black;
	color:#01579B;
	border-radius: 0;
}
.btn-txn-red {
    color: black;
}		
.btn-txn-green {
		border-radius: 0px !important;
		font-size: 15px;
}	
.bs-callout-success h4 {
    color: black;
}	
.bs-callout-success {
    border-left-color: yellowgreen;
    padding: 10px;
    border-left-width: 5px;
    border-radius: 3px;
    background-color: white;
}
.bs-callout-lineups {

    border-left-color: black;
    padding: 10px;
    border-left-width: 5px;
    border-radius: 3px;
    background-color: white;

}
.blue {
	color: #01579B;
	font-weight: 600;
}
.nav-tabs {
    border-bottom: 2px solid black;
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
.mark, mark {
    padding: .2em;
    background-color: yellow;
}
span.orangeText {
    font-weight: bold;
}
</style>
</head>
<body>
<!--#include virtual="Common/headerMain.inc"-->

<%
					Set objRSMatchups= Server.CreateObject("ADODB.RecordSet")
		
					objRSMatchups.Open "SELECT * FROM qry_matchupLineups", objConn,3,3,1
					gameday = objRSMatchups.Fields("gameday").Value
					gameToday = objRSMatchups.RecordCount
					objRSMatchups.close
%>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
			<strong>Matchups | Lineups</strong>
			</div>
		</div>
	</div>
</div>
<% if gameToday > 0 then %>
<div class="container">
	<div class="row">		
		<div class="col-md-12 col-sm-12 col-xs-12">			
				<span style=""><strong>Game Day Matchups</strong></span><br>
				<%=(FormatDateTime(gameday,1))%>
		</div>
	</div>
</div>
<br>
<%end if%>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<ul class="nav nav-tabs">
			 <% if sAction = "Reforecast" or sAction = "ReforecastAdd" then %>
				<li><a data-toggle="tab" href="#matchups"><i class="far fa-ticket-alt"></i>&nbsp;Matchups</a></li>
				<li class="active"><a data-toggle="tab" href="#VS"><i class="far fa-chart-line"></i>&nbsp;Forecasted</a></li>
				<li><a data-toggle="tab" href="#mine"><i class="fal fa-file-alt"></i>&nbsp;My Lineups</a></li>	
			 <%elseif sAction = "Submit Lineup" then %>
				<li><a data-toggle="tab" href="#matchups"><i class="far fa-ticket-alt"></i>&nbsp;Matchups</a></li>
				<li><a data-toggle="tab" href="#VS"><i class="far fa-chart-line"></i>&nbsp;Forecasted</a></li>
				<li class="active"><a data-toggle="tab" href="#mine"><i class="fal fa-file-alt"></i>&nbsp;My Lineups</a></li>						
			 <%else %>
				<li class="active"><a data-toggle="tab" href="#matchups"><i class="far fa-ticket-alt"></i>&nbsp;Matchups</a></li>
				<li><a data-toggle="tab" href="#VS"><i class="far fa-chart-line"></i>&nbsp;Forecasted</a></li>
				<li><a data-toggle="tab" href="#mine"><i class="fal fa-file-alt"></i>&nbsp;Lineups</a></li>			 
			 <%end if%>
			</ul>
		</div>
	</div>
	</br>
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="tab-content">
			<% if sAction = "Submit Lineup" then %>
			<div id="mine" class="tab-pane fade in active">
			<%else%>
			<div id="mine" class="tab-pane fade">			
			<%end if%>
			<%
						Set objRSLineups      = Server.CreateObject("ADODB.RecordSet")
						Set objRSLineupsPics 	= Server.CreateObject("ADODB.RecordSet")
						
						objRSLineups.Open	"SELECT * FROM qrySubLineups WHERE qrySubLineups.ownerID =" & ownerid & " ", objConn,3,3,1
					%>
					<%
						 While Not objRSLineups.EOF
					%>
					<% 
						if objRSLineups.Fields("gameday").Value = date() and objRSLineups.Fields("gameStaggerDeadline").Value < (time() - 1/24) then
							showPencil    = false
						else
							showPencil    = true 
						end if
						
						cFirstName    = ""
						cLastName     = ""
						sForward2     = ""
						f1LastName    = ""
						sGuard2       = ""
				
						sCenterBarps  = 0
						sForwardBarps = 0
						sForward2Barps= 0
						sGuardBarps   = 0
						sGuard2Barps  = 0

						sCenterTime   = 0
						sForwardTime  = 0
						sForward2Time = 0
						sGuardTime    = 0
						sGuard2Time   = 0
						
						sCenter       = objRSLineups.Fields("sCenter").Value
						sForward      = objRSLineups.Fields("sforward").Value
						sForward2     = objRSLineups.Fields("sforward2").Value
						sGuard        = objRSLineups.Fields("sguard").Value
						sGuard2       = objRSLineups.Fields("sguard2").Value
						
						sCenterBarps  = objRSLineups.Fields("sCenterBarps").Value
						sForwardBarps = objRSLineups.Fields("sforwardBarps").Value
						sForward2Barps= objRSLineups.Fields("sForward2Barps").Value
						sGuardBarps   = objRSLineups.Fields("sguardBarps").Value
						sGuard2Barps  = objRSLineups.Fields("sguard2Barps").Value

						sCenterTime   = objRSLineups.Fields("sCenterTip").Value
						if len(objRSLineups.Fields("sCenterTip").Value) = 10 then
							sCenterTime = Left(objRSLineups.Fields("sCenterTip").Value,4) & Right(objRSLineups.Fields("sCenterTip").Value,3)
						else
							sCenterTime = Left(objRSLineups.Fields("sCenterTip").Value,5) & Right(objRSLineups.Fields("sCenterTip").Value,3)
						end if	
						
						sForwardTime  = objRSLineups.Fields("sForwardTip").Value
						if len(objRSLineups.Fields("sForwardTip").Value) = 10 then
							sForwardTime = Left(objRSLineups.Fields("sForwardTip").Value,4) & Right(objRSLineups.Fields("sForwardTip").Value,3)
						else
							sForwardTime = Left(objRSLineups.Fields("sForwardTip").Value,5) & Right(objRSLineups.Fields("sForwardTip").Value,3)
						end if	
						
						sForward2Time = objRSLineups.Fields("sForwardTip2").Value
						if len(objRSLineups.Fields("sForwardTip2").Value) = 10 then
							sForward2Time = Left(objRSLineups.Fields("sForwardTip2").Value,4) & Right(objRSLineups.Fields("sForwardTip2").Value,3)
						else
							sForward2Time = Left(objRSLineups.Fields("sForwardTip2").Value,5) & Right(objRSLineups.Fields("sForwardTip2").Value,3)
						end if	
						
						sGuardTime    = objRSLineups.Fields("sGuardTip").Value
						if len(objRSLineups.Fields("sGuardTip").Value) = 10 then
							sGuardTime = Left(objRSLineups.Fields("sGuardTip").Value,4) & Right(objRSLineups.Fields("sGuardTip").Value,3)
						else
							sGuardTime = Left(objRSLineups.Fields("sGuardTip").Value,5) & Right(objRSLineups.Fields("sGuardTip").Value,3)
						end if	
						
						sGuard2Time   = objRSLineups.Fields("sGuardTip2").Value
						if len(objRSLineups.Fields("sGuardTip2").Value) = 10 then
							sGuard2Time = Left(objRSLineups.Fields("sGuardTip2").Value,4) & Right(objRSLineups.Fields("sGuardTip2").Value,3)
						else
							sGuard2Time = Left(objRSLineups.Fields("sGuardTip2").Value,5) & Right(objRSLineups.Fields("sGuardTip2").Value,3)
						end if	
										
						objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sCenter  & " = PID " , objConn,3,3,1
						cFirstName = objRSLineupsPics.Fields("firstName").Value
						cLastName  = objRSLineupsPics.Fields("lastName").Value

						objRSLineupsPics.Close
						
						objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sForward  & " = PID " , objConn,3,3,1
						f1FirstName= objRSLineupsPics.Fields("firstName").Value
						f1LastName = objRSLineupsPics.Fields("lastName").Value

						objRSLineupsPics.Close
						
						objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sForward2  & " = PID " , objConn,3,3,1
						f2FirstName= objRSLineupsPics.Fields("firstName").Value
						f2LastName = objRSLineupsPics.Fields("lastName").Value

						objRSLineupsPics.Close
						
						objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sGuard  & " = PID " , objConn,3,3,1
						g1FirstName= objRSLineupsPics.Fields("firstName").Value
						g1LastName = objRSLineupsPics.Fields("lastName").Value

						objRSLineupsPics.Close
						
						objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sGuard2 & " = PID " , objConn,3,3,1
						g2FirstName= objRSLineupsPics.Fields("firstName").Value
						g2LastName = objRSLineupsPics.Fields("lastName").Value

						objRSLineupsPics.Close 
				
					%>
					<table class="table table-custom-black table-bordered table-condensed">
						<tr bgcolor="#FFFFFF">
							<td class="big" colspan="2" style="text-align:center;">
								<%if cLastName = "***CEN***" then %>
									<i class="fas fa-user-slash blue"></i>
								<%else%>
									<a class="blue" href="playerprofile.asp?pid=<%=sCenter %>"><%=left(cFirstName,1)%>.&nbsp;<%=left(cLastName,10)%></a>&nbsp;<span class="orangeText">CEN</span>&nbsp;<span class="gameTip"><small><i class="fas fa-clock"></i>&nbsp;<%=sCenterTime%></small></span>
								<%end if%>
							</td>
						</tr>
						
						<tr bgcolor="#FFFFFF">
							<td class="big" width="50%" style="text-align:center">
								<%if f1LastName = "***FOR***" then %>
									<i class="fas fa-user-slash blue"></i>
								<%else%>
									<a class="blue" href="playerprofile.asp?pid=<%=sForward %>"><%=left(f1FirstName,1)%>.&nbsp;<%=left(f1LastName,10)%></a>&nbsp;<span class="orangeText">FOR</span>&nbsp;<span class="gameTip"><small><i class="fas fa-clock"></i>&nbsp;<%=sForwardTime%></small></span>
								<%end if%>
							</td>
							<td class="big" width="50%" style="text-align:center">
								<%if f2LastName = "***FOR***" then %>							
									<i class="fas fa-user-slash blue"></i>
								<%else%>
									<a class="blue" href="playerprofile.asp?pid=<%=sForward2 %>"><%=left(f2FirstName,1)%>.&nbsp;<%=left(f2LastName,10)%></a>&nbsp;<span class="orangeText">FOR</span>&nbsp;<span class="gameTip"><small><i class="fas fa-clock"></i>&nbsp;<%=sForward2Time%></small></span>
								<%end if%>	
							</td>		
						</tr>
	
						<tr bgcolor="#FFFFFF">
							<td class="big" width="50%" style="text-align:center">
								<%if g1LastName = "***GUA***" then %>
									<i class="fas fa-user-slash blue"></i>
								<%else%>
									<a class="blue" href="playerprofile.asp?pid=<%=sGuard %>"><%=left(g1FirstName,1)%>.&nbsp;<%=left(g1LastName,10)%></a>&nbsp;<span class="orangeText">GUA</span>&nbsp;<span class="gameTip"><small><i class="fas fa-clock"></i>&nbsp;<%=sGuardTime%></small></span>								<%end if%>
							</td>
							<td class="big" width="50%" style="text-align:center">
								<%if g2LastName = "***GUA***" then %>							
									<i class="fas fa-user-slash blue"></i>
								<%else%>
									<a class="blue" href="playerprofile.asp?pid=<%=sGuard2 %>"><%=left(g2FirstName,1)%>.&nbsp;<%=left(g2LastName,10)%></a>&nbsp;<span class="orangeText">GUA</span>&nbsp;<span class="gameTip"><small><i class="fas fa-clock"></i>&nbsp;<%=sGuard2Time%></small></span>								<%end if%>	
							</td>		
						</tr>
						
						<% if showPencil = true then %>					
						<tr bgcolor="white" class="font-weight:bold;">
							<td colspan="3"><a href="dashboard.asp?ownerid=<%= ownerid %>&Action=Retrieve Lineup&currentDate=<%= objRSLineups.Fields("gameday").Value%>"><button type="submit" class="btn btn-block  btn-default"><i class="fa fa-pencil-square-o red" aria-hidden="true"></i>&nbsp;EDIT <%= objRSLineups.Fields("gameday").Value %></button></a></td>
						</tr>
						<%else%>
							<td style="font-weight:bold;background-color: yellow;vertical-align: middle;text-align:center;" colspan="3"><span style="font-weight:bold;background-color: yellow;vertical-align: middle;">Game Deadline Passed!</span></td>
						<%end if %>
					</table>
					<br>
					<%
					objRSLineups.MoveNext	
					Wend
					%>
			
					<%
						objRSLineups.Close
						Set objRS = Nothing
						Set objRSLineupsPics = Nothing
					%>

				</div>	
					
			 <% if sAction = "Reforecast" or sAction = "ReforecastAdd" or sAction = "Submit Lineup" then %>
					<div id="matchups" class="tab-pane fade">
				<%else %>
					<div id="matchups" class="tab-pane fade in active">				
				<%end if%>
				
				<% 
					Set objRSMatchups= Server.CreateObject("ADODB.RecordSet")
					Set objRSBarps   = Server.CreateObject("ADODB.RecordSet")
					Set objRSPics    = Server.CreateObject("ADODB.RecordSet")   
					Set objRSTeamLogos   = Server.CreateObject("ADODB.RecordSet")
					Set objRSNBASked  = Server.CreateObject("ADODB.RecordSet")
					
					objRSMatchups.Open "SELECT * FROM qry_matchupLineups", objConn,3,3,1
					gameday = objRSMatchups.Fields("gameday").Value

					dim loopcnt 
					loopcnt = 1
				%>
				<!--<table class="table table-custom-black table-bordered table-condensed">
					<% if objRSMatchups.RecordCount > 0 then %>				
					<tr style="background-color:white;text-align:center;vertical-align:middle">
						<th style="width:50%;color:#a89a7a;font-weight:bold;text-align:center;"><i class="fa fa-user" aria-hidden="true"></i>&nbsp;IDLE PLAYER</th>
						<th style="width:50%;color:yellowgreen;font-weight:bold;text-align:center;">ACTIVE PLAYER&nbsp;<i class="fa fa-user" aria-hidden="true"></i></th>
					</tr>
					<%end if%>
				</table>-->
				<% if objRSMatchups.RecordCount > 0 then %>				
				<% 
				
				While Not objRSMatchups.EOF
	
				hometeam      = objRSMatchups.Fields("HomeTeam").Value
				hometeamshort = objRSMatchups.Fields("HomeTeamShort").Value
				homeowner     = objRSMatchups.Fields("HomeOwner").Value
				hometeampen   = objRSMatchups.Fields("HomeTeamPen").Value
				awayteam      = objRSMatchups.Fields("VisitingTeam").Value
				awayteamshort = objRSMatchups.Fields("VisitingTeamShort").Value
				awayowner     = objRSMatchups.Fields("VisitingOwner").Value
				awayteampen   = objRSMatchups.Fields("VisitingTeamPen").Value
				
				objRSTeamLogos.Open "Select TeamLogo  from TBLOWNERS where  " & homeowner  & " = ownerID " , objConn,3,3,1
				homeLogo = objRSTeamLogos.Fields("TeamLogo").Value
				objRSTeamLogos.Close		
				
				objRSTeamLogos.Open "Select TeamLogo  from TBLOWNERS where  " & awayowner  & " = ownerID " , objConn,3,3,1
				awayLogo = objRSTeamLogos.Fields("TeamLogo").Value
				objRSTeamLogos.Close		
				
				homerec  = objRSMatchups.Fields("HomeRecord").Value
				awayrec  = objRSMatchups.Fields("VisRecord").Value
				homeStamp= objRSMatchups.Fields("HomeStamp").Value
				awayStamp= objRSMatchups.Fields("Timestamp").Value
				'******************************
				'** Retrieving Starters PID  **
				'******************************
				sCenterAway   = objRSMatchups.Fields("sCenter").Value
				sForward1Away = objRSMatchups.Fields("sForward").Value
				sForward2Away = objRSMatchups.Fields("sForward2").Value
				sGuard1Away   = objRSMatchups.Fields("sGuard").Value
				sGuard2Away   = objRSMatchups.Fields("sGuard2").Value	
				sCenterHome   = objRSMatchups.Fields("HomeC").Value
				sForward1Home = objRSMatchups.Fields("HomeF").Value
				sForward2Home = objRSMatchups.Fields("HomeF2").Value
				sGuard1Home   = objRSMatchups.Fields("HomeG").Value
				sGuard2Home   = objRSMatchups.Fields("HomeG2").Value
				lineupTimeChk = time() - 1/24	
				
				homePlayerCnt = 0
				awayPlayerCnt = 0 
				
				cTip = objRSMatchups.Fields("sCenterTip").Value
				if len(objRSMatchups.Fields("sCenterTip").Value) = 10 then
					cTip = Left(objRSMatchups.Fields("sCenterTip").Value,4) & Right(objRSMatchups.Fields("sCenterTip").Value,3)
				else
					cTip = Left(objRSMatchups.Fields("sCenterTip").Value,5) & Right(objRSMatchups.Fields("sCenterTip").Value,3)
				end if	

				if 	CDATE(cTip) < CDATE(lineupTimeChk) then 	
					awayCenTipDeadlinePassed = true
					awayPlayerCnt = awayPlayerCnt + 1
				end if				
				
				for1Tip       = objRSMatchups.Fields("sForwardTip").Value
				if len(objRSMatchups.Fields("sForwardTip").Value) = 10 then
					for1Tip = Left(objRSMatchups.Fields("sForwardTip").Value,4) & Right(objRSMatchups.Fields("sForwardTip").Value,3)
				else
					for1Tip = Left(objRSMatchups.Fields("sForwardTip").Value,5) & Right(objRSMatchups.Fields("sForwardTip").Value,3)
				end if	
				
				if 	CDATE(for1Tip) < CDATE(lineupTimeChk) then 	
					awayFor1TipDeadlinePassed = true
					awayPlayerCnt = awayPlayerCnt + 1
				end if
				
				for2Tip       = objRSMatchups.Fields("sForwardTip2").Value
				if len(objRSMatchups.Fields("sForwardTip2").Value) = 10 then
					for2Tip = Left(objRSMatchups.Fields("sForwardTip2").Value,4) & Right(objRSMatchups.Fields("sForwardTip2").Value,3)
				else
					for2Tip = Left(objRSMatchups.Fields("sForwardTip2").Value,5) & Right(objRSMatchups.Fields("sForwardTip2").Value,3)
				end if	
				
				if 	CDATE(for2Tip) < CDATE(lineupTimeChk) then 	
					awayFor2TipDeadlinePassed = true
					awayPlayerCnt = awayPlayerCnt + 1
				end if
	
				gua1Tip       = objRSMatchups.Fields("sGuardTip").Value
				if len(objRSMatchups.Fields("sGuardTip").Value) = 10 then
					gua1Tip = Left(objRSMatchups.Fields("sGuardTip").Value,4) & Right(objRSMatchups.Fields("sGuardTip").Value,3)
				else
					gua1Tip = Left(objRSMatchups.Fields("sGuardTip").Value,5) & Right(objRSMatchups.Fields("sGuardTip").Value,3)
				end if	

				if 	CDATE(gua1Tip) < CDATE(lineupTimeChk) then 	
					awayGua1TipDeadlinePassed = true
					awayPlayerCnt = awayPlayerCnt + 1
				end if
				
				gua2Tip       = objRSMatchups.Fields("sGuardTip2").Value
				if len(objRSMatchups.Fields("sGuardTip2").Value) = 10 then
					gua2Tip = Left(objRSMatchups.Fields("sGuardTip2").Value,4) & Right(objRSMatchups.Fields("sGuardTip2").Value,3)
				else
					gua2Tip = Left(objRSMatchups.Fields("sGuardTip2").Value,5) & Right(objRSMatchups.Fields("sGuardTip2").Value,3)
				end if	

				if 	CDATE(gua2Tip) < CDATE(lineupTimeChk) then 	
					awayGua2TipDeadlinePassed = true
					awayPlayerCnt = awayPlayerCnt + 1
				end if
	
				
				hCenTip       = objRSMatchups.Fields("HomeCTip").Value
				if len(objRSMatchups.Fields("HomeCTip").Value) = 10 then
					hCenTip = Left(objRSMatchups.Fields("HomeCTip").Value,4) & Right(objRSMatchups.Fields("HomeCTip").Value,3)
				else
					hCenTip = Left(objRSMatchups.Fields("HomeCTip").Value,5) & Right(objRSMatchups.Fields("HomeCTip").Value,3)
				end if	
				
				if 	CDATE(hCenTip) < CDATE(lineupTimeChk) then 	
					homeCenTipDeadlinePassed = true
					homePlayerCnt = homePlayerCnt + 1
				end if	
				
				hForTip       = objRSMatchups.Fields("HomeFTip").Value
				if len(objRSMatchups.Fields("HomeFTip").Value) = 10 then
					hForTip = Left(objRSMatchups.Fields("HomeFTip").Value,4) & Right(objRSMatchups.Fields("HomeFTip").Value,3)
				else
					hForTip = Left(objRSMatchups.Fields("HomeFTip").Value,5) & Right(objRSMatchups.Fields("HomeFTip").Value,3)
				end if	

				if 	CDATE(hForTip) < CDATE(lineupTimeChk) then 	
					homeFor1TipDeadlinePassed = true
					homePlayerCnt = homePlayerCnt + 1
				end if
				
				
				hFor2Tip      = objRSMatchups.Fields("HomeF2Tip").Value
				if len(objRSMatchups.Fields("HomeF2Tip").Value) = 10 then
					hFor2Tip = Left(objRSMatchups.Fields("HomeF2Tip").Value,4) & Right(objRSMatchups.Fields("HomeF2Tip").Value,3)
				else
					hFor2Tip = Left(objRSMatchups.Fields("HomeF2Tip").Value,5) & Right(objRSMatchups.Fields("HomeF2Tip").Value,3)
				end if	

				if 	CDATE(hFor2Tip) < CDATE(lineupTimeChk) then 	
					homeFor2TipDeadlinePassed = true	
					homePlayerCnt = homePlayerCnt + 1
				end if				
				
				
				hGuaTip       = objRSMatchups.Fields("HomeGTip").Value
				if len(objRSMatchups.Fields("HomeGTip").Value) = 10 then
					hGuaTip = Left(objRSMatchups.Fields("HomeGTip").Value,4) & Right(objRSMatchups.Fields("HomeGTip").Value,3)
				else
					hGuaTip = Left(objRSMatchups.Fields("HomeGTip").Value,5) & Right(objRSMatchups.Fields("HomeGTip").Value,3)
				end if	

				if 	CDATE(hGuaTip) < CDATE(lineupTimeChk) then 	
					homeGua1TipDeadlinePassed = true
					homePlayerCnt = homePlayerCnt + 1
				end if
					
				hGua2Tip      = objRSMatchups.Fields("HomeG2Tip").Value 
				if len(objRSMatchups.Fields("HomeG2Tip").Value) = 10 then
					hGua2Tip = Left(objRSMatchups.Fields("HomeG2Tip").Value,4) & Right(objRSMatchups.Fields("HomeG2Tip").Value,3)
				else
					hGua2Tip = Left(objRSMatchups.Fields("HomeG2Tip").Value,5) & Right(objRSMatchups.Fields("HomeG2Tip").Value,3)
				end if	
				
				if 	CDATE(hGua2Tip) < CDATE(lineupTimeChk) then 	
					homeGua2TipDeadlinePassed = true
					homePlayerCnt = homePlayerCnt + 1
				end if
				
				ACenBarps  		= 0
				AFor1Barps 		= 0
				AFor2Barps 		= 0
				AGua1Barps 		= 0
				AGua2Barps 		= 0
				HCenBarps  		= 0
				HFor1Barps 		= 0
				HFor2Barps 		= 0
				HGua1Barps 		= 0
				HGua2Barps 		= 0
				
				objRSPics.Open "Select firstName,lastName,NBATeamID from tblPlayers where  " & sCenterAway  & " = PID " , objConn,3,3,1
				CenterAwayNBATM= objRSPics.Fields("NBATeamID").Value

				CenterAwayNF  = left(objRSPics.Fields("firstName").Value,12)				
				CenterAwayNL  = objRSPics.Fields("lastName").Value
				
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1


				if IsNull(objRSNBASked.Fields("opponent").value) or objRSNBASked.Fields("opponent").value = " " then 
					CenterAwayOpp = "GHOST"
				else
					CenterAwayOpp = objRSNBASked.Fields("opponent").value
				end if 
				
				objRSNBASked.Close
				objRSPics.Close  

				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sCenterAway  & " = PID " , objConn,3,3,1
				ACenBarps     = round(objRSBarps.Fields("barps").Value,2)
				ACenTM = objRSBarps.Fields("team").Value
				objRSBarps.Close
				
				objRSPics.Open "Select firstName,lastName,NBATeamID   from tblplayers where " & sForward1Away & " = PID" , objConn,3,3,1
				Forward1AwayNBATM= objRSPics.Fields("NBATeamID").Value
				Forward1AwayNF= left(objRSPics.Fields("firstName").Value,12)
				Forward1Away= objRSPics.Fields("lastName").Value
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				'Forward1AwayOpp = objRSNBASked.Fields("opponent").value

				if IsNull(objRSNBASked.Fields("opponent").value) or objRSNBASked.Fields("opponent").value = " " then 
					Forward1AwayOpp = "GHOST"
				else
					Forward1AwayOpp = objRSNBASked.Fields("opponent").value
				end if 
				
				objRSNBASked.Close
				objRSPics.Close				

				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sForward1Away & " = PID " , objConn,3,3,1
				AFor1Barps    = round(objRSBarps.Fields("barps").Value,2)
				AFor1TM = objRSBarps.Fields("team").Value
				objRSBarps.Close

				objRSPics.Open "Select firstName,lastName,NBATeamID   from tblplayers where " & sForward2Away &" = PID" , objConn,3,3,1
				Forward2Away  = objRSPics.Fields("lastName").Value
				Forward2AwayNF= left(objRSPics.Fields("firstName").Value,12)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				
				if IsNull(objRSNBASked.Fields("opponent").value) or objRSNBASked.Fields("opponent").value = " " then 
					Forward2AwayOpp = "GHOST"
				else
					Forward2AwayOpp = objRSNBASked.Fields("opponent").value
				end if	
				
				objRSNBASked.Close				
				objRSPics.Close 

				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sForward2Away & " = PID " , objConn,3,3,1
				AFor2Barps    = round(objRSBarps.Fields("barps").Value,2)
				AFor2TM = objRSBarps.Fields("team").Value
				objRSBarps.Close
				
				objRSPics.Open "Select firstName,lastName,NBATeamID   from tblplayers where  " & sGuard1Away & " = PID" , objConn,3,3,1
				Guard1Away    = objRSPics.Fields("lastName").Value
				Guard1AwayNF  = left(objRSPics.Fields("firstName").Value,12)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				'Guard1AwayOpp = objRSNBASked.Fields("opponent").value

				if IsNull(objRSNBASked.Fields("opponent").value) or objRSNBASked.Fields("opponent").value = " " then 
					Guard1AwayOpp = "GHOST"
				else
					Guard1AwayOpp = objRSNBASked.Fields("opponent").value
				end if
				
				objRSNBASked.Close
				objRSPics.Close 
				
				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sGuard1Away & " = PID " , objConn,3,3,1
				AGua1Barps    = round(objRSBarps.Fields("barps").Value,2)
				AGua1TM = objRSBarps.Fields("team").Value
				objRSBarps.Close

				objRSPics.Open "Select firstName,lastName,NBATeamID   from tblplayers where " & sGuard2Away & "  = PID" , objConn,3,3,1
				Guard2Away    = objRSPics.Fields("lastName").Value
				Guard2AwayNF  = left(objRSPics.Fields("firstName").Value,12)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				'Guard2AwayOpp = objRSNBASked.Fields("opponent").value

				if IsNull(objRSNBASked.Fields("opponent").value) or objRSNBASked.Fields("opponent").value = " " then 
					Guard2AwayOpp = "GHOST"
				else
					Guard2AwayOpp = objRSNBASked.Fields("opponent").value
				end if
				
				objRSNBASked.Close
				objRSPics.Close
				
				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sGuard2Away & " = PID " , objConn,3,3,1
				AGua2Barps    = round(objRSBarps.Fields("barps").Value,2)
				AGua2TM = objRSBarps.Fields("team").Value
				objRSBarps.Close
		
				'******************************************************************************************************
				'**** RETRIEVING THE BARPS AD THE IMAGES FOR THE STARTING PLAYERS ON THE HOME TEAM
				'******************************************************************************************************
				
				objRSPics.Open "Select firstName,lastName,NBATeamID from tblplayers where " & sCenterHome & "  = PID" , objConn,3,3,1
				CenterHomeN = objRSPics.Fields("lastName").Value
				CenterHomeNF= left(objRSPics.Fields("firstName").Value,12)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				'CenterHomeOpp = objRSNBASked.Fields("opponent").value

				if IsNull(objRSNBASked.Fields("opponent").value) or objRSNBASked.Fields("opponent").value = " " then 
					CenterHomeOpp = "GHOST"
				else
					CenterHomeOpp = objRSNBASked.Fields("opponent").value
				end if	
				
				objRSNBASked.Close
				objRSPics.Close
				
				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sCenterHome & " = PID " , objConn,3,3,1
				HCenBarps  = round(objRSBarps.Fields("barps").Value,2)
				HCenTM = objRSBarps.Fields("team").Value
				objRSBarps.Close
				
				objRSPics.Open "Select firstName,lastName,NBATeamID from tblplayers where " & sForward1Home & " = PID" , objConn,3,3,1
				Forward1HomeN = objRSPics.Fields("lastName").Value
				Forward1HomeNF = left(objRSPics.Fields("firstName").Value,12)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				'Forward1HomeOpp = objRSNBASked.Fields("opponent").value
				
				if IsNull(objRSNBASked.Fields("opponent").value) or objRSNBASked.Fields("opponent").value = " " then 
					Forward1HomeOpp = "GHOST"
				else
					Forward1HomeOpp = objRSNBASked.Fields("opponent").value
				end if	
				objRSNBASked.Close
				objRSPics.Close
				
				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sForward1Home & " = PID " , objConn,3,3,1
				HGua2HomeNBATM= objRSPics.Fields("NBATeamID").Value
				HFor1Barps  = round(objRSBarps.Fields("barps").Value,2)
				HFor1TM = objRSBarps.Fields("team").Value
				objRSBarps.Close

				objRSPics.Open "Select firstName,lastName,NBATeamID from tblplayers where " & sForward2Home & " = PID" , objConn,3,3,1
				Forward2HomeN = objRSPics.Fields("lastName").Value
				Forward2HomeNF= left(objRSPics.Fields("firstName").Value,12)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1

				if IsNull(objRSNBASked.Fields("opponent").value) or objRSNBASked.Fields("opponent").value = " " then 
					Forward2HomeOpp = "GHOST"
				else
					Forward2HomeOpp = objRSNBASked.Fields("opponent").value
				end if	

				objRSNBASked.Close
				objRSPics.Close	
				
				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sForward2Home & " = PID " , objConn,3,3,1
				HFor2Barps = round(objRSBarps.Fields("barps").Value,2)
				HFor2TM = objRSBarps.Fields("team").Value
				objRSBarps.Close
				
				objRSPics.Open "Select firstName,lastName,NBATeamID from tblplayers where " & sGuard1Home & " = PID" , objConn,3,3,1
				Guard1HomeN = objRSPics.Fields("lastName").Value
				Guard1HomeNF= left(objRSPics.Fields("firstName").Value,12)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1

				if IsNull(objRSNBASked.Fields("opponent").value) or objRSNBASked.Fields("opponent").value = " " then 
					Guard1HomeOpp = "GHOST"
				else
					Guard1HomeOpp = objRSNBASked.Fields("opponent").value
				end if	

				objRSNBASked.Close
				objRSPics.Close	
				
				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sGuard1Home & " = PID " , objConn,3,3,1
				HGua2HomeNBATM= objRSPics.Fields("NBATeamID").Value
				HGua1Barps = round(objRSBarps.Fields("barps").Value,2)
				HGua1TM = objRSBarps.Fields("team").Value
				objRSBarps.Close

				objRSPics.Open "Select firstName,lastName,NBATeamID from tblplayers where " & sGuard2Home & " = PID" , objConn,3,3,1
				Guard2HomeN   = objRSPics.Fields("lastName").Value
				Guard2HomeNF  = left(objRSPics.Fields("firstName").Value,12)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1

				if IsNull(objRSNBASked.Fields("opponent").value) or objRSNBASked.Fields("opponent").value = " " then 
					Guard2HomeOpp = "GHOST"
				else
					Guard2HomeOpp = objRSNBASked.Fields("opponent").value
				end if	

				objRSNBASked.Close
				objRSPics.Close
				
				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sGuard2Home & " = PID " , objConn,3,3,1
				HGua2Barps = round(objRSBarps.Fields("barps").Value,2)
				HGua2TM = objRSBarps.Fields("team").Value
				objRSBarps.Close
				
	
				htBarps = cDbl(HGua1Barps)  + cDbl(HGua2Barps) + cDbl(HFor1Barps)  + cDbl(HFor2Barps) + cDbl(HCenBarps)
				atBarps = cDbl(AGua1Barps)  + cDbl(AGua2Barps) + cDbl(AFor1Barps)  + cDbl(AFor2Barps) + cDbl(ACenBarps)		

				%>

						<table class="table table-custom-black table-bordered table-condensed">
							<tr style="text-align:center;vertical-align:middle;background-color:black;font-weight:bold;">
								<td style="color:#a89a7a;">Active Player Count: <span style="color:white;"><%=homePlayerCnt%></span></td>
								<td style="color:#a89a7a;">Active Player Count: <span style="color:white;"><%=awayPlayerCnt%></span></td>
							</tr>
							<tr  style="background-color:white;text-align:center;vertical-align:middle">
								<% if hometeampen = true Then %>
									<td style="width:50%" align="center"><img class="img-responsive" src="images/penalty.png" style="width:151x;height:100px;"></td>
								<% else %>
									<td style="width:50%" align="center"><img class="img-responsive" src="<%=homeLogo%>" style="width:151x;height:100px;"></td>
								<%end if %>
								<% if awayteampen = true Then %>
									<td style="width:50%" align="center"><img class="img-responsive" src="images/penalty.png" style="width:151x;height:100px;"></td>
								<% else %>
									<td style="width:50%" align="center"><img class="img-responsive" src="<%=awayLogo%>" style="width:151x;height:100px;"></td>
								<%end if %>
							</tr>
						</table>
						<table class="table table-custom table-bordered table-condensed">
						<tr style="background-color:white;text-align:center;vertical-align:middle">
							<td style="width:25%">
								<% if hometeampen = true Then %>
								<table class="table table-custom-black table-responsive table-condensed">
									<tr>
										<th class="text-center" style="vertical-align:middle;color:#9a1400;"><%=hometeamShort%></br><%=homerec%></th>
									</tr>
								</table>
							<% else %>
								<table class="table table-custom table-responsive table-condensed">
									<tr>
										<th class="text-center" style="vertical-align:middle;"><%=hometeamShort%><br><%=homerec%></th>
									</tr>
								</table>
							<%end if %>	
							</td>
							<%if CDbl(atBarps) > CDbl(htBarps) then%>
							<td style="width:25%;">proj.</br><span class="badgeDown"><%=round(htBarps,2)%></span></td>
							<%elseif CDbl(htBarps) > CDbl(atBarps) then%>
							<td style="width:25%;">proj.</br><span class="badgeUp"><%=round(htBarps,2)%></span></td>
							<%else%>
							<td style="width:25%;">proj.</br><span class="badgeEven"><%=round(htBarps,2)%></span></td>
							<%end if%>

							<%if CDbl(atBarps) > CDbl(htBarps) then%>
							<td style="width:25%;">proj.</br><span class="badgeUp"><%=round(atBarps,2)%></span></td>
							<%elseif CDbl(htBarps) > CDbl(atBarps) then%>
							<td style="width:25%;">proj.</br><span class="badgeDown"><%=round(atBarps,2)%></span></td>
							<%else%>
							<td style="width:25%;">proj.</br><span class="badgeEven"><%=round(atBarps,2)%></span></td>
							<%end if%>	
							<td style="width:25%">
								<% if awayteampen = true Then %>
										<table class="table table-custom table-responsive table-condensed">
									<tr>
										<th class="text-center" style="vertical-align:middle;color:#9a1400;"><%=awayteamShort%></br> <%=awayrec%></th>
									</tr>
								</table>
								<%else %>
									<table class="table table-custom table-responsive table-condensed">
										<tr>
											<th class="text-center" style="vertical-align:middle;"><%=awayteamShort%><br><%=awayrec%></th>
										</tr>
									</table>								
								<% end if%>
							</td>
						</tr>
					</table>	
					<table class="table table-custom table-bordered table-condensed">	
					<!-- CENTERS -->	
					<tr style="vertical-align:middle;background-color:white;">						
						<%if CenterHomeN = "***CEN***" then%>
						<td class="text-left" style="color:white;background-color:#9a1400;width:50%;">
							<redIcon2><i class="fas fa-user-slash"></i></redIcon2></br>
							<small><%=CenterHomeOpp %>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=hCenTip%></small>
						</td>	
						<%elseif homeCenTipDeadlinePassed then%>
						<td class="text-left big" style="background-color:black;color:#a89a7a;font-weight:bold;width:50%;">
							<span style="color:white;">CEN</span>&nbsp;<span style="font-weight:bold;color:yellowgreen;"><%=left(CenterHomeNF,1) %>.&nbsp;<%=CenterHomeN %></span><br>
							<small><%=CenterHomeOpp %>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=hCenTip%></small>
						</td>							
						<%else%>
						<td class="text-left" style="width:50%;">
							<span class="orangeText">CEN</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=sCenterHome %>"><%=(left(CenterHomeNF,1))%>.&nbsp;<%=CenterHomeN %></a>&nbsp;<%=HCenBarps%></br>	
							<span class="gameTip"><small><%=CenterHomeOpp %>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=hCenTip%></small></span>
						</td>
						<%end if%>					
						<!-- CENTER 2-->						
						<%if CenterAwayNL = "***CEN***" then%>
							<td class="text-left" style="color:white;background-color:#9a1400;width:50%;">
							<redIcon2><i class="fas fa-user-slash"></i></redIcon2></br>
							<small><%=CenterAwayOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=cTip%></small>
						</td>	
						<%elseif awayCenTipDeadlinePassed then%>
						<td class="text-left  big" style="width:50%;background-color:black;color:#a89a7a;font-weight:bold;">
							<span style="color:white;">CEN</span>&nbsp;<span style="font-weight:bold;color:yellowgreen;"><%=CenterAwayNF %>.&nbsp;<%=CenterAwayNL %></span><br>
							<small><%=CenterAwayOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=cTip%></small>		
						</td>			
						<%else%>
						<td class="text-left" style="width:50%;">
							<span class="orangeText">CEN</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=sCenterAway %>"><%=(left(CenterAwayNF,1))%>.&nbsp;<%=CenterAwayNL %></a>&nbsp;<%=ACenBarps%></br>
							<span class="gameTip"><small><%=CenterAwayOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=cTip%></small></span>
						</td>
						<%end if%>						
					</tr>
					<!-- FORWARD 1 -->	
					<tr style="vertical-align:middle;background-color:white;">
						<%if Forward1HomeN = "***FOR***"   then%>
							<td class="text-left" style="color:white;background-color:#9a1400;width:50%;">
								<redIcon2><i class="fas fa-user-slash"></i></redIcon2><br>
								<small><%=Forward1HomeOpp  %>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=hForTip%></small>
							</td>
						<%elseif homeFor1TipDeadlinePassed then%>
							<td class="text-left  big" style="background-color:black;color: #a89a7a;font-weight:bold;">
								<span style="color:white;">FOR</span>&nbsp;<span style="font-weight:bold;color:yellowgreen;"><%=left(Forward1HomeNF,1) %>.&nbsp;<%=Forward1HomeN %></span><br>
								<small><%=Forward1HomeOpp  %>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=hForTip%></small>
							</td>					
						<%else%>
							<td class="text-left">
								<span class="orangeText">FOR</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=sForward1Home %>"><%=(left(Forward1HomeNF,1))%>.&nbsp;<%=Forward1HomeN %></a>&nbsp;<%=HFor1Barps%></br>
								<span class="gameTip"><small><%=Forward1HomeOpp  %>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=hForTip%></small></span>
							</td>
						<%end if%>
						
						<%if Forward1Away = "***FOR***" then%>
							<td class="text-left" style="color:white;background-color:#9a1400;width:50%;">
								<redIcon2><i class="fas fa-user-slash"></i></redIcon2><br>
								<small><%=Forward1AwayOpp %>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=for1Tip%></small>
							</td>
						<%elseif awayFor1TipDeadlinePassed then%>
						  <td class="text-left  big" style="background-color:black;color: #a89a7a;font-weight:bold;">
								<span style="color:white;">FOR</span>&nbsp;<span style="font-weight:bold;color:yellowgreen;"><%=left(Forward1AwayNF,1) %>.&nbsp;<%=Forward1Away %></span><br>
								<small><%=Forward1AwayOpp  %>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=for1Tip%></small>
							</td>
						<%else%>
							<td class="text-left">
								<span class="orangeText">FOR</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=sForward1Away %>"><%=(left(Forward1AwayNF,1))%>.&nbsp;<%=Forward1Away %></a>&nbsp;<%=AFor1Barps%></br>								
								<span class="gameTip"><small><%=Forward1AwayOpp  %>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=for1Tip%></small></span>
							</td>
						<%end if%>						
					</tr>
					<!-- FORWARD 2 -->	
					<tr style="vertical-align:middle;background-color:white;">					
						<% if Forward2HomeN = "***FOR***" then  %>
						<td class="text-left" style="color:white;background-color:#9a1400;width:50%;">
							<redIcon2><i class="fas fa-user-slash"></i></redIcon2><br>
							<small><%=Forward2HomeOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=hFor2Tip%></small>
						</td>
						<%elseif homeFor2TipDeadlinePassed then %>
						<td class="text-left  big" style="background-color:black;color: #a89a7a;font-weight:bold;">
							<span style="color:white;">FOR</span>&nbsp;<span style="font-weight:bold;color:yellowgreen;"><%=left(Forward2HomeNF,1) %>.&nbsp;<%=Forward2HomeN %></span><br>
							<small><%=Forward2HomeOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=hFor2Tip%></small>
						</td>
						<%else%>
						<td class="text-left">
							<span class="orangeText">FOR</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=sForward2Home %>"><%=(left(Forward2HomeNF,1))%>.&nbsp;<%=Forward2HomeN %></a>&nbsp;<%=HFor2Barps%></br>							
							<span class="gameTip"><small><%=Forward2HomeOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=hFor2Tip%></small></span>
						</td>
						<%end if %>

						<% if Forward2Away = "***FOR***" then  %>
						<td class="text-left" style="color:white;background-color:#9a1400;width:50%;">
							<redIcon2><i class="fas fa-user-slash"></i></redIcon2></br>
							<small><%=Forward2AwayOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=for2Tip%></small>
						</td>	
						<%elseif awayFor2TipDeadlinePassed then%>
						<td class="text-left  big" style="background-color:black;color: #a89a7a;font-weight:bold;">
							<span style="color:white;">FOR</span>&nbsp;<span style="font-weight:bold;color:yellowgreen;"><%=Forward2AwayNF %>.&nbsp;<%=Forward2Away %></span><br>
							<small><%=Forward2AwayOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=for2Tip%></small>
						</td>	
						<%else%>
						<td class="text-left">
							<span class="orangeText">FOR</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=sForward2Away %>"><%=(left(Forward2AwayNF,1))%>&nbsp;<%=Forward2Away %></a>&nbsp;<%=AFor2Barps%></br>							
							<span class="gameTip"><small><%=Forward2AwayOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=for2Tip%></small></span>
						</td>	
						<%end if %>						
					</tr>
					<!--GUARDS 1  -->
					<tr style="vertical-align:middle;background-color:white;">
  					<%if Guard1HomeN = "***GUA***" then%>
						<td class="text-left" style="color:white;background-color:#9a1400;width:50%;">
							<redIcon2><i class="fas fa-user-slash"></i></redIcon2></br>
							<%=Guard1HomeOpp  %>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=hGuaTip%>
						</td>
						<%elseif homeGua1TipDeadlinePassed then %>
					 <td class="text-left  big" style="background-color:black;color: #a89a7a;font-weight:bold;">
							<span style="color:white;">GUA</span>&nbsp;<span style="font-weight:bold;color:yellowgreen;"><%=left(Guard1HomeNF,1) %>.&nbsp;<%=Guard1HomeN %></span><br>
							<small><%=Guard1HomeOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=hGuaTip%>
						</td>						
						<%else%>
						<td class="text-left">
							<span class="orangeText">GUA</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=sGuard1Home %>"><%=(left(Guard1HomeNF,1)) %>.&nbsp;<%=Guard1HomeN %></a>&nbsp;<%=HGua1Barps%></br>							
							<span class="gameTip"><small><%=Guard1HomeOpp  %>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=hGuaTip%></small></span>
						</td>
						<%end if%>
						
						<!--GUARD AWAY 2-->
						<%if Guard1Away = "***GUA***" then%>
						<td class="text-left" style="color:white;background-color:#9a1400;width:50%;">
							<redIcon2><i class="fas fa-user-slash"></i></redIcon2></br>
							</small><%=Guard1AwayOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=gua1Tip%></small>
						</td>	
						<%elseif awayGua1TipDeadlinePassed then %>
						<td class="text-left  big" style="background-color:black;color: #a89a7a;font-weight:bold;">
							<span style="color:white;">GUA</span>&nbsp;<span style="font-weight:bold;color:yellowgreen;"><%=left(Guard1AwayNF,1) %>.&nbsp;<%=Guard1Away %></span><br>
							<small><%=Guard1AwayOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=gua1Tip%></small>
						</td>						
						<%else%>
						<td class="text-left">
							<span class="orangeText">GUA</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=sGuard1Away %>"><%=(left(Guard1AwayNF,1))%>.&nbsp;<%=Guard1Away %></a>&nbsp;<%=AGua1Barps%></br>							
							<span class="gameTip"><small><%=Guard1AwayOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=gua1Tip%></small></span>
						</td>	
						<%end if%>
					</tr>
					<!--GUARDS 2  -->
					<tr style="vertical-align:middle;background-color:white;">
  					<%if Guard2HomeN = "***GUA***" then%>
							<td class="text-left" style="color:white;background-color:#9a1400;width:50%;">
							<redIcon2><i class="fas fa-user-slash"></i></redIcon2></br>
							<small><%=Guard2HomeOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=hGua2Tip %></small>
						</td>
						<%elseif homeGua2TipDeadlinePassed then %>
						<td class="text-left  big" style="background-color:black;color: #a89a7a;font-weight:bold;">
							<span style="color:white;">GUA</span>&nbsp;<span style="font-weight:bold;color:yellowgreen;"><%=left(Guard2HomeNF,1) %>.&nbsp;<%=Guard2HomeN %></span><br>
						  <small><%=Guard2HomeOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=hGua2Tip %></small>
						</td>
						<%else%>
						<td class="text-left">
							<span class="orangeText">GUA</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=sGuard2Home%>"><%=(left(Guard2HomeNF,1))%>.&nbsp;<%=Guard2HomeN %></a>&nbsp;<%=HGua2Barps%></br>							
							<span class="gameTip"><small><%=Guard2HomeOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=hGua2Tip %></small></span>
						</td>
						<%end if%>
						
						<!--AWAY GUARD 2-->
  					<%if Guard2Away = "***GUA***" then%>
							<td class="text-left" style="color:white;background-color:#9a1400;width:50%;">
							<redIcon2><i class="fas fa-user-slash"></i></redIcon2></br>
							<small><%=Guard2AwayOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=gua2Tip%></small>
						</td>	
						<%elseif awayGua2TipDeadlinePassed then %>
						<td class="text-left  big" style="background-color:black;color: #a89a7a;font-weight:bold;">
							<span style="color:white;">GUA</span>&nbsp;<span style="font-weight:bold;color:yellowgreen;"><%=left(Guard2AwayNF,1) %>.&nbsp;<%=Guard2Away %></span><br>
						  <small><%=Guard2AwayOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=gua2Tip %></small>
						</td>
						<%else%>
						<td class="text-left">
							<span class="orangeText">GUA</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=sGuard2Away %>"><%=(left(Guard2AwayNF,1))%>.&nbsp;<%=Guard2Away %></a>&nbsp;<%=AGua2Barps%></br>
							<span class="gameTip"><small><%=Guard2AwayOpp%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%=gua2Tip%></small></span>
						</td>	
						<%end if%>				
					</tr>
				</table>
				<br>

			<%
				objRSMatchups.MoveNext
				loopcnt = loopcnt + 1
				awayCenTipDeadlinePassed  = false
				awayFor1TipDeadlinePassed = false
				awayFor2TipDeadlinePassed = false
				awayGua1TipDeadlinePassed = false
				awayGua2TipDeadlinePassed = false
				homeCenTipDeadlinePassed  = false
				homeFor1TipDeadlinePassed = false
				homeFor2TipDeadlinePassed = false	
				homeGua1TipDeadlinePassed = false
				homeGua2TipDeadlinePassed = false
				homePlayerCnt = 0
				awayPlayerCnt = 0 

				Wend
			%>
			<%elseif sAction <> "Submit Lineup" then%>
					<div class="row">
						<div class="col-xs-12">
							<div class="jumbotron" style="background-color:#9a1400;color:white;text-align:center;vertical-align:middle;font-weight:bold;">
								<i class="far fa-basketball-hoop"></i>&nbsp;NO MATCHUPS SET CURRENTLY!&nbsp;<i class="far fa-basketball-ball"></i>
							</div>
						</div>
					</div>
			<%end if%>	
			<%
				objRSMatchups.Close
				objRSBarps.Close
				objRSPics.Close  
				objRSBarps.Close 			
				
				Set objRSMatchups= Nothing
				Set objRSBarps   = Nothing
				Set objRSPics    = Nothing  
				Set objRSBarps   = Nothing  
				
			%>			
			</div>
		  <% if sAction = "Reforecast" or sAction = "ReforecastAdd" then %>
			<div id="VS" class="tab-pane fade in active">
			<%else%>
			<div id="VS" class="tab-pane fade">			
			<%end if%>
				<%
				  Set objRSHome				= Server.CreateObject("ADODB.RecordSet")
					Set objRSAway				= Server.CreateObject("ADODB.RecordSet")
					Set objRSAll				= Server.CreateObject("ADODB.RecordSet")
					Set objRSGameDay		= Server.CreateObject("ADODB.RecordSet")
					Set objRSPlayers    = Server.CreateObject("ADODB.RecordSet")
					Set objRSMyTeam     = Server.CreateObject("ADODB.RecordSet")
					Set objRSNoTeam     = Server.CreateObject("ADODB.RecordSet")
					Set	objRSHurt       = Server.CreateObject("ADODB.RecordSet")					 
					Set objRSLU         = Server.CreateObject("ADODB.RecordSet")	
				
					objRSHurt.Open "SELECT * from tblplayers where IR = true and ownerID =  " & ownerid & " order by firstname",objConn,3,3,1	
					
				%>
				<%
					objRSMyTeam.Open "SELECT * from tblplayers where ownerID =  " & ownerid & " order by firstname",objConn,3,3,1	
					objRSNoTeam.Open "SELECT * from tblplayers where ownerID = 0 order by firstname",objConn,3,3,1	
				%>
				<div class="bs-callout bs-callout-success">
					<h4><span style="font-weight:bold;">Re-Forecast Rules</span></h4>
					<ol>
						<li><strong>Add to Forecast</strong> - Add Free-Agents</li>
						<li>Select Player from Drop-Down List</li>
						<li>Click Re-Forecast Button</li>
						<li>Click Re-Forecast Only for Default Analysis</li>
					</ol>
				</div>
				</br>
				<form action="viewLineups.asp" name="frmMain" method="POST">				
				<table class="table table-custom table-responsive table-bordered table-condensed">
					<tr>
					 <th  colspan="2" style="color:black;font-size:12px !important;text-align:center">Re-Forecast Lineups</th>
					</tr>
					<tr style="background-color:white;color:black;">
						<td style="width:50%;">        
						<select class="form-control input-sm" name="addPlayerPID" >
						<option value="" selected>Add to Forecast</option>
							<%
								While Not objRSNoTeam.EOF
							%>
							<option value="<%=objRSNoTeam("pid")%>"><%=objRSNoTeam.Fields("firstName")%>&nbsp;<%=objRSNoTeam("lastName")%>&nbsp;|&nbsp;<%=objRSNoTeam("pos")%></option>
							<%
								objRSNoTeam.MoveNext
								Wend
							%>								
						</select>
						</td>
						<td>
							<button type="submit" value="ReforecastAdd" name="Action" class="btn btn-block btn-default">RE-FORECAST</button>
						</td>
					</tr>
				</table>
		 		</form>
				<%if objRSHurt.RecordCount > 0 then %>
				</br>
				<table class="table table-custom-black table-bordered table-condensed">
					<tr style="background-color:red;vertical-align:middle">
					<td>
						<table class="table table-striped table-responsive table-custom table-condensed">
						<td style="text-align:left;"><i class="fas fa-briefcase-medical red"></i>&nbsp;Your Injured Players are Omitted from Forecast. Visit the Player Profile Page by Clicking the link to set the Indicator to Off.</td>
						<%
						While Not objRSHurt.EOF
						%>
							<tr>
								<td><a class="blue" href="playerprofile.asp?pid=<%=objRSHurt.Fields("PID").Value %>"><%=objRSHurt.Fields("firstName").Value%>&nbsp;<%=objRSHurt.Fields("LastName").Value%></a></td>
							</tr>
						<%
						objRSHurt.MoveNext
						Wend
						%>
						</table>
					</td>
					</tr>
				</table>
				<%end if%>
				</br>
				<div class="bs-callout bs-callout-lineups">
					<h4><span style="font-weight:bold;">Quick Lineup</span></h4>
					<ol>
						<li><strong>Submit Quick Lineup - <span style="color:#468847;font-weight:bold;">Optimal Available Lineup!</span></strong></li>						
						<li>Only Allowed When 5 Players are Available</li>
						<li>Only Allowed Prior to Tip of First Game of the Day</li>
						<li><i class="fas fa-briefcase-medical red"></i>&nbsp;Players are Not Forecasted</li>
					</ol>
				</div>
				</br>
          <%
						objRSAll.Open "SELECT * FROM tblGameDeadLines where gameDay >= Date() order by gameDay", objConn
								
						While Not objRSAll.EOF
					
					    dgameday       = objRSAll("gameday")
							dgameDeadline  = objRSAll("gamedeadline")
							
					    objRSGameDay.Open "SELECT * "&_
																"FROM qryAllGames "&_
																"WHERE gameday = #"&dgameday&"# " & _
																"AND (qryAllGames.AwayTeamInd = " & ownerid & " or qryAllGames.HomeTeamInd = " & ownerid & ")", objConn,3,3,1										  
                        
							if objRSGameDay.Recordcount > 0 then						
								 if objRSGameDay("HomeTeamInd") = ownerid then
									OpponentName = objRSGameDay("AwayTeamShort")
									OpponentID = objRSGameDay("AwayTeamInd")
								 else 
									OpponentName = objRSGameDay("HomeTeamShort")
									OpponentID = objRSGameDay("HomeTeamInd")
								 end if
								else
								 OpponentName = "TBD"
								 OpponentID = 999
							end if   
							param_stagger = "STAGGER_WINDOW"
							objParams.Open  "SELECT * FROM tblParameterCtl WHERE param_name = '"&param_stagger&"'", objConn,3,3,1
							staggerPeriod = objParams.Fields("param_indicator").value
							objParams.Close
							'Response.Write "TABLE GAME DAY      = "&dgameday&".<br>"
							'Response.Write "TIP TIME            = "&dgameDeadline&".<br>"
							'Response.Write "DATE                = "&DATE()&".<br>"
							'Response.Write "STAGGER GAME TIME   = "&objRSGameDay.Fields("gameStaggerDeadline").Value&".<br>"
	
							lineupTimeChk = time() - 1/24

							'Response.Write "LINEUP TIME CHECK   = "&lineupTimeChk&".<br>"
							if(dgameday  = date() and objRSGameDay.Fields("gameDeadline").Value < lineupTimeChk) or (dgameday  = date() and staggerPeriod = true) then
								showButton = false
							else
								showButton = true 
							end if
							
							objRSGameDay.Close
							
							CenName       = false			
							For1Name      = false					
							For2Name      = false				
							Guard1Name    = false					
							Guard2Name    = false
							OppCenName    = false
							OppFor1Name   = false
							OppFor2Name   = false
							OppGuard1Name = false
							OppGuard2Name = false

							if omitPID > 0 and sAction = "Reforecast" or sAction = "ReforecastAdd" then 
								wRetcd = RE_Forecast_Lineup (dgameday,ownerid, CenName,CenBarps, CenPID, For1Name, For1Barps, For1PID, For2Name, For2Barps, For2Pid, Guard1Name, Guard1Barps, Gua1PID, Guard2Name, Guard2Barps, Gua2PID, omitPID)
							else
								wRetcd = Forecast_HurtLineup(dgameday,ownerid, CenName,CenBarps, CenPID, For1Name, For1Barps, For1PID, For2Name, For2Barps, For2Pid, Guard1Name, Guard1Barps, Gua1PID, Guard2Name, Guard2Barps, Gua2PID)
							end if
							
							wRetcd = Forecast_Lineup(dgameday, OpponentID, OppCenName, OppCenBarps, OppFor1Name, OppFor1Barps, OppFor2Name, OppFor2Barps, OppGuard1Name, OppGuard1Barps, OppGuard2Name, OppGuard2Barps)						
							
							if IsNull(CenBarps)       then CenBarps       = 0 end if
							if IsNull(For1Barps)      then For1Barps      = 0 end if
							if IsNull(For2Barps)      then For2Barps      = 0 end if
							if IsNull(Guard1Barps)    then Guard1Barps    = 0 end if
							if IsNull(Guard2Barps)    then Guard2Barps    = 0 end if
							if IsNull(OppCenBarps)    then OppCenBarps    = 0 end if
							if IsNull(OppFor1Barps)   then OppFor1Barps   = 0 end if
							if IsNull(OppFor2Barps)   then OppFor2Barps   = 0 end if
							if IsNull(OppGuard1Barps) then OppGuard1Barps = 0 end if
							
							if CenName     = false	  then showButton = false end if		
							if For1Name    = false	  then showButton = false end if				
							if For2Name    = false	  then showButton = false end if			
							if Guard1Name  = false	  then showButton = false end if				
							if Guard2Name  = false    then showButton = false end if
						 
						 
							'*****************************************************************
							'*** CHECK PIDS TO VERIFY THEY ARE ROSTERED AND NOT FREE-AGENTS
							'*****************************************************************
							
							Set objsStatus = Server.CreateObject("ADODB.RecordSet")								
							objsStatus.Open "SELECT * FROM TBLPLAYERS WHERE PID in ("&CenPID&","&For1PID&","&For2Pid&","&Gua1PID&","&Gua2PID&")  AND PlayerStatus = 'O' ", objConn,3,3,1
							
							if objsStatus.RecordCount = 5 then 
								sPlayerCntValid = true
							else
								sPlayerCntValid = false
							end if	
									
							'Response.Write "Number Rostered Players = "&objsStatus.RecordCount&".<br>"
							objsStatus.Close
							
							myTotal  = cDbl(CenBarps)+cDbl(For1Barps)+cDbl(For2Barps)+cDbl(Guard1Barps)+cDbl(Guard2Barps)
							OppTotal = cDbl(OppCenBarps)+cDbl(OppFor1Barps)+cDbl(OppFor2Barps)+cDbl(OppGuard1Barps)+cDbl(OppGuard2Barps)
							MyFav = (myTotal - OppTotal)
							OppFav = (OppTotal - myTotal)
				%>

					<form action="viewLineups.asp" method="POST" onSubmit="return functionLineup(this)" name="frmLineups"  language="JavaScript">
					<input type="hidden" name="var_ownerid"  value="<%= ownerid %>"/>
					<input type="hidden" name="wCBarps"  value="<%=round(CenBarps,2)%>"/>
					<input type="hidden" name="wF1Barps" value="<%=round(For1Barps,2)%>"/>
					<input type="hidden" name="wF2Barps" value="<%=round(For2Barps,2)%>"/>
					<input type="hidden" name="wG1Barps" value="<%=round(Guard1Barps,2)%>"/>
					<input type="hidden" name="wG2Barps" value="<%=round(Guard2Barps,2)%>"/>
					<input type="hidden" name="wCPID"    value="<%=CenPID%>"/>
					<input type="hidden" name="wF1Pid"   value="<%=For1PID%>"/>
					<input type="hidden" name="wF2Pid"   value="<%=For2Pid%>"/>
					<input type="hidden" name="wG1Pid"   value="<%=Gua1PID%>"/>
					<input type="hidden" name="wG2Pid"   value="<%=Gua2PID%>"/>	
					<input type="hidden" name="GameDays"  value="<%=objRSAll("gameday")%>"/>							
					<table class="table table-custom-black table-bordered table-condensed">
						<tr style="background-color: white;">
							<th colspan="2">Forecast for <%=objRSAll("gameday")%></th>
						</tr>
						<%
							myLine = cDbl(OppTotal) - cDbl(myTotal)
							vLine  = cDbl(myTotal)  - cDbl(OppTotal)						
						%>
						<tr style="background-color: white;">
							<td style="vertical-align:middle;width:50%;color:black"><span style="font-weight:bold;black">Line:</span>&nbsp;<%=round(myLine,1)%>&nbsp;<span style="font-weight:bold;color:black;">Proj:</span>&nbsp;<%= round(myTotal,2)%></td>
							<td style="vertical-align:middle;width:50%;color:black"><span style="text-left;color: #9a1400;font-weight: bold;" class="text-uppercase">[<%=OpponentName%>]</span>&nbsp;<span style="font-weight:bold;black">Line:&nbsp;</span><%=round(vLine,1)%>&nbsp;<span style="font-weight:bold;color:black;">Proj:</span>&nbsp;<%=round(oppTotal,2)%></td>
						</tr>
						<tr style="vertical-align:middle;background-color:white;">						
							<%if CenName = false then%>
								<td style="vertical-align:middle;width:50%;" class="opp">				
									<span class="blue"><i class="fas fa-user-slash"></i></span>
								</td>
							<%else%>
								<td style="vertical-align:middle;width:50%;" class="opp">
									<%if cDBL(CenBarps) > cDBL(OppCenBarps) then %>
										<span class="orangeText">CEN</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=cenPID %>"><%=left(CenName,11)%></a>&nbsp;<%=round(CenBarps,2)%>&nbsp;<i class="far fa-check green fa-sm"></i>
									<%else%>
										<span class="orangeText">CEN</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=cenPID %>"><%=left(CenName,11)%></a>&nbsp;<%=round(CenBarps,2)%>
									<%end if%>
								</td>
							<%end if%>
							
							
							<%if OppCenName = false then%>
								<td style="vertical-align:middle;text-align:left;" class="opp">
									<span class="blue"><i class="fas fa-user-slash"></i></span>
								</td>
							<%else%>
								<td style="vertical-align:middle;text-align:left;">
								<%if cDBL(OppCenBarps) > cDBL(CenBarps) then %>		
									<span class="orangeText">CEN</span>&nbsp;<span class="blue"><%=OppCenName%></span>&nbsp;<%=round(OppCenBarps,2)%>&nbsp;<i class="far fa-check green fa-sm"></i>							
								<%else%>
									<span class="orangeText">CEN</span>&nbsp;<span class="blue"><%=OppCenName%></span>&nbsp;<%=round(OppCenBarps,2)%>								
								<%end if%>
								</td>
							<%end if %>
						</tr> 
				
						<tr style="vertical-align:middle;background-color:white;">
							<%if For1Name = false then%>
								<td style="vertical-align:middle;text-align:left;">
									<span class="blue"><i class="fas fa-user-slash"></i></span>
								</td>
							<%else%>
								<%if Cdbl(For1Barps) > Cdbl(OppFor1Barps) then %>
									<td class="text-left"><span class="orangeText">FOR</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=for1PID %>"><%=left(For1Name,11)%></a>&nbsp;<%=round(For1Barps,2)%>&nbsp;<i class="far fa-check green fa-sm"></i></td>
								<%else%>
									<td class="text-left"><span class="orangeText">FOR</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=for1PID %>"><%=left(For1Name,11)%></a>&nbsp;<%=round(For1Barps,2)%></td>
								<%end if%>				
							<%end if %>
							<%if OppFor1Name = false then%>
								<td style="vertical-align:middle;text-align:left;">
									<span class="blue"><i class="fas fa-user-slash"></i></span>
								</td>
							<%else%>
								<%if Cdbl(OppFor1Barps) > Cdbl(For1Barps) then%>
									<td class="text-left"><span class="orangeText">FOR</span>&nbsp;<span class="blue"><%=OppFor1Name %></span>&nbsp;<%=round(OppFor1Barps,2)%>&nbsp;<i class="far fa-check green fa-sm"></i></td>
								<%else%>
									<td class="text-left"><span class="orangeText">FOR</span>&nbsp;<span class="blue"><%=OppFor1Name %></span>&nbsp;<%=round(OppFor1Barps,2)%></td>
								<%end if%>
							<%end if %>
						</tr>
						<tr style="vertical-align:middle;background-color:white;">
							<% if For2Name = false then %>
								<td style="vertical-align:middle;text-align:left;">
									<span class="blue"><i class="fas fa-user-slash"></i></span>
								</td>
							<%else%>
								<%if cDBL(For2Barps) > cDBL(OppFor2Barps) then%>
									<td class="text-left"><span class="orangeText">FOR</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=for2PID %>"><%=left(For2Name,11)%></a>&nbsp;<%=round(For2Barps,2)%>&nbsp;<i class="far fa-check green fa-sm"></i></td>
								<%else%>
									<td class="text-left"><span class="orangeText">FOR</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=for2PID %>"><%=left(For2Name,11)%></a>&nbsp;<%=round(For2Barps,2)%></td>
								<%end if%>									
							<%end if %>
								
							<% if OppFor2Name = false then %>
								<td style="vertical-align:middle;text-align:left;">
									<span class="blue"><i class="fas fa-user-slash"></i></span>
								</td>
							<%else%>
								<%if cDBL(OppFor2Barps) > cDBL(For2Barps) then%>
									<td class="text-left"><span class="orangeText">FOR</span>&nbsp;<span class="blue"><%=OppFor2Name %></span>&nbsp;<%=round(OppFor2Barps,2)%>&nbsp;<i class="far fa-check green fa-sm"></i></td>
								<%else%>
									<td class="text-left"><span class="orangeText">FOR</span>&nbsp;<span class="blue"><%=OppFor2Name %></span>&nbsp;<%=round(OppFor2Barps,2)%></td>
								<%end if%>
							<%end if %>
						</tr>				
						<tr style="vertical-align:middle;background-color:white;">
							<%if Guard1Name = false then%>
								<td style="vertical-align:middle;text-align:left;">
									<span class="blue"><i class="fas fa-user-slash"></i></span>
								</td>
							<%else%>
								<%if cDBL(Guard1Barps) > cDBL(OppGuard1Barps) then %>
									<td class="text-left"><span class="orangeText">GUA</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=gua1PID %>"><%=left(Guard1Name,11)%></a>&nbsp;<%=round(Guard1Barps,2)%>&nbsp;<i class="far fa-check green fa-sm"></i></td>
								<%else%>
									<td class="text-left"><span class="orangeText">GUA</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=gua1PID %>"><%=left(Guard1Name,11)%></a>&nbsp;<%=round(Guard1Barps,2)%></td>
								<%end if%>
							<%end if %>
							<%if OppGuard1Name = false then %>
								<td style="vertical-align:middle;text-align:left;">
									<span class="blue"><i class="fas fa-user-slash"></i></span>
								</td>
							<%else%>
								<%if cDBL(OppGuard1Barps) > cDBL(Guard1Barps) then %>
									<td class="text-left"><span class="orangeText">GUA</span>&nbsp;<span class="blue"><%=OppGuard1Name %></span>&nbsp;<%=round(OppGuard1Barps,2)%>&nbsp;<i class="far fa-check green fa-sm"></i></td>
								<%else%>
									<td class="text-left"><span class="orangeText">GUA</span>&nbsp;<span class="blue"><%=OppGuard1Name %></span>&nbsp;<%=round(OppGuard1Barps,2)%></td>
								<%end if%>
							<%end if %>
						</tr> 
						<tr style="vertical-align:middle;background-color:white;">
							<% if Guard2Name = false then %>
								<td style="vertical-align:middle;text-align:left;">
									<span class="blue"><i class="fas fa-user-slash"></i></span>
								</td>
							<%else%>
								<%if cDBL(Guard2Barps) > cDBL(OppGuard2Barps) then %>
									<td class="text-left"><span class="orangeText">GUA</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=gua2PID %>"><%=left(Guard2Name,11)%></a>&nbsp;<%=round(Guard2Barps,2)%>&nbsp;<i class="far fa-check green fa-sm"></i></td>							
								<%else%>
									<td class="text-left"><span class="orangeText">GUA</span>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=gua2PID %>"><%=left(Guard2Name,11)%></a>&nbsp;<%=round(Guard2Barps,2)%></td>							
								<%end if%>
							<%end if %>
							<% if OppGuard2Name = false then %>
								<td style="vertical-align:middle;text-align:left;">
									<span class="blue"><i class="fas fa-user-slash"></i></span>
								</td>
							<%else%>
								<%if cDBL(OppGuard2Barps) > cDBL(Guard2Barps) then%>
									<td class="text-left "><span class="orangeText">GUA</span>&nbsp;<span class="blue"><%=OppGuard2Name %></span>&nbsp;<%=round(OppGuard2Barps,2)%>&nbsp;<i class="far fa-check green fa-sm"></i></td>
								<%else%>
									<td class="text-left "><span class="orangeText">GUA</span>&nbsp;<span class="blue"><%=OppGuard2Name %></span>&nbsp;<%=round(OppGuard2Barps,2)%></td>
								<%end if%>
							<%end if %>
						</tr>					
						<% if showButton = true and sPlayerCntValid = true then %>
						<tr>
							<td colspan="5">					
								<button type="submit" id="idSubmitLineup"  value="Submit Lineup" name="Action" class="btn btn-block btn-default"><span class="glyphicon glyphicon-save"></span>&nbsp;Submit Quick Lineup for&nbsp;<%=objRSAll("gameday")%></button>
							</td>
						</tr>
						<%else%>
						<tr>
							<td colspan="5" style="text-align:center;vertical-align:middle;font-weight:bold;background-color: yellow;vertical-align: middle;"><i class="fa fa-user-lock red " aria-hidden="true" style="vertical-align: sub;"></i>&nbsp;<span style="font-weight:bold;background-color: yellow;vertical-align: middle;">Submit Lineup Via Dashboard!</span></td>
						</tr>
						<%end if%>
					</table>
				</form>
			<br>
			<%
			objRSAll.MoveNext
			Wend
			objRSAll.Close			
			%>
					<!--#include virtual="Common/missingLineups.inc"-->
					<!--#include virtual="Common/forecastLineups.inc"-->
					<!--#include virtual="Common/hurtPLayers.inc"-->							
			</div>
		</div>
	</div>
</div>
<%
  objConn.Close
	objRSHurt.Close
  Set objConn = Nothing
%>
<script>
$(document).ready(function(){
    $('[data-toggle="tooltip"]').tooltip(); 
});

$(document).ready(function() {
    $('#example').DataTable( {
        "ordering": false,
				"paging":  	true,
				"info":     false,
				"lengthMenu": [[10, 20, 30, -1], [10, 20, 30, "All"]]
    } );
} );
</script>
</body>
</html