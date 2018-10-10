<form action="dashboard.asp" method="POST" language="JavaScript" name="FrontPage_Form1">
<%
	Set objConn        = Server.CreateObject("ADODB.Connection")
	Set objRS          = Server.CreateObject("ADODB.RecordSet")
	Set objRSteams     = Server.CreateObject("ADODB.RecordSet")
	Set objRS1         = Server.CreateObject("ADODB.RecordSet")
	Set objRS          = Server.CreateObject("ADODB.RecordSet")
	Set objRSCenters   = Server.CreateObject("ADODB.RecordSet")
	Set objRSForwards  = Server.CreateObject("ADODB.RecordSet")
	Set objRSGuards    = Server.CreateObject("ADODB.RecordSet") 

	objConn.Open Application("lineupstest_ConnectionString")

	teamloop = 10
	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"

	GetAnyParameter "Action", sAction	

	if sAction > "" then
		TEAM_Split        = Split(Request.Form("Action"), ";")
		sAction           = TEAM_Split(0) 'SAction
		tradepartner      = TEAM_Split(1) 'Team ID
		sTeam             = TEAM_Split(1) 'Team			
	end if	
		
	select case sAction

	case "Continue"
		sURL = "tradeanalyzer.asp"
		conAction = "Continue"
		AddLinkParameter "Action", conAction, sURL
		AddLinkParameter "cmbTeam", tradepartner, sURL
		Response.Redirect sURL		
	end select
%>
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Team Rosters</strong>
			</div>
		</div>
	</div>
<%
	teamloop = 10
	teamcnt = 0
%>
<%
	While teamloop > 0
%>
<%
	teamcnt = teamcnt + 1
	objRS.Open          "SELECT * FROM qryRosters WHERE OwnerID =" & teamcnt & "", objConn,3,3,1
	objRSteams.Open     "SELECT * FROM qryowners  WHERE OwnerID =" & teamcnt & "", objConn,3,3,1
							
	objRSCenters.Open   "SELECT * FROM tblPlayers WHERE (tblPlayers.POS  = 'CEN' or tblPlayers.POS  = 'F-C') and tblPlayers.OwnerID=" & teamcnt & "",objConn,3,3,1
	objRSForwards.Open  "SELECT * FROM tblPlayers WHERE (tblPlayers.POS  = 'FOR' or tblPlayers.POS  = 'F-C' or tblPlayers.POS   = 'G-F') and tblPlayers.OwnerID=" & teamcnt & "",objConn,3,3,1
	objRSGuards.Open  	"SELECT * FROM tblPlayers WHERE (tblPlayers.POS  = 'GUA' or tblPlayers.POS  = 'G-F') and tblPlayers.OwnerID=" & teamcnt & "",objConn,3,3,1

	centerCntMe  = objRSCenters.RecordCount
	forwardCntMe = objRSForwards.RecordCount
	guardCntME   = objRSGuards.RecordCount

	objRSCenters.Close  	
	objRSForwards.Close  
	objRSGuards.Close 

	dim w_max_count
	w_max_count = 0

	if objRS.Recordcount > w_max_count then
		w_max_count = objRS.Recordcount
	end if
%>

	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="panel panel-override">
			<div id="<%= objRSteams.Fields("teamName").Value%>"></div>
			<table class="table table-custom-black table-bordered table-condensed">
			
			<!--Verify Owner Id and Check if Owner is Accepting Trade Offers-->
			<%if objRSteams.Fields("OwnerID").value = ownerid then%>
				<tr>
					<td  colspan="4" valign="middle" align="center">
						<button class="btn  btn-team btn-block" disabled><%= objRSteams.Fields("teamName").Value%>&nbsp;<small>(<%= objRSteams.Fields("ShortName").Value%>)</small><span style="color:white;font-weight:bold;" class="pull-right"><%= FormatCurrency(objRSteams.Fields("waiverBal").Value)%></span></button>
					</td>	
				</tr>	
			<%elseif myTradeInd = false or objRSteams.Fields("acceptTradeOffers").Value = false then%>
				<tr>
					<td  colspan="4" valign="middle" align="center">
						<button class="btn  btn-team btn-block" disabled><span style="text-transform: uppercase;"><%= objRSteams.Fields("TeamName").value%></span> <redIcon><i class="fa fa-ban fa-1x"></i></redIcon> TRADE OFFERS<span style="color:white;font-weight:bold;" class="pull-right"><%= FormatCurrency(objRSteams.Fields("waiverBal").Value)%></span></button>
					</td>	
				</tr>	
			<%else%>
				<tr>
					<td style="text-align:center"colspan="4"><button class="btn btn-block btn-team " value="Continue;<%= Trim(objRSteams.Fields("OwnerID").value) %>;<%= objRSteams.Fields("TeamName").value%>" name="Action" type="submit">Trade w/<%= objRSteams.Fields("TeamName").value%>&nbsp;<small>(<%= objRSteams.Fields("ShortNAme").Value%>)</small><span style="color:white;font-weight:bold;" class="pull-right"><%= FormatCurrency(objRSteams.Fields("waiverBal").Value)%></span></button></td>
				</tr>						
			<%end if%>	
				<tr style="text-align:center";>
					<th style="color:black" colspan="4">Roster Configuration</th>
				</tr>
				<tr style="text-align:center;">
					<td colspan="4">
						<table class="table table-striped table-bordered table-condensed table-responsive">
							<tr>
								<td class ="big" style="background: yellowgreen;color:black;font-weight:bold;width:12%">CEN</td>
								<td class ="big" style="background-color:white" width="12%"><%=centerCntMe%></td>
								<td class ="big" style="background: yellowgreen;color:black;font-weight:bold;width:12%">FOR</td>
								<td class ="big" style="background-color:white" width="12%"><%=forwardCntME%></td>
								<td class ="big" style="background: yellowgreen;color:black;font-weight:bold;width:12%">GUA</td>
								<td class ="big" style="background-color:white" width="12%"><%=guardCntME%></td>
								<td class ="big" style="background: yellowgreen;color:black;font-weight:bold;width:12%">TOT</td>
								<td class ="big" style="background-color:white" width="12%">
								<%if myPlayerCnt < 14 then %>
								<span style="color:black;background-color:#FFEB3B;font-weight:bold;"><%=myPlayerCnt%></span>
								<%else %>
								<span style="color:black;"><%=myPlayerCnt%></span>
								<%end if %>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>					
					<th style="width:50%;text-align:left;"><span>Player - Team | Pos</span></th>
					<th style="width:20%;"><span>AVG</span></th>					
					<th style="width:20%;"><span>L/5</span></th>
					<th style="width:10%;"><span>OTB</span></th>
				</tr>
	<%
		While w_max_count > 0
		'****************************'
		'* Set up Rental Indicators *'
		'****************************'

		if 	objRS.Fields("rentalPlayer").Value = 0 then
			cen_rent_ind = " "
		else
			cen_rent_ind = true
		end if

		if objRS.Fields("IR").Value = true then
			cen_hurt = true 
		else
			cen_hurt = " "
		end if

		myPlayerCnt = objRS.recordcount
	%>
				<tr style="background-color:white;">					
					<td style="text-align: left; vertical-align: middle;">
					<a class="blue" href="playerprofile.asp?pid=<%=objRS.Fields("pid").Value %>" target="_self"><%=left(objRS.Fields("firstName").Value,1)%>.&nbsp;<%=left(objRS.Fields("lastName").Value,10) %></a>,&nbsp;
					<%if cen_rent_ind = true then %>
						<i class="fa fa-registered blue" aria-hidden="true"></i>
					<%end if%>
					<%if cen_hurt = true then %>
						<i class="fas fa-briefcase-medical red"></i>
					<%end if%>						
					<span class="greenTrade"><%=objRS.Fields("TeamshortName").Value%></span>&nbsp;|&nbsp;<span class="orange"><%=objRS.Fields("POS").Value%></span>
					</td>
					<td style="vertical-align: middle;text-align:center;"><%= round(objRS.Fields ("barps").Value,2) %></td>
					<td class="big" style="vertical-align:middle;text-align:center"><%= round(objRS.Fields ("l5barps").Value,2)%></td>		
				<% if objRS.Fields("ontheblock").Value = true then %>
					<td style="vertical-align: middle;text-align:center;"><span><i class="far fa-check"></i></span></td>
				<%else%>
					<td></td>
				<%end if%>
				</tr>

	<%
	objRS.MoveNext
	w_max_count = w_max_count - 1
	cen_hurt = " "
	cen_rent_ind = " "
	Wend
	%>
			</table>				
			</div>
		</div>
	</div>
<%
	objRS.Close
	objRSteams.Close

	objRS.MoveFirst
	teamloop  = teamloop - 1

	Wend
%>
</br>
</form>	