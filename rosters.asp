<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""

	Dim ownerID, teamcnt, objRSteams, teamloop,sAction
	Dim objConn, objConn1,objRS,objRS1,objRS2,objRS3,objRS4,objRS5,objRS6

	Set objConn   = Server.CreateObject("ADODB.Connection")
	Set objRS     = Server.CreateObject("ADODB.RecordSet")
	Set objRSteams= Server.CreateObject("ADODB.RecordSet")
	Set objRS1    = Server.CreateObject("ADODB.RecordSet")
	Set objRS2    = Server.CreateObject("ADODB.RecordSet")
	Set objRS3    = Server.CreateObject("ADODB.RecordSet")
	Set objRS4    = Server.CreateObject("ADODB.RecordSet")
	Set objRS5    = Server.CreateObject("ADODB.RecordSet")
	Set objRS6    = Server.CreateObject("ADODB.RecordSet")
	Set objRS     = Server.CreateObject("ADODB.RecordSet")


	objConn.Open Application("lineupstest_ConnectionString")

	Dim strDatabaseType

  strDatabaseType = "Access"
	teamloop = 10
  objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"
	%>
	<!--#include virtual="Common/session.inc"-->
	<%	
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
<!--#include virtual="Common/noTrades.inc"-->
<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
.greenTrade {
	color:darkorange;
}
.item {
	background: #333;
	text-align: center;
	height: 120px !important;
}
td {
	font-size:11px;
	text-align:center;
}

.panel-override {
    color: darkorange;
    background-color:white;
    border-color: #01579B;
}
	
a.greenTrade {
    color: darkorange;
    text-decoration: underline;
    font-size: 11px;
    font-weight: 600;
}
		
.panel-title {
    color: yellowgreen;
		text-transform: none;	
		font-size: 14px  !important;
}

.panel-override {
  background-color: #ddd;
  border-color: black;
	border-width: 1px;
	background-color: #000000;
}

.bs-callout-success {
    border-left-color: #000000;
    padding: 10px;
    border-left-width: 4px;
    border-radius: 3px;
    background-color: white;
}

.panel-heading {
    background-image: none;
    background-color: #000000  !important;
    color: white;
    height: 30px;
    padding: 5px 5px;
}

.black {
    font-weight: 600;
}
.redText {
	color:#d6300e;
}
.btn-team {
    color: black;
    background-color: yellowgreen;
    border-color: #000000;
    font-weight: 600;
}
a.btn-team:hover{
     background: #d6300e;
}
.th{
border-radius:0px;
}
.badge {

	display: inline-block;
	min-width: 10px;
	padding: 3px 6px;    
	line-height: 1;
	color: white;
	text-align: center;
	white-space: nowrap;
	vertical-align: baseline;
	background-color: black;
	border-radius: 14px;
	border:white;
	border-style: double;
	color:yellowgreen;
	font-weight: 600 !important;
	border-width: thick;
}
</style>
</head>
<body>
<!--#include virtual="Common/headerMain.inc"-->
<form action="rosters.asp" method="POST" language="JavaScript" name="FrontPage_Form1" onSubmit="return FrontPage_Form1_Validator(this)">
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong><i class="fas fa-users"></i>&nbsp;Team Rosters</strong>
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
					<% if len(objRSteams.Fields("teamName")) >= 26 then %>
					<td style="width:50%;text-align:left;"><a class="blue" href="#<%=objRSteams.Fields("teamName").Value%>"><%=objRSteams.Fields("ShortName").Value%></a></td>
					<% else %>
					<td style="width:50%;text-align:left;"><a class="blue" href="#<%=objRSteams.Fields("teamName").Value%>"><%=objRSteams.Fields("teamName").Value%></a></td>
					<%end if%>
					<%
					objRSteams.MoveNext
					teamloop  = teamloop - 1
					teamcnt = teamcnt + 1
					%>
					<% if len(objRSteams.Fields("teamName")) >= 26 then %>
					<td style="width:50%;text-align:left;"><a class="blue" href="#<%=objRSteams.Fields("teamName").Value%>"><%=objRSteams.Fields("ShortName").Value%></a></td>
					<% else %>
					<td style="width:50%;text-align:left;"><a class="blue" href="#<%=objRSteams.Fields("teamName").Value%>"><%=objRSteams.Fields("teamName").Value%></a></td>
					<%end if%>					<%
					objRSteams.MoveNext
					teamloop  = teamloop - 1
					teamcnt = teamcnt + 1
					Wend
					%>
				</tr>
			</table>
    </div>
  </div>
</div>
</br>
<div class="container">
	<div class="row">
			<div class="col-md-12 col-sm-12 col-xs-12">
				<table style="width:100%;border-collapse: collapse;bordercolor:#111111;cellspacing:0;border:0">	
					<tr><td style="color:black;text-align:center;text-transform: uppercase;" class="big"	><strong><small>Players in <redText>Red</redText> are On the BLock!</small></strong></td></tr>
				</table>	
		</div>
	</div>
</div>
<br>
  <%
	objRSteams.Close
	teamloop = 10
	teamcnt = 0
%>
  <%
 While teamloop > 0
%>
  <%
	teamcnt = teamcnt + 1
	objRS.Open   "SELECT * FROM qryRosters WHERE OwnerID  =" & teamcnt & "  and  pos = 'CEN'  ", objConn,3,3,1
  objRS2.Open  "SELECT * FROM qryRosters WHERE OwnerID =" & teamcnt & "  and  pos = 'FOR'  ", objConn,3,3,1
  objRS3.Open  "SELECT * FROM qryRosters WHERE OwnerID =" & teamcnt & "  and  pos = 'F-C'  ", objConn,3,3,1
  objRS5.Open  "SELECT * FROM qryRosters WHERE OwnerID =" & teamcnt & "  and  pos = 'GUA'  ", objConn,3,3,1
  objRS6.Open  "SELECT * FROM qryRosters WHERE OwnerID =" & teamcnt & "  and  pos = 'G-F'  ", objConn,3,3,1
	objRSteams.Open "SELECT * FROM qryowners WHERE OwnerID = " & teamcnt & " ", objConn,3,3,1

 	dim w_max_count
	w_max_count = 0

	if objRS.Recordcount > w_max_count then
		w_max_count = objRS.Recordcount
	end if

	if objRS2.Recordcount > w_max_count then
		w_max_count = objRS2.Recordcount
	end if

	if objRS3.Recordcount > w_max_count then
		w_max_count = objRS3.Recordcount
	end if

	'if objRS4.Recordcount > w_max_count then
	'	w_max_count = objRS4.Recordcount
	'end if

	if objRS5.Recordcount > w_max_count then
		w_max_count = objRS5.Recordcount
	end if

	if objRS6.Recordcount > w_max_count then
		w_max_count = objRS6.Recordcount
	end if
	
%>

<div class="container">
  <div class="row">
    <div class="col-md-12 col-sm-12 col-xs-12">
      <div class="panel panel-override">
			<div id="<%= objRSteams.Fields("teamName").Value%>"></div>
       <!-- <div class="panel-heading">
           <h5 class="panel-title"><%= objRSteams.Fields("teamName").Value%>&nbsp;<small>(<%= objRSteams.Fields("ShortNAme").Value%>)</small></h5>
        </div>-->
          <table class="table table-striped table-custom table-bordered table-condensed">
					<%if objRSteams.Fields("OwnerID").value = ownerid then%>
						<tr>
							<td  colspan="5" valign="middle" align="center">
								<button class="btn  btn-team btn-block" disabled><%= objRSteams.Fields("teamName").Value%>&nbsp;<small>(<%= objRSteams.Fields("ShortNAme").Value%>)</small> <%= objRSteams.Fields("ActivePlayerCnt").Value%></button>
							</td>	
						</tr>	
					<%elseif myTradeInd = false or objRSteams.Fields("acceptTradeOffers").Value = false then%>
					  <tr>
							<td  colspan="5" valign="middle" align="center">
								<button class="btn  btn-team btn-block" disabled><span style="text-transform: uppercase;"><%= objRSteams.Fields("TeamName").value%></span> <redIcon><i class="fa fa-ban fa-1x"></i></redIcon> TRADE OFFERS</button>
							</td>	
						</tr>	
					<%else%>
						<tr>
              <td style="text-align:center"colspan="5"><button class="btn btn-block btn-team " value="Continue;<%= Trim(objRSteams.Fields("OwnerID").value) %>;<%= objRSteams.Fields("TeamName").value%>" name="Action" type="submit">Trade w/<%= objRSteams.Fields("TeamName").value%>&nbsp;<small><span style="color:white;text-transform:uppercase;">[<%= objRSteams.Fields("ShortNAme").Value%>]</span></small></button></td>
            </tr>						
					<%end if %>	
            <tr>
              <th style="text-align:center;width:20%;color:black;border-radius: 0px;"><strong>CEN</strong></th>
              <th style="text-align:center;width:20%;color:black;border-radius: 0px;"><strong>FOR</strong></th>
              <th style="text-align:center;width:20%;color:black;border-radius: 0px;"><strong>F-C</strong></th>
              <th style="text-align:center;width:20%;color:black;border-radius: 0px;"><strong>GUA</strong></th>
              <th style="text-align:center;width:20%;color:black;border-radius: 0px;"><strong>G-F</strong></th>
            </tr>
            <%

 While w_max_count > 0
 	'****************************'
 	'* Set up Rental Indicators *'
 	'****************************'

	if 	objRS.Fields("rentalPlayer").Value = 0 then
		cen_rent_ind = " "
	else
		cen_rent_ind = "   (R)"
	end if

	if objRS.Fields("injury").Value = true then
		cen_hurt = "<red>(I)</red>" 
	else
		cen_hurt = " "
	end if
	
	if objRS2.Fields("injury").Value = true then
		for_hurt = "<red>(I)</red>" 
	else
		for_hurt = " "
	end if
	
	if objRS3.Fields("injury").Value = true then
		for_hurt2 = "<red>(I)</red>" 
	else
		for_hurt2 = " "
	end if
	
	if objRS5.Fields("injury").Value = true then
		gua_hurt = "<red>(I)</red>" 
	else
		gua_hurt = " "
	end if
	
	if objRS6.Fields("injury").Value = true then
		gua_hurt2 = "<red>(I)</red>" 
	else
		gua_hurt2 = " "
	end if
	
	if 	objRS2.Fields("rentalPlayer").Value = 0 then
		for_rent_ind = " "
	else
		for_rent_ind = "   (R)"
	end if

	if 	objRS3.Fields("rentalPlayer").Value = 0 then
		forcen_rent_ind = " "
	else
		forcen_rent_ind = "   (R)"
	end if

	if 	objRS5.Fields("rentalPlayer").Value = 0 then
		guard_rent_ind = " "
	else
		guard_rent_ind = "   (R)"
	end if

	if 	objRS6.Fields("rentalPlayer").Value = 0 then
		guafor_rent_ind = " "
	else
		guafor_rent_ind = "   (R)"
	end if
%>
            <tr style="background-color:#eeeded;">
              <td style="background-color:#FFFFFF;text-align:center;width:20%">
								<table style="width:100%;border-collapse:collapse;bordercolor:#111111;cellspacing:0;border:0">								 
                  <tr>
									<% if objRS.Fields("ontheblock").Value = true then %>
                    <td><a class="red" href="playerprofile.asp?pid=<%=objRS.Fields("pid").Value %>" target="_self"><%= left(objRS.Fields("lastName").Value,8) + cen_rent_ind + cen_hurt%></a></td>
                  <%else %>
                    <td><a class="blue" href="playerprofile.asp?pid=<%=objRS.Fields("pid").Value %>" target="_self"><%= left(objRS.Fields("lastName").Value,8) + cen_rent_ind + cen_hurt%></a></td>
									<%end if %>									
									</tr>
                  <tr>
                    <td><span class="greenTrade"><%= left(objRS.Fields("teamShortName").Value,3)%></span></td>
                  </tr>
									<tr>
									<% if objRS.Fields("barps").Value > 0 then %>
										<td><span class="badge"><%=round(objRS.Fields("barps").Value,2) %></span></td>
									<%else%>
										<td><span class="badge">0</span></td>
									<%end if%>
									</tr>
								</table>
							</td>
              <td style="background-color:#FFFFFF;text-align:center;width:20%">
								<table style="width:100%;border-collapse:collapse;border:0">
									<tr>
									<% if objRS2.Fields("ontheblock").Value = true then %>
                    <td><a class="red" href="playerprofile.asp?pid=<%=objRS2.Fields("pid").Value %>" target="_self"><%= left(objRS2.Fields("lastName").Value,8) + for_rent_ind + for_hurt%></a></td>
                  <%else %>
                    <td><a class="blue" href="playerprofile.asp?pid=<%=objRS2.Fields("pid").Value %>" target="_self"><%= left(objRS2.Fields("lastName").Value,8) + for_rent_ind + for_hurt%></a></td>
									<%end if %>	
									</tr>
									<tr>
										<td><span class="greenTrade"><%= left(objRS2.Fields("teamShortName").Value,3)%></span></td>
									</tr>
									<tr>
									<% if objRS2.Fields("barps").Value > 0 then %>
										<td><span class="badge"><%=round(objRS2.Fields("barps").Value,2) %></span></td>
									<%else%>
										<td><span class="badge">0</span></td>
									<%end if%>
									</tr>
                </table>
							</td>
							<td style="background-color:#FFFFFF;text-align:center;width:20%">
								<table style="width:100%;border-collapse: collapse;bordercolor:#111111;cellspacing:0;border:0">
                  <tr>
									<% if objRS3.Fields("ontheblock").Value = true then %>
                    <td><a class="red" href="playerprofile.asp?pid=<%=objRS3.Fields("pid").Value %>" target="_self"><%= left(objRS3.Fields("lastName").Value,8) + forcen_rent_ind + for_hurt2%></a></td>
                  <%else %>
                    <td><a class="blue" href="playerprofile.asp?pid=<%=objRS3.Fields("pid").Value %>" target="_self"><%= left(objRS3.Fields("lastName").Value,8) + forcen_rent_ind + for_hurt2%></a></td>
									<%end if %>	
                  </tr>
                  <tr>
										<td><span class="greenTrade"><%= left(objRS3.Fields("teamShortName").Value,3)%></span></td>
                  </tr>
									<tr>
									<% if objRS3.Fields("barps").Value > 0 then %>
										<td><span class="badge"><%=round(objRS3.Fields("barps").Value,2) %></span></td>
									<%else%>
										<td><span class="badge">0</span></td>
									<%end if%>
									</tr>
									</table>
							</td>
              <td style="background-color:#FFFFFF;text-align:center;width:20%">
								<table style="width:100%;border-collapse: collapse;bordercolor:#111111;cellspacing:0;border:0">
                  <tr>
									<% if objRS5.Fields("ontheblock").Value = true then %>
                    <td><a class="red" href="playerprofile.asp?pid=<%=objRS5.Fields("pid").Value %>" target="_self"><%= left(objRS5.Fields("lastName").Value,8) + guard_rent_ind + gua_hurt%></a></td>
                  <%else %>
                    <td><a class="blue" href="playerprofile.asp?pid=<%=objRS5.Fields("pid").Value %>" target="_self"><%= left(objRS5.Fields("lastName").Value,8) + guard_rent_ind + gua_hurt%></a></td>
									<%end if %>	
                  </tr>
                  <tr>
                    <td><span class="greenTrade"><%= left(objRS5.Fields("teamShortName").Value,3)%></span></td>
									</tr>
									<tr>
									<% if objRS5.Fields("barps").Value > 0 then %>
										<td><span class="badge"><%=round(objRS5.Fields("barps").Value,2) %></span></td>
									<%else%>
										<td><span class="badge">0</span></td>
									<%end if%>
									</tr>                
                </table>
							</td>
              <td style="background-color:#FFFFFF;text-align:center;width:20%">
								<table style="width:100%;border-collapse: collapse;bordercolor:#111111;cellspacing:0;border:0">	
                  <tr>
									<% if objRS6.Fields("ontheblock").Value = true then %>
                    <td><a class="red" href="playerprofile.asp?pid=<%=objRS6.Fields("pid").Value %>" target="_self"><%= left(objRS6.Fields("lastName").Value,8) + guafor_rent_ind + gua_hurt2%></a></td>
                  <%else %>
                    <td><a class="blue" href="playerprofile.asp?pid=<%=objRS6.Fields("pid").Value %>" target="_self"><%= left(objRS6.Fields("lastName").Value,8) + guafor_rent_ind + gua_hurt2%></a></td>
									<%end if %>	
                  </tr>
                  <tr>
										<td><span class="greenTrade"><%= left(objRS6.Fields("teamShortName").Value,3)%></span></td>
                  </tr>
									<tr>
									<% if objRS6.Fields("barps").Value > 0 then %>
										<td><span class="badge"><%=round(objRS6.Fields("barps").Value,2) %></span></td>
									<%else%>
										<td><span class="badge">0</span></td>
									<%end if%>
									</tr> 
                </table>
							</td>
            </tr>

            <%
						objRS.MoveNext
						objRS1.MoveNext
						objRS2.MoveNext
						objRS3.MoveNext
						objRS4.MoveNext
						objRS5.MoveNext
						objRS6.MoveNext
						w_max_count = w_max_count - 1
							 Wend
						%>
         </table>		
		
      </div>
    </div>
  </div>
</div>
<center>
  <a class="blue" href="#top">Return to top</a>
</center>
				<br>	
<%
	objRS.Close
	objRS1.Close
	objRS2.Close
	objRS3.Close
	objRS4.Close
	objRS5.Close
	objRS6.Close
	objRSteams.Close

	objRS.MoveFirst
	objRS1.MoveFirst
	objRS2.MoveFirst
	objRS3.MoveFirst
	objRS4.MoveFirst
	objRS5.MoveFirst
	objRS6.MoveFirst
	teamloop  = teamloop - 1

    Wend
%>
<%
  	objRS.Close
  	objRS1.Close
  	objRS2.Close
  	objRS3.Close
  	objRS4.Close
  	objRS5.Close
  	objRS6.Close
	  objRSteams.Close

    objConn.Close


  	Set objRS = Nothing
  	Set objRS1 = Nothing
  	Set objRS2 = Nothing
  	Set objRS3 = Nothing
  	Set objRS4 = Nothing
  	Set objRS5 = Nothing
  	Set objRS6 = Nothing
  	Session.CodePage = Session("FP_OldCodePage")
  	Session.LCID     = Session("FP_OldLCID")
%>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<table style="width:100%;border-collapse: collapse;bordercolor:#111111;cellspacing:0;border:0">	
				<tr><td style="color:black;text-align:center;text-transform:uppercase;" class="big"><strong><small>Players in <redText>Red</redText> are On the BLock!</small></strong></td></tr>
			</table>	
		</div>
	</div>	
</div>
</br>
</form>
</body>
</html>