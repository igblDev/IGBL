<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%

On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""

	Dim objRSLineups,objRSTeamLogos,ownerid

	ownerid = session("ownerid")	
	
  if ownerid = "" then
    	GetAnyParameter "var_ownerid", ownerid
	end if
	
	if ownerid = "" then
		Response.Redirect("timeout.asp")
	end if	
	
	Set objConn             = Server.CreateObject("ADODB.Connection")
	Set objRSLineups        = Server.CreateObject("ADODB.RecordSet")
	Set objRSMissingLineups = Server.CreateObject("ADODB.RecordSet")
	Set objNextRun          = Server.CreateObject("ADODB.RecordSet")
	Set objRSLineupsPics    = Server.CreateObject("ADODB.RecordSet")
	Set objRSTeamLogos      = Server.CreateObject("ADODB.RecordSet")
	Set objRSBarps          = Server.CreateObject("ADODB.RecordSet")
	
	objConn.Open Application("lineupstest_ConnectionString")


	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"
				  
				  

%>
<!DOCTYPE HTML>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=no">
<meta name="description" content="">
<meta name="author" content="">
<title>IGBL 2016-2017</title>
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/app.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>

black{
	color: black;
}
yellow{
	color: yellow;
}
small {
	font-size:12px;
}
td {
    vertical-align: middle;
}
.badge2 {
    display: inline-block;
    min-width: 10px;
    padding: 3px 7px;
    font-size: 12px;
    font-weight: 700;
    line-height: 1;
    color: #fff;
    text-align: center;
    white-space: nowrap;
    vertical-align: baseline;
    background-color: black;
    border-radius: 10px;
}
.badge3 {
    display: inline-block;
    min-width: 10px;
    padding: 3px 7px;
    font-size: 12px;
    font-weight: 700;
    line-height: 1;
    color: #fff;
    text-align: center;
    white-space: nowrap;
    vertical-align: baseline;
    background-color: #354478;
    border-radius: 10px;
}
</style>
</head>
<body>
<% 
	
	objRSLineups.Open "SELECT * FROM qry_matchupLineups", objConn,3,3,1
	gameday = objRSLineups.Fields("gameday").Value
	ownerid = session("ownerid")
	dim loopcnt 
	loopcnt = 1
	

	%>
<form>
<!--#include virtual="Common/headerMain.inc"-->

	<% if objRSLineups.RecordCount > 0 then %>
  <% 
	
	While Not objRSLineups.EOF
	
	hometeam      = objRSLineups.Fields("HomeTeam").Value
	hometeamshort = objRSLineups.Fields("HomeTeamShort").Value
	homeowner     = objRSLineups.Fields("HomeOwner").Value
	hometeampen   = objRSLineups.Fields("HomeTeamPen").Value
	awayteam      = objRSLineups.Fields("VisitingTeam").Value
	awayteamshort = objRSLineups.Fields("VisitingTeamShort").Value
	awayowner     = objRSLineups.Fields("VisitingOwner").Value
	awayteampen   = objRSLineups.Fields("VisitingTeamPen").Value
	
	objRSTeamLogos.Open "Select TeamLogo  from TBLOWNERS where  " & homeowner  & " = ownerID " , objConn,3,3,1
	homeLogo = objRSTeamLogos.Fields("TeamLogo").Value
	objRSTeamLogos.Close		
	
	objRSTeamLogos.Open "Select TeamLogo  from TBLOWNERS where  " & awayowner  & " = ownerID " , objConn,3,3,1
	awayLogo = objRSTeamLogos.Fields("TeamLogo").Value
	objRSTeamLogos.Close		
	
	homerec  = objRSLineups.Fields("HomeRecord").Value
	awayrec  = objRSLineups.Fields("VisRecord").Value
	homeStamp= objRSLineups.Fields("HomeStamp").Value
	awayStamp= objRSLineups.Fields("Timestamp").Value
	'******************************
	'** Retrieving Starters PID  **
	'******************************
	sCenterAway   = objRSLineups.Fields("sCenter").Value
	sForward1Away = objRSLineups.Fields("sForward").Value
	sForward2Away = objRSLineups.Fields("sForward2").Value
	sGuard1Away   = objRSLineups.Fields("sGuard").Value
	sGuard2Away   = objRSLineups.Fields("sGuard2").Value	
	sCenterHome   = objRSLineups.Fields("HomeC").Value
	sForward1Home = objRSLineups.Fields("HomeF").Value
	sForward2Home = objRSLineups.Fields("HomeF2").Value
	sGuard1Home   = objRSLineups.Fields("HomeG").Value
	sGuard2Home   = objRSLineups.Fields("HomeG2").Value
	lineupTimeChk = time() - 1/24	
	
	cTip          = objRSLineups.Fields("sCenterTip").Value
	for1Tip       = objRSLineups.Fields("sForwardTip").Value
	for2Tip       = objRSLineups.Fields("sForwardTip2").Value
	gua1Tip       = objRSLineups.Fields("sGuardTip").Value
	gua2Tip       = objRSLineups.Fields("sGuardTip2").Value
	hCenTip       = objRSLineups.Fields("HomeCTip").Value
	hForTip       = objRSLineups.Fields("HomeFTip").Value
	hFor2Tip      = objRSLineups.Fields("HomeF2Tip").Value
	hGuaTip       = objRSLineups.Fields("HomeGTip").Value
	hGua2Tip      = objRSLineups.Fields("HomeG2Tip").Value 

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
	
	objRSLineupsPics.Open "Select image,lastName  from tblPlayers where  " & sCenterAway  & " = PID " , objConn,3,3,1
	ACenPic     = objRSLineupsPics.Fields("image").Value
	CenterAway  = objRSLineupsPics.Fields("lastName").Value
	objRSLineupsPics.Close  
	
	objRSBarps.Open "SELECT barps FROM qry_tblbarps where  " & sCenterAway  & " = PID " , objConn,3,3,1
	ACenBarps   = objRSBarps.Fields("barps").Value
	objRSBarps.Close
	
	'Response.Write "QUERY BARPS PID: " & sCenterAway  
	'Response.Write "LINEUPS PID: " & ACenBarps  

	objRSLineupsPics.Open "Select image,lastName   from tblplayers where " & sForward1Away & " = PID" , objConn,3,3,1
	AFor1Pic    = objRSLineupsPics.Fields("image").Value
	Forward1Away= objRSLineupsPics.Fields("lastName").Value

	objRSLineupsPics.Close
	
	objRSBarps.Open "SELECT barps FROM qry_tblbarps where  " & sForward1Away & " = PID " , objConn,3,3,1
	AFor1Barps  = objRSBarps.Fields("barps").Value
	objRSBarps.Close


	objRSLineupsPics.Open "Select image,lastName   from tblplayers where " & sForward2Away &" = PID" , objConn,3,3,1
	AFor2Pic    = objRSLineupsPics.Fields("image").Value
	Forward2Away= objRSLineupsPics.Fields("lastName").Value

	objRSLineupsPics.Close 

	objRSBarps.Open "SELECT barps FROM qry_tblbarps where  " & sForward2Away & " = PID " , objConn,3,3,1
	AFor2Barps  = objRSBarps.Fields("barps").Value
	objRSBarps.Close
	
	objRSLineupsPics.Open "Select image,lastName   from tblplayers where  " & sGuard1Away & " = PID" , objConn,3,3,1
	AGua1Pic    = objRSLineupsPics.Fields("image").Value
	Guard1Away  = objRSLineupsPics.Fields("lastName").Value

	objRSLineupsPics.Close 

	objRSBarps.Open "SELECT barps FROM qry_tblbarps where  " & sGuard1Away & " = PID " , objConn,3,3,1
	AGua1Barps  = objRSBarps.Fields("barps").Value
	objRSBarps.Close

	objRSLineupsPics.Open "Select image,lastName   from tblplayers where " & sGuard2Away & "  = PID" , objConn,3,3,1
	AGua2Pic    = objRSLineupsPics.Fields("image").Value
	Guard2Away  = objRSLineupsPics.Fields("lastName").Value

	objRSLineupsPics.Close

	objRSBarps.Open "SELECT barps FROM qry_tblbarps where  " & sGuard2Away & " = PID " , objConn,3,3,1
	AGua2Barps  = objRSBarps.Fields("barps").Value
	objRSBarps.Close
	
	
	'******************************************************************************************************
	'**** RETRIEVING THE BARPS AD THE IMAGES FOR THE STARTING PLAYERS ON THE HOME TEAM
	'******************************************************************************************************
	
	objRSLineupsPics.Open "Select image,lastName from tblplayers where " & sCenterHome & "  = PID" , objConn,3,3,1
	HCenPic = objRSLineupsPics.Fields("image").Value
	CenterHomeN = objRSLineupsPics.Fields("lastName").Value
	'Response.Write "LINEUPS PID: " & objRSLineupsPics.Fields("lastName").Value

	objRSLineupsPics.Close
	
	objRSBarps.Open "SELECT barps FROM qry_tblbarps where  " & sCenterHome & " = PID " , objConn,3,3,1
	HCenBarps  = objRSBarps.Fields("barps").Value

	objRSBarps.Close
	
	objRSLineupsPics.Open "Select image,lastName from tblplayers where " & sForward1Home & " = PID" , objConn,3,3,1
	HFor1Pic = objRSLineupsPics.Fields("image").Value
	Forward1HomeN = objRSLineupsPics.Fields("lastName").Value

	objRSLineupsPics.Close

	objRSBarps.Open "SELECT barps FROM qry_tblbarps where  " & sForward1Home & " = PID " , objConn,3,3,1
	HFor1Barps  = objRSBarps.Fields("barps").Value
	
	objRSBarps.Close

	objRSLineupsPics.Open "Select image,lastName from tblplayers where " & sForward2Home & " = PID" , objConn,3,3,1
	HFor2Pic = objRSLineupsPics.Fields("image").Value
	Forward2HomeN = objRSLineupsPics.Fields("lastName").Value

	objRSLineupsPics.Close	

	objRSBarps.Open "SELECT barps FROM qry_tblbarps where  " & sForward2Home & " = PID " , objConn,3,3,1
	HFor2Barps = objRSBarps.Fields("barps").Value
	objRSBarps.Close
	
	objRSLineupsPics.Open "Select image,lastName from tblplayers where " & sGuard1Home & " = PID" , objConn,3,3,1
	HGua1Pic = objRSLineupsPics.Fields("image").Value
	Guard1HomeN = objRSLineupsPics.Fields("lastName").Value

	objRSLineupsPics.Close	

	objRSBarps.Open "SELECT barps FROM qry_tblbarps where  " & sGuard1Home & " = PID " , objConn,3,3,1
	HGua1Barps = objRSBarps.Fields("barps").Value
	objRSBarps.Close

	objRSLineupsPics.Open "Select image,lastName from tblplayers where " & sGuard2Home & " = PID" , objConn,3,3,1
	HGua2Pic = objRSLineupsPics.Fields("image").Value
	Guard2HomeN = objRSLineupsPics.Fields("lastName").Value

	objRSLineupsPics.Close

	objRSBarps.Open "SELECT barps FROM qry_tblbarps where  " & sGuard2Home & " = PID " , objConn,3,3,1
	HGua2Barps = objRSBarps.Fields("barps").Value
	objRSBarps.Close
	
	htBarps = HGua1Barps  + HGua2Barps + HFor1Barps  + HFor2Barps + HCenBarps
	atBarps = AGua1Barps  + AGua2Barps + AFor1Barps  + AFor2Barps + ACenBarps		
	
	%>
<div class="container">
    <div class="row">
			<div class="col-md-12 col-sm-12 col-xs-12">
				<div class="panel panel-primary">
								<div class="panel-heading clearfix">
					<h4 class="panel-title"><i class="fa fa-dribbble" aria-hidden="true"></i>&nbsp;Match-up #<%=loopcnt %>&nbsp;<small><yellow>for</small></yellow> <%=gameday %>&nbsp;<div class="pull-right"><small><yellow>all times cst</yellow></small></div></h4>
				</div>
        <table class="table table-condensed table-bordered">
          <tr  style="background:#821111; color:white" align="center">
            <% if awayteampen = true Then %>
            	<td  style="background-color:red"><%=awayteam%><br><%=awayrec%></td>
            <% else %> 
            	<td style="width:50%"><%=awayteam%><br><%=awayrec%></td>
            <%end if %>
            
						<% if hometeampen = true Then %>
							<td  style="background-color:red"><%=hometeam%><br><%=homerec%></td>          
						<% else %> 
							<td  style="width:50%"><%=hometeam%><br><%=homerec%></td>
            <%end if %>
					</tr>
          <tr  style="background-color:white" align="center">
						<% if awayteampen = true Then %>
							<td style="width:50%" align="center"><img style="max-height:100px; max-width:100px" class="img-responsive" border="0" src="images/penalty.png"></td>
						<% else %>
	            <td style="width:50%" align="center"><img style="max-height:100px; max-width:100px" class="img-responsive" border="0" src="<%=awayLogo%>"></td>
						<%end if %>

						<% if hometeampen = true Then %>
						  <td style="width:50%" align="center"><img style="max-height:100px; max-width:100px" class="img-responsive" border="0" src="images/penalty.png"></td>
						<% else %>
						  <td style="width:50%" align="center"><img style="max-height:100px; max-width:100px" class="img-responsive" border="0" src="<%=homeLogo%>"></td>
						<%end if %>
          </tr>
					<tr align="center">
            <td style="background-color:gray;color:white" colspan="2">Score Projections</td>
          </tr>
          <tr>
            <td align="center"><span class="badge"><%=atBarps%></span></td>
            <td align="center"><span class="badge"><%=htBarps%></span></td>
          </tr> 
          <tr  align="center">
            <td style="background-color:#468847;color:white" colspan="4">Starting Centers</td>
          </tr>
					<tr class="active">
            <td  align="center"><small><%=CenterAway %> | <%=cTip%><br><span class="badge3"><%=ACenBarps%></span></small></td>
            <td  align="center"><small><%=CenterHomeN %> | <%=hCenTip%><br><span class="badge2"><%=HCenBarps%></span></small></td>
					</tr>
					<tr>
						<td style="width:50%" align="center"><img class="img-responsive img-circle" width="30" height="40" border="0" src="<%=ACenPic%>"></td>
						<td style="width:50%" align="center"><img class="img-responsive img-circle" width="30" height="40" border="0" src="<%=HCenPic%>"></td>
					</tr>
          <tr align="center">
            <td style="background-color:#468847;color:white" colspan="4">Starting Forwards</td>
          </tr>
						<tr class="active">
						<td align="center"><small><%=Forward1Away %> | <%=for1Tip%><br><span class="badge3"><%=AFor1Barps%></span></small></td></small></td>
            <td align="center"><small><%=Forward1HomeN %> | <%=hForTip%><br><span class="badge2"><%=HFor1Barps%></span></small></td>
					</tr>
          <tr>
            <td style="width:50%" align="center"><img class="img-responsive img-circle" width="30" height="40" border="0" src="<%=AFor1Pic%>"></td>
						<td style="width:50%" align="center"><img class="img-responsive img-circle" width="30" height="40" border="0" src="<%=HFor1Pic%>"></td>
					</tr>
						<tr class="active">
            <td align="center"><small><%=Forward2Away %> | <%=for2Tip%><br><span class="badge3"><%=AFor2Barps%></span></small></td></small></td>
            <td align="center"><small><%=Forward2HomeN %> | <%=hFor2Tip%><br><span class="badge2"><%=HFor2Barps%></span></small></td>
					</tr>
          <tr>
            <td style="width:50%" align="center"><img class="img-responsive img-circle" width="30" height="40" border="0" src="<%=AFor2Pic%>"></td>
						<td style="width:50%" align="center"><img class="img-responsive img-circle" width="30" height="40" border="0" src="<%=HFor2Pic%>"></td>
          </tr>
          <tr  align="center">
            <td style="background-color:#468847;color:white" colspan="4">Starting Guards</td>
          </tr>
					<tr class="active">
            <td align="center"><small><%=Guard1Away %> | <%=gua1Tip%><br><span class="badge3"><%=AGua1Barps%></span></small></td></small></td>
            <td align="center"><small><%=Guard1HomeN%> | <%=hGuaTip%><br><span class="badge2"><%=HGua1Barps%></span></small></td>
					</tr>
          <tr>
            <td style="width:50%" align="center"><img class="img-responsive img-circle" width="30" height="40" border="0" src="<%=AGua1Pic%>"></td>
						<td style="width:50%" align="center"><img class="img-responsive img-circle" width="30" height="40" border="0" src="<%=HGua1Pic%>"></td>
          </tr>
					 <tr class="active">
            <td align="center"><small><%=Guard2Away %> | <%=gua2Tip%><br><span class="badge3"><%=AGua2Barps %></span></small></td></small></td>
            <td align="center"><small><%=Guard2HomeN %> | <%=hGua2Tip %><br><span class="badge2"><%=HGua2Barps%></span></small></td>
					</tr>
          <tr>
            <td align="center"><img class="img-responsive img-circle" width="30" height="40" border="0" src="<%=AGua2Pic%>"></td>
						<td align="center"><img class="img-responsive img-circle" width="30" height="40" border="0" src="<%=HGua2Pic%>"></td>
						</tr>
				</table>
      </div>
    </div>
  </div>
</div>
<%
	objRSLineups.MoveNext
	loopcnt = loopcnt + 1
	Wend
%>
<%else%>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-info">
				<strong><div style="text-align:center;vertical-align:middle"><i class="fa fa-exclamation-triangle fa-lg" aria-hidden="true"></i>&nbsp;NO MATCHUPS SET FOR TODAY!</div></strong>
			</div>
		</div>
	</div>
</div>
<%end if%>
</form>
<% 
objRSLineups.Close
objRSLineups.Close
objConn.Close
Set objConn = Nothing
 %>
</body>
</html>
