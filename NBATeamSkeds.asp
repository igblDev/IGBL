<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%

On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""

	Dim objConn,objLineups,objRSAllDates
	Dim strSQL,I,ownerid,sAction
	Dim errorCode, errorDesc, txnteamname, mailTxt
	GetAnyParameter "Action", sAction

	Set objConn       = Server.CreateObject("ADODB.Connection")
	Set objRSTMSkeds  = Server.CreateObject("ADODB.RecordSet")
	Set objRSAllDates = Server.CreateObject("ADODB.RecordSet")
	
	objConn.Open Application("lineupstest_ConnectionString")

  objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"

%>
	<!--#include virtual="Common/session.inc"-->
<%
  GetAnyParameter "Action", sAction
	GetAnyParameter "var_sPid", ppPID	
%>
<!--#include virtual="Common/functions.inc"-->
<!DOCTYPE HTML>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=no">
<meta name="description" content="">
<meta name="author" content="">
<title>IGBL 2017-2018</title>
<!-- Bootstrap core CSS -->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Custom styles for this template -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style type="text/css">
.panel-override {
  background-color:white;
  border-color:#354478;
	color:black
}
td {
	font-size:11px;
}
</style>
<script>
$(document).ready(function(){
    $('[data-toggle="tooltip"]').tooltip();   
});
</script>
</head>
<body>
<script type="text/javascript">
function processTeams(theForm) {
	
	var teamCnt = 2;
	var teamsSelected = 0;
	var timeEntered 
	timeEntered = theForm.newTime.value
	
	for(var i=0; i < theForm.chkTeamSked.length; i++){
		if(theForm.chkTeamSked[i].checked) {
		teamsSelected +=1;
		}
	}

	if(teamsSelected < 2  ) {
		alert("Two Teams Required for Update!" ); 
		return (false);
	}else if(teamsSelected > 2  ) {
		alert("Only Two Teams Can be Selected for Update!" ); 
		return (false); 
	}

	if(!timeEntered) {
		alert("New Time Required for Update!" ); 
		return (false);
	}

	return (true);
}	

function processNewTeams(theForm) {
	
	var teamCnt = 2;
	var teamsSelected = 0;
	var timeEntered 
	timeEntered = theForm.newAddGameTime.value
	
	for(var i=0; i < theForm.chkTeamSkedAdd.length; i++){
		if(theForm.chkTeamSkedAdd[i].checked) {
		teamsSelected +=1;
		}
	}

	if(teamsSelected < 2  ) {
		alert("Two Teams Required for Update!" ); 
		return (false);
	}else if(teamsSelected > 2  ) {
		alert("Only Two Teams Can be Selected for Update!" ); 
		return (false); 
	}

	if(!timeEntered) {
		alert("New Time Required for Update!" ); 
		return (false);
	}

	return (true);
}	

</script>
<!--#include virtual="Common/headerMain.inc"-->
<%
		objRSAllDates.Open "SELECT DISTINCT GameDate FROM tblLeagueSetup WHERE GameDate >= DATE() ORDER BY GameDate", objConn	
								
%>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>NBA Schedule</strong>
			</div>
		</div>
	</div>
</div>
<form action="NBATeamSkeds.asp" name="maintainTeamSkeds" method="POST" onSubmit="return processTeams(this)">
<input type="hidden" name="gameDay" value="<%= objRSTMSkeds.Fields("gameday").Value %>" />
<!--#include virtual="Common/headermain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="panel panel-override">
				<div class="panel-body">
					<%						
						While Not objRSAllDates.EOF
						loopGameDate = objRSAllDates.Fields("GameDate").Value
						
						objRSTMSkeds.Open   "SELECT TeamName,'at' as Location, Opponent, GameDate, TipTimeEst - 1/24 as Tip_CST " &_ 
																"FROM tblLeagueSetup " &_ 
																"WHERE GameLoc = '@' AND gamedate = #"&loopGameDate&"# ORDER BY GameDate, TipTimeEst",objConn,3,3,1	

						gameCnt = objRSTMSkeds.recordCount										
					%>
					<% if gameCnt >=7 then %>
						<span style="font-weight:bold;font-size:16px;text-transform:uppercase;"><%=FormatDateTime(loopGameDate,1)%></span>&nbsp;<mark>[IGBL GAME]</mark></br></br>
					<%else%>
						<span style="font-weight:bold;font-size:16px;text-transform:uppercase;"><%=FormatDateTime(loopGameDate,1)%></span></br></br>
					<%end if%>
					
					<table class="table table-custom-black table-responsive table-condensed">
						<tr style="font-weight:bold;background-color:#dff0d8;">
							<td class="big" style="text-align:left;">Team</td>
							<td class="big" style="text-align:left;">Opponent</td>
							<td class="big" style="text-align:left;">Tip-Time</td>
						</tr>
						<%
							
							While Not objRSTMSkeds.EOF
						%>
							<tr style="background-color:#dff0d8;">
								<td class="big" style="vertical-align:middle;text-align:left;width:38%;background-color:white;"><%=objRSTMSkeds.Fields("TeamName").Value%></td>
								<td class="big" style="vertical-align:middle;text-align:left;width:38%;background-color:white;"><%=objRSTMSkeds.Fields("Opponent").Value%></td>
								<td class="big" style="vertical-align:middle;text-align:left;width:24%;background-color:white;"><%=objRSTMSkeds.Fields("Tip_CST").Value%></td>
							</tr>					
						<%
							objRSTMSkeds.MoveNext
							Wend
							objRSTMSkeds.Close							
						%>
				</table>
				<br>
					
					
					
					
				<%					
					objRSAllDates.MoveNext
					Wend 
					objRSAllDates.Close					
				%>				
				</div>
			</div>
		</div>
	</div>
</div>
</form>
<br>
<%
	objRSTMSkeds.Close
	objConn.Close
	Set objRSTMSkeds = Nothing
	Set ObjConn      = Nothing

%>
</body>
</html>