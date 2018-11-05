<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""

	Dim objConn,objRSAl,objRSLogs,objRSWork
	
	Set objConn = Server.CreateObject("ADODB.Connection")
	Set objRSWork = Server.CreateObject("ADODB.RecordSet") 

	objConn.Open Application("lineupstest_ConnectionString")
	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
							 "Data Source=lineupstest.mdb;" & _
							 "Persist Security Info=False"
					 
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
.panel-calendar {
  color:black;
  background-color:white;
  border-color: black;
  border-radius:0px;
}
.modal-header {
    padding: 7px;
    border-bottom: 1px solid #e5e5e5;
}
.modal-header-modal {
    color: white;
    padding: 9px 15px;
    border-bottom: 1px solid #eee;
    background-color:#468847;
    -webkit-border-top-left-radius: 5px;
    -webkit-border-top-right-radius: 5px;
    -moz-border-radius-topleft: 5px;
    -moz-border-radius-topright: 5px;
    border-top-left-radius: 5px;
    border-top-right-radius: 5px;
    font-weight: bold;
}
.item {
	background: #333;
	text-align: center;
	height: 120px !important;
}
redW {
	color: #9a1400;
}
white {
	color: white;
}
.bs-callout-success {
	border-left-color: darkorange;
	padding: 10px;
	border-left-width: 5px;
	border-radius: 3px;
	background-color: white;
}
.bs-callout-success h4 {
	color: darkorange;
}
dodgerblue {
	color: dodgerblue;
}
white {
	color: white;
}
th {
	vertical-align: middle;
	text-align: center;
}


darkorange {
	color:darkorange;
}
td.opp {
	vertical-align: middle;		
	font-size:9px;	
}
td.opp {
	vertical-align: middle;		
	font-size:11px;	
}
table.box{
	border-bottom-color:black;
	border-bottom-style:double;
	border-bottom-width:thick;
	border-left-color:black;
	border-left-style:solid;
	border-left-width:thin;
	border-right-color:black;
	border-right-style:solid;
	border-right-width:thin;
}

span.redText {
	color:#9a1400;	
}
span.greenText {
	color:#468847;	 
}
div.contained_table{
	width:100%;
	overflow-x: auto;
}
gameTip {
    color: #616161;
    font-weight: normal;
}
orange {
    color: darkorange;
}
.auctionText {
    color: #48494a;
    font-weight: bold;
}
td.green{
	color:yellowgreen;
}
td.blue{
	color:dodgerblue;
}
.nav-tabs {
	border-bottom: 2px solid black;
}
.badgeTie {
	display: inline-block;
	min-width: 10px;
	padding: 3px 6px;
	line-height: 1;
	color: white;
	text-align: center;
	white-space: nowrap;
	vertical-align: baseline;
	background-color: yellow;
	border-radius: 14px;
	border: white;
	color: black;
	font-weight: 700 !important;
	border-width: thick;
}
button{
	min-width: 20px;
	min-height: 20px;
}
.fire {
	color: orangered;
}
td {
	vertical-align: middle;
	text-align: center;
	font-size:11px;	
}
</style>
</head>
<body>
<script language="JavaScript" type="text/javascript">
$(document).ready(function() {
    $('#example').DataTable( {
        "order": [[ 0, "desc" ]],
			  "lengthMenu": [[25, 50, 75, -1], [25, 50, 75, "All"]]
    } );
} );
</script>
<!--#include virtual="Common/headerMain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Results</strong>
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<ul class="nav nav-tabs">
				<li class="active"><a data-toggle="tab" href="#box"><i class="fal fa-calculator"></i>&nbsp;Box</a></li>
				<!--<li><a data-toggle="tab" href="#leaders"><i class="fal fa-sort-amount-down"></i>&nbsp;Nightly Leaders</a></li>-->
				<li><a data-toggle="tab" href="#gameLogs"><i class="fas fa-folder-open"></i>&nbsp;Game Logs</a></li>
				<li><a data-toggle="tab" href="#scoring"><i class="fas fa-user-cog"></i>&nbsp;Scoring Settings</a></li>
				<%
				objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'PLAYOFF_START_DATE' ",objConn
				wPlayoffStart = objRSWork.Fields("param_date").Value
				objRSWork.Close				
				if (date() >= wPlayoffStart) then 
				%>
					<li><a data-toggle="tab" href="#gameLogsPO">PO Logs</a></li>
				<% end if %>
				</ul>
		</div>
	</div>
	</br>
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="tab-content">
				<div id="scoring" class="tab-pane">
					<!--#include virtual="Common/scoring.inc"-->
				</div>
				<div id="gameLogs" class="tab-pane">
					<!--#include virtual="Common/gamelogs.inc"-->
				</div>
				<div id="gameLogsPO" class="tab-pane">
					<!--#include virtual="Common/gamelogsPO.inc"-->					
				</div>
				<div id="box" class="tab-pane fade in active">
				<% 
				Set objRSLastNight= Server.CreateObject("ADODB.RecordSet")
				Set objRSLongDate = Server.CreateObject("ADODB.RecordSet")
				Set objRSAwaylogo = Server.CreateObject("ADODB.RecordSet")
				Set objRSHomelogo = Server.CreateObject("ADODB.RecordSet")
				Set objRS         = Server.CreateObject("ADODB.RecordSet")
				Set objRSGameDeadLines = Server.CreateObject("ADODB.RecordSet")
				
				objRSLastNight.Open  "SELECT MAX(gameDate) as gameday FROM newbox", objConn,3,3,1		
				if IsNull(objRSLastNight.Fields("gameday").Value) then
					gameDate = date()
				else
					gameDate = objRSLastNight.Fields("gameday").Value
				end if

				dim icount, loopcount, hcount, acount,htToWon,atToWon
				icount   = 0 
				loopcount= 1
				hcount   = 0 
				acount   = 6
				dim hscore, ascore, turn, turn1

				objRSLongDate.Open "SELECT format(#"&gameDate&"#, 'Long Date') as DisplayDate",objConn,3,3,1
				wDisplayDate = objRSLongDate.Fields("DisplayDate").Value 
				
				objRSLongDate.Close

				objRS.Open  	"SELECT * FROM NEWBOX WHERE gameDate = #"&gameDate&"#",objConn,3,3,1		
				objRSGameDeadLines.open   "SELECT * FROM tblGameDeadLines WHERE gameDay = #"&gameDate&"#",objConn,3,3,1							
				%>

				<div class="row">		
					<div class="col-md-12 col-sm-12 col-xs-12">			
					<span style="font-size:14px;"><strong>IGBL Scoreboard</strong></span></br>
					<%=FormatDateTime(gameDate,1)%></br><span style="color:#01579B;font-weight:bold;">GAME#:</span> <%=objRSGameDeadLines.Fields("IGBLGAMENUM").Value%> | <span style="color:#9a1400;font-weight:bold;">CYCLE:</span> <%=objRSGameDeadLines.Fields("cycle").Value%>
					</div>
				</div>
					<!--#include virtual="Common/gamedayResultNoTab.inc"-->				
				</div>
			</div>		
		</div>
	</div>
</div>

<!--#include virtual="Common/functions.inc"-->

<%
	objRSLastNight.Close
	objRSLongDate.Close 
	objRSAwaylogo.Close 
	objRSHomelogo.Close 
	objRBox.Close
	objRS.Close         
  objConn.Close
	objRSLeaders.close
	objRS.Close
	objRSAll.close
	objRSLogs.close
	objRSStagger.close
  Set objConn = Nothing
%>
</body>
</html>