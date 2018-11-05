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
	
	Dim objConn, objRS,objRSHomeLogo,objRSAwaylogo
	
	Set objConn            = Server.CreateObject("ADODB.Connection")
	Set objRS              = Server.CreateObject("ADODB.RecordSet")
	Set objRSHomeLogo      = Server.CreateObject("ADODB.RecordSet")
	Set objRSAwaylogo      = Server.CreateObject("ADODB.RecordSet")
	Set objRSGameDeadLines = Server.CreateObject("ADODB.RecordSet")
	Set objRSPlayer        = Server.CreateObject("ADODB.RecordSet")
	Set objRSWork          = Server.CreateObject("ADODB.RecordSet")

	objConn.Open Application("lineupstest_ConnectionString")
	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"

									
	gameDate = Request.querystring("gameDate")

	dim icount, loopcount, hcount, acount,htToWon,atToWon
	icount   = 0 
	loopcount= 1
	hcount   = 0 
	acount   = 6
	dim hscore, ascore, turn, turn1

	objRS.Open            "SELECT * FROM NEWBOX WHERE gameDate = #"&gameDate&"#",objConn,3,3,1		
	objRSGameDeadLines.open   "SELECT * FROM tblGameDeadLines WHERE gameDay = #"&gameDate&"#",objConn,3,3,1			
%>
<!DOCTYPE HTML>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="">
<meta name="author" content="">
<!--#include virtual="Common/pageHeader.inc"-->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/tabs.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
th {
	vertical-align: middle;
	text-align: center;
	}
td {
	vertical-align: middle;
	text-align: center;
	font-size: 12px;
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
<!--#include virtual="Common/session.inc"-->
<!--#include virtual="Common/headerMain.inc"-->

<body>
<div class="container">
	<div class="row">		
		<div class="col-md-12 col-sm-12 col-xs-12">			
				<span style="font-size:14px;"><strong>IGBL Scoreboard</strong></span></br>
				<%=FormatDateTime(gameDate,1)%></br><span style="color:#01579B;font-weight:bold;">GAME#:</span> <%=objRSGameDeadLines.Fields("IGBLGAMENUM").Value%> | <span style="color:#9a1400;font-weight:bold;">CYCLE:</span> <%=objRSGameDeadLines.Fields("cycle").Value%>
		</div>
	</div>

	
<!--#include virtual="Common/gameDayResultNoTabBox.inc"-->
<!--#include virtual="Common/functions.inc"-->


<%
objRSLastNight.Close
objRSLongDate.Close 
objRSAwaylogo.Close 
objRSHomelogo.Close 
objRBox.Close
objRS.Close   
ObjConn.Close
Set objRS = Nothing
Set objConn = Nothing
Session.CodePage = Session("FP_OldCodePage")
Session.LCID = Session("FP_OldLCID")
%>
</div>
</body>
</html>