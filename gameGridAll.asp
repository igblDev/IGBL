<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""

	Dim objRSgames,objRS,objConn, objRSwaivers, ownerid,objParams,objrsTeams
	Dim objrsATL,objrsBKN,objrsBOS,objrsCHA,objrsCHI,objrsCLE,objrsDAL,objrsDEN,objrsDET,objrsGSW
	Dim objrsHOU,objrsIND,objrsLAC,objrsLAL,objrsMEM,objrsMIA,objrsMIL,objrsMIN,objrsNOP,objrsNYK
	Dim objrsOKC,objrsORL,objrsPHI,objrsPHX,objrsPOR,objrsSAC,objrsSAS,objrsTOR,objrsUTA,objrsWAS
	
	%>
	<!--#include virtual="Common/SESSION.inc"-->
	<%	

	Set objConn       = Server.CreateObject("ADODB.Connection")	
	Set objParams     = Server.CreateObject("ADODB.RecordSet")
	Set objRS         = Server.CreateObject("ADODB.RecordSet")
	Set objRSgames    = Server.CreateObject("ADODB.RecordSet")
	Set objrsATL      = Server.CreateObject("ADODB.RecordSet")
	Set objrsBKN      = Server.CreateObject("ADODB.RecordSet")	
	Set objrsBOS      = Server.CreateObject("ADODB.RecordSet")
	Set objrsCHA      = Server.CreateObject("ADODB.RecordSet")
	Set objrsCHI      = Server.CreateObject("ADODB.RecordSet")
	Set objrsCLE      = Server.CreateObject("ADODB.RecordSet")
	Set objrsDAL      = Server.CreateObject("ADODB.RecordSet")
	Set objrsDEN      = Server.CreateObject("ADODB.RecordSet")
	Set objrsDET      = Server.CreateObject("ADODB.RecordSet")
	Set objrsGSW      = Server.CreateObject("ADODB.RecordSet")	
	Set objrsHOU      = Server.CreateObject("ADODB.RecordSet")
	Set objrsIND      = Server.CreateObject("ADODB.RecordSet")
	Set objrsLAC      = Server.CreateObject("ADODB.RecordSet")
	Set objrsLAL      = Server.CreateObject("ADODB.RecordSet")	
	Set objrsMEM      = Server.CreateObject("ADODB.RecordSet")
	Set objrsMIA      = Server.CreateObject("ADODB.RecordSet")
	Set objrsMIL      = Server.CreateObject("ADODB.RecordSet")
	Set objrsMIN      = Server.CreateObject("ADODB.RecordSet")
	Set objrsNOP      = Server.CreateObject("ADODB.RecordSet")
	Set objrsNYK      = Server.CreateObject("ADODB.RecordSet")	
	Set objrsOKC      = Server.CreateObject("ADODB.RecordSet")
	Set objrsORL      = Server.CreateObject("ADODB.RecordSet")
	Set objrsPHI      = Server.CreateObject("ADODB.RecordSet")
	Set objrsPHX      = Server.CreateObject("ADODB.RecordSet")	
	Set objrsPOR      = Server.CreateObject("ADODB.RecordSet")
	Set objrsSAC      = Server.CreateObject("ADODB.RecordSet")
	Set objrsSAS      = Server.CreateObject("ADODB.RecordSet")
	Set objrsTOR      = Server.CreateObject("ADODB.RecordSet")
	Set objrsUTA      = Server.CreateObject("ADODB.RecordSet")
	Set objrsWAS      = Server.CreateObject("ADODB.RecordSet")	

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
<link href="css/tabs.css" rel="stylesheet">
<style>
.red {
	color:black;
	background-color:red;
	font-weight:bold;
}
th {
    vertical-align: middle;
  	text-align: center;
		font-size:8px;
}

td {
    vertical-align: middle;
		font-size:10px;
		text-align:center;
		width: 9%;
		font-weight:bold;
}

red {
	color: red;
}
white {
	color: white;
}
gray {
		color: gray;
}

.modal-header-success {
    color:#fff;
    padding:9px 15px;
    border-bottom:1px solid #eee;
    background-color: darkorange;
    -webkit-border-top-left-radius: 5px;
    -webkit-border-top-right-radius: 5px;
    -moz-border-radius-topleft: 5px;
    -moz-border-radius-topright: 5px;
     border-top-left-radius: 5px;
     border-top-right-radius: 5px;
}
</style>
</head>
<body>
<script language="JavaScript" type="text/javascript">
</script>

<!--#include virtual="Common/headerMain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				Availability Grid
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<ul class="nav nav-tabs">
				<li class="active"><a data-toggle="tab" href="#reg">REG</a></li>
				<li><a data-toggle="tab" href="#wk">WK</a></li>
				<li><a data-toggle="tab" href="#qtrs">QTR</a></li>
				<li><a data-toggle="tab" href="#semis">SEMI</a></li>
				<li><a data-toggle="tab" href="#finals">FINAL</a></li>
				</ul>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">	
			<div class="tab-content">
				<div id="wk" class="tab-pane fade">
				    <div class="col-md-12 col-sm-12 col-xs-12">
		<table class="table table-bordered table-custom-black table-responsive table-condensed">
				<thead>
					<tr style="color:black;font-weight:bold;">
						<th style="text-align:center;width:10%;">WK</th>
						<th style="text-align:center;width:10%;">M</th>
						<th style="text-align:center;width:10%;">T</th>
						<th style="text-align:center;width:10%;">W</th>
						<th style="text-align:center;width:10%;">TH</th>
						<th style="text-align:center;width:10%;">F</th>
						<th style="text-align:center;width:10%;">S</th>
						<th style="text-align:center;width:10%;">SU</th>
					</tr>
				</thead>
				<tbody>
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">10-22</td>
  					<td style="vertical-align:middle;text-align:center;">9</td>
						<td style="vertical-align:middle;text-align:center">3</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">11</td>
						<td style="vertical-align:middle;text-align:center">4</td>
  					<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3">9</td>
						<td style="vertical-align:middle;text-align:center">4</td>
					</tr>
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">10-29</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3">9</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
						<td style="vertical-align:middle;text-align:center;background-color: red;color:white;">7</td>
						<td style="vertical-align:middle;text-align:center">6</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
						<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
					</tr>							
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">11-5</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
						<td style="vertical-align:middle;text-align:center">4</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>
						<td style="vertical-align:middle;text-align:center">4</td>
  					<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>
						<td style="vertical-align:middle;text-align:center">6</td>
					</tr>
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">11-12</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
						<td style="vertical-align:middle;text-align:center">3</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">11</td>
						<td style="vertical-align:middle;text-align:center">3</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>
						<td style="vertical-align:middle;text-align:center">5</td>
					</tr>
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">11-19</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
						<td style="vertical-align:middle;text-align:center">4</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">13</td>
						<td style="vertical-align:middle;text-align:center;background-color:blue;color: white;">TG</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">14</td>
						<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
					</tr>							
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">11-26</td>
  					<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
						<td style="vertical-align:middle;text-align:center">5</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>
						<td style="vertical-align:middle;text-align:center">3</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">11</td>
  					<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
						<td style="vertical-align:middle;text-align:center">6</td>
					</tr>
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">12-3</td>
						<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
						<td style="vertical-align:middle;text-align:center">5</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>
						<td style="vertical-align:middle;text-align:center">3</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
						<td style="vertical-align:middle;text-align:center">4</td>
					</tr>	
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">12-10</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">11</td>
						<td style="vertical-align:middle;text-align:center">3</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">11</td>
						<td style="vertical-align:middle;text-align:center">4</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
						<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
						<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
					</tr>	
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">12-17</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
						<td style="vertical-align:middle;text-align:center">4</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">12</td>
						<td style="vertical-align:middle;text-align:center">2</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>
						<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">11</td>
					</tr>	
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">12-24</td>
						<td style="vertical-align:middle;text-align:center;background-color:blue;color: white;">XE</td>
						<td style="vertical-align:middle;text-align:center">5</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>
						<td style="vertical-align:middle;text-align:center">5</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
						<td style="vertical-align:middle;text-align:center">6</td>
					</tr>	
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">12-31</td>
  					<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
						<td style="vertical-align:middle;text-align:center">5</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
						<td style="vertical-align:middle;text-align:center">3</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>	
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
						<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
					</tr>	
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">01-7</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>
						<td style="vertical-align:middle;text-align:center">4</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
						<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
					</tr>	
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">01-14</td>
						<td style="vertical-align:middle;text-align:center">6</td>
						<td style="vertical-align:middle;text-align:center">6</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
						<td style="vertical-align:middle;text-align:center">6</td>
						<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>
						<td style="vertical-align:middle;text-align:center">3</td>
					</tr>	
					<tr style="text-align:center;background-color: white;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">01-21</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">11</td>
						<td style="vertical-align:middle;text-align:center">4</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>
						<td style="vertical-align:middle;text-align:center">4</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>
  					<td style="vertical-align:middle;text-align:center">5</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
					</tr>					
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">01-28</td>
						<td style="vertical-align:middle;text-align:center">5</td>
						<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
						<td style="vertical-align:middle;text-align:center">6</td>
  					<td style="vertical-align:middle;text-align:center">5</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">12</td>
						<td style="vertical-align:middle;text-align:center">3</td>
					</tr>	
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">02-4</td>
						<td style="vertical-align:middle;text-align:center">6</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
						<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
						<td style="vertical-align:middle;text-align:center">6</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
						<td style="vertical-align:middle;text-align:center">5</td>
					</tr>	
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">02-11</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
						<td style="vertical-align:middle;text-align:center">5</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">11</td>
						<td style="vertical-align:middle;text-align:center">3</td>
						<td style="vertical-align:middle;text-align:center;background-color:blue;color: white;">A</td>
						<td style="vertical-align:middle;text-align:center;background-color:blue;color: white;">S</td>
						<td style="vertical-align:middle;text-align:center;background-color:blue;color: white;">B</td>
					</tr>	
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">02-18</td>
						<td style="vertical-align:middle;text-align:center;background-color:blue;color: white;">A</td>
						<td style="vertical-align:middle;text-align:center;background-color:blue;color: white;">S</td>
						<td style="vertical-align:middle;text-align:center;background-color:blue;color: white;">B</td>
						<td style="vertical-align:middle;text-align:center">6</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">12</td>
						<td style="vertical-align:middle;text-align:center">3</td>
					</tr>	
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">02-25</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">11</td>
						<td style="vertical-align:middle;text-align:center">3</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">11</td>
						<td style="vertical-align:middle;text-align:center">6</td>
  					<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
					</tr>	
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">03-4</td>
  					<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
						<td style="vertical-align:middle;text-align:center">6</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>
						<td style="vertical-align:middle;text-align:center">2</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
  					<td style="vertical-align:middle;text-align:center">6</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
					</tr>	
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">03-11</td>
						<td style="vertical-align:middle;text-align:center">6</td>
  					<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
						<td style="vertical-align:middle;text-align:center">6</td>
						<td style="vertical-align:middle;text-align:center">6</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
					</tr>	
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">03-18</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
						<td style="vertical-align:middle;text-align:center">6</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">9</td>
						<td style="vertical-align:middle;text-align:center">6</td>
						<td style="vertical-align:middle;text-align:center;background-color: red;color: white;">7</td>
  					<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">8</td>
					</tr>	
					<tr style="text-align:center;background-color: white;">
  					<td style="vertical-align:middle;text-align:center">03-25</td>
						<td style="vertical-align:middle;text-align:center;">4</td>
						<td style="vertical-align:middle;text-align:center;background-color: #fcf8e3;">10</td>
						<td style="vertical-align:middle;text-align:center;">O</td>
						<td style="vertical-align:middle;text-align:center;">V</td>
  					<td style="vertical-align:middle;text-align:center;">E</td>
  					<td style="vertical-align:middle;text-align:center;">R</td>
						<td style="vertical-align:middle;text-align:center;">!</td>
					</tr>	
				<thead>
					<tr style="color:black;font-weight:bold;">
						<th style="text-align:center;width:10%;">WK</th>
						<th style="text-align:center;width:10%;">M</th>
						<th style="text-align:center;width:10%;">T</th>
						<th style="text-align:center;width:10%;">W</th>
						<th style="text-align:center;width:10%;">TH</th>
						<th style="text-align:center;width:10%;">F</th>
						<th style="text-align:center;width:10%;">S</th>
						<th style="text-align:center;width:10%;">SU</th>
					</tr>
				</thead>
				</tbody>
			</table>
			</br>
    </div>
				</div>
				<div id="reg" class="tab-pane fade in active">
				<!--#include virtual="Common/gamegridReg.inc"-->
				<div class="col-md-4">
					<table class="table table-bordered table-striped table-responsive table-condensed">
							<tr style="color:black;font-weight:bold;">
								<td></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/ATL.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/BKN.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/BOS.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/CHA.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/CHI.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/CLE.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/DAL.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/DEN.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/DET.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/GSW.svg" alt="Atlanta Hawks"></td>
							</tr>
						<%
							objRS.Open   "SELECT * FROM tblGameGrid where tblGameGrid.GameDate >= date()", objConn,3,3,1
						%>
						<%
							While Not obJRS.EOF
								if objRS.Fields("GameDate").Value < cdate(dPODate) then
						%>
						<tr>
					<% if len(objRS.Fields("GameDate")) = 8 then %>
						<td><%=left(objRS.Fields("GameDate").Value,3) %></td>	
					<% elseif len(objRS.Fields("GameDate")) = 9 then %>
						<td><%=left(objRS.Fields("GameDate").Value,4) %></td>					
					<% else %>
						<td><%=left(objRS.Fields("GameDate").Value,5) %></td>					
					<%end if%>
							<td><%=objRS.Fields("ATL").Value %></td>
							<td><%=objRS.Fields("BKN").Value %></td>
							<td><%=objRS.Fields("BOS").Value %></td>
							<td><%=objRS.Fields("CHA").Value %></td>
							<td><%=objRS.Fields("CHI").Value %></td>
							<td><%=objRS.Fields("CLE").Value %></td>
							<td><%=objRS.Fields("DAL").Value %></td>
							<td><%=objRS.Fields("DEN").Value %></td>
							<td><%=objRS.Fields("DET").Value %></td>
							<td><%=objRS.Fields("GSW").Value %></td>
						</tr>
						<%
								end if
								objRS.MoveNext
							Wend
						%>
						<%
						objRS.Close
						%>
						<tr style="color:black;font-weight:bold;">
							<td width="10%">GM</td>
							<td>ATL</td>
							<td>BKN</td>
							<td>BOS</td>
							<td>CHA</td>
							<td>CHI</td>
							<td>CLE</td>
							<td>DAL</td>
							<td>DEN</td>
							<td>DET</td>
							<td>GSW</td>
						</tr>
						<tr class="warning">
							<td>ToT</td>
							<td><%=objrsATL.RecordCount%></td>
							<td><%=objrsBKN.RecordCount%></td>
							<td><%=objrsBOS.RecordCount%></td>
							<td><%=objrsCHA.RecordCount%></td>
							<td><%=objrsCHI.RecordCount%></td>
							<td><%=objrsCLE.RecordCount%></td>
							<td><%=objrsDAL.RecordCount%></td>
							<td><%=objrsDEN.RecordCount%></td>
							<td><%=objrsDET.RecordCount%></td>
							<td><%=objrsGSW.RecordCount%></td>
						</tr>
					</table>
					</br>
				</div>
				<div class="col-md-4">
					<table class="table table-bordered table-striped table-responsive table-condensed">
							<tr style="color:black;font-weight:bold;">
								<td></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/HOU.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/IND.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/LAC.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/LAL.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MEM.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MIA.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MIL.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MIN.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/NOP.svg" alt="Atlanta Hawks"></td>
								<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/NYK.svg" alt="Atlanta Hawks"></td>
						</tr>
						<%
							objRS.Open   "SELECT * FROM tblGameGrid where tblGameGrid.GameDate >= date()", objConn,3,3,1
						%>
						<%
							While Not obJRS.EOF 
								if objRS.Fields("GameDate").Value < cdate(dPODate) then
						%>
						<tr>
					<% if len(objRS.Fields("GameDate")) = 8 then %>
						<td><%=left(objRS.Fields("GameDate").Value,3) %></td>	
					<% elseif len(objRS.Fields("GameDate")) = 9 then %>
						<td><%=left(objRS.Fields("GameDate").Value,4) %></td>					
					<% else %>
						<td><%=left(objRS.Fields("GameDate").Value,5) %></td>					
					<%end if%>
							<td><%=objRS.Fields("HOU").Value %></td>
							<td><%=objRS.Fields("IND").Value %></td>
							<td><%=objRS.Fields("LAC").Value %></td>
							<td><%=objRS.Fields("LAL").Value %></td>
							<td><%=objRS.Fields("MEM").Value %></td>
							<td><%=objRS.Fields("MIA").Value %></td>
							<td><%=objRS.Fields("MIL").Value %></td>
							<td><%=objRS.Fields("MIN").Value %></td>
							<td><%=objRS.Fields("NOP").Value %></td>
							<td><%=objRS.Fields("NYK").Value %></td>
						</tr>
						<%
								end if
								objRS.MoveNext
							Wend
						%>
						<%
						objRS.Close
						%>
						<tr style="color:black;font-weight:bold;">
							<td width="10%">GM</td>
							<td>HOU</td>
							<td>IND</td>
							<td>LAC</td>
							<td>LAL</td>
							<td>MEM</td>
							<td>MIA</td>
							<td>MIL</td>
							<td>MIN</td>
							<td>NOP</td>
							<td>NYK</td>
						</tr>
						<tr class="warning">
							<td>ToT</td>
							<td><%=objrsHOU.RecordCount%></td>
							<td><%=objrsIND.RecordCount%></td>
							<td><%=objrsLAC.RecordCount%></td>
							<td><%=objrsLAL.RecordCount%></td>
							<td><%=objrsMEM.RecordCount%></td>
							<td><%=objrsMIA.RecordCount%></td>
							<td><%=objrsMIL.RecordCount%></td>
							<td><%=objrsMIN.RecordCount%></td>
							<td><%=objrsNOP.RecordCount%></td>
							<td><%=objrsNYK.RecordCount%></td>
						</tr>
					</table>
					</br>
				</div>
				 <div class="col-md-4">
					<table class="table table-bordered table-striped table-responsive table-condensed">
				<tr style="color:black;font-weight:bold;">
					<td></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/OKC.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/ORL.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/PHI.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/PHX.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/POR.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/SAC.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/SAS.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/TOR.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/UTA.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/WAS.svg" alt="Atlanta Hawks"></td>
				</tr>
						<%
							objRS.Open   "SELECT * FROM tblGameGrid where tblGameGrid.GameDate >= date()", objConn,3,3,1
						%>
						<%
							While Not obJRS.EOF 
								if objRS.Fields("GameDate").Value < cdate(dPODate) then
						%>
						<tr>
					<% if len(objRS.Fields("GameDate")) = 8 then %>
						<td><%=left(objRS.Fields("GameDate").Value,3) %></td>	
					<% elseif len(objRS.Fields("GameDate")) = 9 then %>
						<td><%=left(objRS.Fields("GameDate").Value,4) %></td>					
					<% else %>
						<td><%=left(objRS.Fields("GameDate").Value,5) %></td>					
					<%end if%>
							<td><%=objRS.Fields("OKC").Value %></td>
							<td><%=objRS.Fields("ORL").Value %></td>
							<td><%=objRS.Fields("PHI").Value %></td>
							<td><%=objRS.Fields("PHX").Value %></td>
							<td><%=objRS.Fields("POR").Value %></td>
							<td><%=objRS.Fields("SAC").Value %></td>
							<td><%=objRS.Fields("SAS").Value %></td>
							<td><%=objRS.Fields("TOR").Value %></td>
							<td><%=objRS.Fields("UTA").Value %></td>
							<td><%=objRS.Fields("WAS").Value %></td>
						</tr>
						<%
								end if
								objRS.MoveNext
							Wend
						%>
						<%
						objRS.Close
						%>
						<tr style="color:black;font-weight:bold;">
								<td width="10%">GM</td>
								<td>OKC</td>
								<td>ORL</td>
								<td>PHI</td>
								<td>PHX</td>
								<td>POR</td>
								<td>SAC</td>
								<td>SAS</td>
								<td>TOR</td>
								<td>UTA</td>
								<td>WAS</td>	
							</tr>
						<tr class="warning">
							<td>ToT</td>
							<td><%=objrsOKC.RecordCount%></td>
							<td><%=objrsORL.RecordCount%></td>
							<td><%=objrsPHI.RecordCount%></td>
							<td><%=objrsPHX.RecordCount%></td>
							<td><%=objrsPOR.RecordCount%></td>
							<td><%=objrsSAC.RecordCount%></td>
							<td><%=objrsSAS.RecordCount%></td>
							<td><%=objrsTOR.RecordCount%></td>
							<td><%=objrsUTA.RecordCount%></td>
							<td><%=objrsWAS.RecordCount%></td>
						</tr>
					</table>
				</div>
			</div>
			<!--#include virtual="Common/gamegridCloseIO.inc"-->
			<div id="qtrs" class="tab-pane fade">
			<!--#include virtual="Common/gamegridQtrs.inc"-->
			<div class="col-md-4">
				<table class="table table-bordered table-striped table-responsive table-condensed">
				<tr style="color:black;font-weight:bold;">
					<td></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/ATL.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/BKN.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/BOS.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/CHA.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/CHI.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/CLE.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/DAL.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/DEN.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/DET.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/GSW.svg" alt="Atlanta Hawks"></td>
				</tr>
					<%
						objRS.Open "SELECT * FROM tblGameGrid where tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1

					%>
					<%
						While Not objRS.EOF
					%>
					<tr>
					<% if len(objRS.Fields("GameDate")) = 8 then %>
						<td><%=left(objRS.Fields("GameDate").Value,3) %></td>	
					<% elseif len(objRS.Fields("GameDate")) = 9 then %>
						<td><%=left(objRS.Fields("GameDate").Value,4) %></td>					
					<% else %>
						<td><%=left(objRS.Fields("GameDate").Value,5) %></td>					
					<%end if%>
						<td><%=objRS.Fields("ATL").Value %></td>
						<td><%=objRS.Fields("BKN").Value %></td>
						<td><%=objRS.Fields("BOS").Value %></td>
						<td><%=objRS.Fields("CHA").Value %></td>
						<td><%=objRS.Fields("CHI").Value %></td>
						<td><%=objRS.Fields("CLE").Value %></td>
						<td><%=objRS.Fields("DAL").Value %></td>
						<td><%=objRS.Fields("DEN").Value %></td>
						<td><%=objRS.Fields("DET").Value %></td>
						<td><%=objRS.Fields("GSW").Value %></td>
					</tr>
					<%
							objRS.MoveNext
						Wend
					%>
					<%
					objRS.Close
					%>
					<tr style="color:black;font-weight:bold;">
						<td width="10%">GM</td>
						<td>ATL</td>
						<td>BKN</td>
						<td>BOS</td>
						<td>CHA</td>
						<td>CHI</td>
						<td>CLE</td>
						<td>DAL</td>
						<td>DEN</td>
						<td>DET</td>
						<td>GSW</td>
					</tr>
					<tr class="warning">
						<td>ToT</td>
						<td><%=objrsATL.RecordCount%></td>
						<td><%=objrsBKN.RecordCount%></td>
						<td><%=objrsBOS.RecordCount%></td>
						<td><%=objrsCHA.RecordCount%></td>
						<td><%=objrsCHI.RecordCount%></td>
						<td><%=objrsCLE.RecordCount%></td>
						<td><%=objrsDAL.RecordCount%></td>
						<td><%=objrsDEN.RecordCount%></td>
						<td><%=objrsDET.RecordCount%></td>
						<td><%=objrsGSW.RecordCount%></td>
					</tr>
				</table>
				</br>
			</div>
			<div class="col-md-4">
				<table class="table table-bordered table-striped table-responsive table-condensed">
					<tr style="color:black;font-weight:bold;">
						<td></td>
						<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/HOU.svg" alt="Atlanta Hawks"></td>
						<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/IND.svg" alt="Atlanta Hawks"></td>
						<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/LAC.svg" alt="Atlanta Hawks"></td>
						<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/LAL.svg" alt="Atlanta Hawks"></td>
						<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MEM.svg" alt="Atlanta Hawks"></td>
						<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MIA.svg" alt="Atlanta Hawks"></td>
						<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MIL.svg" alt="Atlanta Hawks"></td>
						<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MIN.svg" alt="Atlanta Hawks"></td>
						<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/NOP.svg" alt="Atlanta Hawks"></td>
						<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/NYK.svg" alt="Atlanta Hawks"></td>
				</tr>
				<%
					objRS.Open "SELECT * FROM tblGameGrid where tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
				%>
				<%
					While Not objRS.EOF
				%>
				<tr>
					<% if len(objRS.Fields("GameDate")) = 8 then %>
						<td><%=left(objRS.Fields("GameDate").Value,3) %></td>	
					<% elseif len(objRS.Fields("GameDate")) = 9 then %>
						<td><%=left(objRS.Fields("GameDate").Value,4) %></td>					
					<% else %>
						<td><%=left(objRS.Fields("GameDate").Value,5) %></td>					
					<%end if%>
					<td><%=objRS.Fields("HOU").Value %></td>
					<td><%=objRS.Fields("IND").Value %></td>
					<td><%=objRS.Fields("LAC").Value %></td>
					<td><%=objRS.Fields("LAL").Value %></td>
					<td><%=objRS.Fields("MEM").Value %></td>
					<td><%=objRS.Fields("MIA").Value %></td>
					<td><%=objRS.Fields("MIL").Value %></td>
					<td><%=objRS.Fields("MIN").Value %></td>
					<td><%=objRS.Fields("NOP").Value %></td>
					<td><%=objRS.Fields("NYK").Value %></td>
				</tr>
					<%
						objRS.MoveNext
							Wend
					%>
					<%
					objRS.Close
					%>
					<tr style="color:black;font-weight:bold;">
						<td width="10%">GM</td>
						<td>HOU</td>
						<td>IND</td>
						<td>LAC</td>
						<td>LAL</td>
						<td>MEM</td>
						<td>MIA</td>
						<td>MIL</td>
						<td>MIN</td>
						<td>NOP</td>
						<td>NYK</td>
					</tr>
					<tr class="warning">
						<td>ToT</td>
						<td><%=objrsHOU.RecordCount%></td>
						<td><%=objrsIND.RecordCount%></td>
						<td><%=objrsLAC.RecordCount%></td>
						<td><%=objrsLAL.RecordCount%></td>
						<td><%=objrsMEM.RecordCount%></td>
						<td><%=objrsMIA.RecordCount%></td>
						<td><%=objrsMIL.RecordCount%></td>
						<td><%=objrsMIN.RecordCount%></td>
						<td><%=objrsNOP.RecordCount%></td>
						<td><%=objrsNYK.RecordCount%></td>
					</tr>
				</table>
				</br>
			</div>
			<div class="col-md-4">
				<table class="table table-bordered table-striped table-responsive table-condensed">
				<tr style="color:black;font-weight:bold;">
					<td></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/OKC.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/ORL.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/PHI.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/PHX.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/POR.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/SAC.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/SAS.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/TOR.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/UTA.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/WAS.svg" alt="Atlanta Hawks"></td>
				</tr>
					<%
						objRS.Open "SELECT * FROM tblGameGrid where tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
					%>
					<%
						While Not objRS.EOF
					%>
					<tr>
					<% if len(objRS.Fields("GameDate")) = 8 then %>
						<td><%=left(objRS.Fields("GameDate").Value,3) %></td>	
					<% elseif len(objRS.Fields("GameDate")) = 9 then %>
						<td><%=left(objRS.Fields("GameDate").Value,4) %></td>					
					<% else %>
						<td><%=left(objRS.Fields("GameDate").Value,5) %></td>					
					<%end if%>
						<td><%=objRS.Fields("OKC").Value %></td>
						<td><%=objRS.Fields("ORL").Value %></td>
						<td><%=objRS.Fields("PHI").Value %></td>
						<td><%=objRS.Fields("PHX").Value %></td>
						<td><%=objRS.Fields("POR").Value %></td>
						<td><%=objRS.Fields("SAC").Value %></td>
						<td><%=objRS.Fields("SAS").Value %></td>
						<td><%=objRS.Fields("TOR").Value %></td>
						<td><%=objRS.Fields("UTA").Value %></td>
						<td><%=objRS.Fields("WAS").Value %></td>
					</tr>

					<%
						objRS.MoveNext
							Wend
					%>
					<%
					objRS.Close
					%>
					<tr style="color:black;font-weight:bold;">
					<td width="10%">GM</td>
						<td>OKC</td>
						<td>ORL</td>
						<td>PHI</td>
						<td>PHX</td>
						<td>POR</td>
						<td>SAC</td>
						<td>SAS</td>
						<td>TOR</td>
						<td>UTA</td>
						<td>WAS</td>	
					</tr>
					<tr class="warning">
						<td>ToT</td>
						<td><%=objrsOKC.RecordCount%></td>
						<td><%=objrsORL.RecordCount%></td>
						<td><%=objrsPHI.RecordCount%></td>
						<td><%=objrsPHX.RecordCount%></td>
						<td><%=objrsPOR.RecordCount%></td>
						<td><%=objrsSAC.RecordCount%></td>
						<td><%=objrsSAS.RecordCount%></td>
						<td><%=objrsTOR.RecordCount%></td>
						<td><%=objrsUTA.RecordCount%></td>
						<td><%=objrsWAS.RecordCount%></td>
					</tr>
				</table>
			</div>
			</div>
			<!--#include virtual="Common/gamegridCloseIO.inc"-->
			<div id="semis" class="tab-pane fade">
			<!--#include virtual="Common/gamegridSemis.inc"-->
			<div class="col-md-4">
			<table class="table table-bordered table-striped table-responsive table-condensed">
				<tr style="color:black;font-weight:bold;">
					<td></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/ATL.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/BKN.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/BOS.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/CHA.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/CHI.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/CLE.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/DAL.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/DEN.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/DET.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/GSW.svg" alt="Atlanta Hawks"></td>
				</tr>
				<%
					objRS.Open "SELECT * FROM tblGameGrid where tblGameGrid.GameDate >= cdate('"&dSemiDate&"') and tblGameGrid.gamedate < cdate('"&dFinalDate&"') ", objConn,3,3,1

				%>
				<%
					While Not objRS.EOF
				%>
				<tr>
					<% if len(objRS.Fields("GameDate")) = 8 then %>
						<td><%=left(objRS.Fields("GameDate").Value,3) %></td>	
					<% elseif len(objRS.Fields("GameDate")) = 9 then %>
						<td><%=left(objRS.Fields("GameDate").Value,4) %></td>					
					<% else %>
						<td><%=left(objRS.Fields("GameDate").Value,5) %></td>					
					<%end if%>
					<td><%=objRS.Fields("ATL").Value %></td>
					<td><%=objRS.Fields("BKN").Value %></td>
					<td><%=objRS.Fields("BOS").Value %></td>
					<td><%=objRS.Fields("CHA").Value %></td>
					<td><%=objRS.Fields("CHI").Value %></td>
					<td><%=objRS.Fields("CLE").Value %></td>
					<td><%=objRS.Fields("DAL").Value %></td>
					<td><%=objRS.Fields("DEN").Value %></td>
					<td><%=objRS.Fields("DET").Value %></td>
					<td><%=objRS.Fields("GSW").Value %></td>
				</tr>
				<%
					objRS.MoveNext
						Wend
				%>
				<%
				objRS.Close
				%>
				<tr style="color:black;font-weight:bold;">
					<td width="10%">GM</td>
					<td>ATL</td>
					<td>BKN</td>
					<td>BOS</td>
					<td>CHA</td>
					<td>CHI</td>
					<td>CLE</td>
					<td>DAL</td>
					<td>DEN</td>
					<td>DET</td>
					<td>GSW</td>
				</tr>
				<tr class="warning">
					<td>ToT</td>
					<td><%=objrsATL.RecordCount%></td>
					<td><%=objrsBKN.RecordCount%></td>
					<td><%=objrsBOS.RecordCount%></td>
					<td><%=objrsCHA.RecordCount%></td>
					<td><%=objrsCHI.RecordCount%></td>
					<td><%=objrsCLE.RecordCount%></td>
					<td><%=objrsDAL.RecordCount%></td>
					<td><%=objrsDEN.RecordCount%></td>
					<td><%=objrsDET.RecordCount%></td>
					<td><%=objrsGSW.RecordCount%></td>
				</tr>
			</table>
			</br>
    </div>
		<div class="col-md-4">
			<table class="table table-bordered table-striped table-responsive table-condensed">
				<tr style="color:black;font-weight:bold;">
					<td></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/HOU.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/IND.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/LAC.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/LAL.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MEM.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MIA.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MIL.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MIN.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/NOP.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/NYK.svg" alt="Atlanta Hawks"></td>
			</tr>
			<%
				objRS.Open "SELECT * FROM tblGameGrid where tblGameGrid.GameDate >=  cdate('"&dSemiDate&"') and tblGameGrid.gamedate < cdate('"&dFinalDate&"') ", objConn,3,3,1
			%>
			<%
				While Not objRS.EOF
			%>
			<tr>
					<% if len(objRS.Fields("GameDate")) = 8 then %>
						<td><%=left(objRS.Fields("GameDate").Value,3) %></td>	
					<% elseif len(objRS.Fields("GameDate")) = 9 then %>
						<td><%=left(objRS.Fields("GameDate").Value,4) %></td>					
					<% else %>
						<td><%=left(objRS.Fields("GameDate").Value,5) %></td>					
					<%end if%>
				<td ><%=objRS.Fields("HOU").Value %></td>
				<td ><%=objRS.Fields("IND").Value %></td>
				<td ><%=objRS.Fields("LAC").Value %></td>
				<td><%=objRS.Fields("LAL").Value %></td>
				<td><%=objRS.Fields("MEM").Value %></td>
				<td><%=objRS.Fields("MIA").Value %></td>
				<td><%=objRS.Fields("MIL").Value %></td>
				<td><%=objRS.Fields("MIN").Value %></td>
				<td><%=objRS.Fields("NOP").Value %></td>
				<td><%=objRS.Fields("NYK").Value %></td>
			</tr>
				<%
					objRS.MoveNext
						Wend
				%>
				<%
				objRS.Close
				%>
			<tr style="color:black;font-weight:bold;">
					<td width="10%">GM</td>
					<td>HOU</td>
					<td>IND</td>
					<td>LAC</td>
					<td>LAL</td>
					<td>MEM</td>
					<td>MIA</td>
					<td>MIL</td>
					<td>MIN</td>
					<td>NOP</td>
					<td>NYK</td>
				</tr>
				<tr class="warning">
					<td>ToT</td>
					<td><%=objrsHOU.RecordCount%></td>
					<td><%=objrsIND.RecordCount%></td>
					<td><%=objrsLAC.RecordCount%></td>
					<td><%=objrsLAL.RecordCount%></td>
					<td><%=objrsMEM.RecordCount%></td>
					<td><%=objrsMIA.RecordCount%></td>
					<td><%=objrsMIL.RecordCount%></td>
					<td><%=objrsMIN.RecordCount%></td>
					<td><%=objrsNOP.RecordCount%></td>
					<td><%=objrsNYK.RecordCount%></td>
				</tr>
			</table>
			</br>
    </div>
		<div class="col-md-4">
			<table class="table table-bordered table-striped table-responsive table-condensed">
				<tr style="color:black;font-weight:bold;">
					<td></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/OKC.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/ORL.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/PHI.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/PHX.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/POR.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/SAC.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/SAS.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/TOR.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/UTA.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/WAS.svg" alt="Atlanta Hawks"></td>
				</tr>
				<%
					objRS.Open "SELECT * FROM tblGameGrid where tblGameGrid.GameDate >=  cdate('"&dSemiDate&"') and tblGameGrid.gamedate < cdate('"&dFinalDate&"') ", objConn,3,3,1
				%>
				<%
					While Not objRS.EOF
				%>
				<tr>
					<% if len(objRS.Fields("GameDate")) = 8 then %>
						<td><%=left(objRS.Fields("GameDate").Value,3) %></td>	
					<% elseif len(objRS.Fields("GameDate")) = 9 then %>
						<td><%=left(objRS.Fields("GameDate").Value,4) %></td>					
					<% else %>
						<td><%=left(objRS.Fields("GameDate").Value,5) %></td>					
					<%end if%>
					<td><%=objRS.Fields("OKC").Value %></td>
					<td><%=objRS.Fields("ORL").Value %></td>
					<td><%=objRS.Fields("PHI").Value %></td>
					<td><%=objRS.Fields("PHX").Value %></td>
					<td><%=objRS.Fields("POR").Value %></td>
					<td><%=objRS.Fields("SAC").Value %></td>
					<td><%=objRS.Fields("SAS").Value %></td>
					<td><%=objRS.Fields("TOR").Value %></td>
					<td><%=objRS.Fields("UTA").Value %></td>
					<td><%=objRS.Fields("WAS").Value %></td>
				</tr>

				<%
					objRS.MoveNext
						Wend
				%>
				<%
				objRS.Close
				%>
					<tr style="color:black;font-weight:bold;">
					<td width="10%">GM</td>
					<td>OKC</td>
					<td>ORL</td>
					<td>PHI</td>
					<td>PHX</td>
					<td>POR</td>
					<td>SAC</td>
					<td>SAS</td>
					<td>TOR</td>
					<td>UTA</td>
					<td>WAS</td>	
				</tr>
				<tr class="warning">
					<td>ToT</td>
					<td><%=objrsOKC.RecordCount%></td>
					<td><%=objrsORL.RecordCount%></td>
					<td><%=objrsPHI.RecordCount%></td>
					<td><%=objrsPHX.RecordCount%></td>
					<td><%=objrsPOR.RecordCount%></td>
					<td><%=objrsSAC.RecordCount%></td>
					<td><%=objrsSAS.RecordCount%></td>
					<td><%=objrsTOR.RecordCount%></td>
					<td><%=objrsUTA.RecordCount%></td>
					<td><%=objrsWAS.RecordCount%></td>
				</tr>
			</table>
    </div>			
			</div>
			<!--#include virtual="Common/gamegridCloseIO.inc"-->
			<div id="finals" class="tab-pane fade">
			<!--#include virtual="Common/gamegridFinals.inc"-->
    <div class="col-md-4">
			<table class="table table-bordered table-striped table-responsive table-condensed">
				<tr style="color:black;font-weight:bold;">
					<td></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/ATL.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/BKN.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/BOS.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/CHA.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/CHI.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/CLE.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/DAL.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/DEN.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/DET.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/GSW.svg" alt="Atlanta Hawks"></td>
				</tr>
				<%
					objRS.Open "SELECT * FROM tblGameGrid where tblGameGrid.GameDate >=  cdate('"&dFinalDate&"') ", objConn,3,3,1

				%>
				<%
					While Not objRS.EOF
				%>
				<tr>
					<% if len(objRS.Fields("GameDate")) = 8 then %>
						<td><%=left(objRS.Fields("GameDate").Value,3) %></td>	
					<% elseif len(objRS.Fields("GameDate")) = 9 then %>
						<td><%=left(objRS.Fields("GameDate").Value,4) %></td>					
					<% else %>
						<td><%=left(objRS.Fields("GameDate").Value,5) %></td>					
					<%end if%>
					<td><%=objRS.Fields("ATL").Value %></td>
					<td><%=objRS.Fields("BKN").Value %></td>
					<td><%=objRS.Fields("BOS").Value %></td>
					<td><%=objRS.Fields("CHA").Value %></td>
					<td><%=objRS.Fields("CHI").Value %></td>
					<td><%=objRS.Fields("CLE").Value %></td>
					<td><%=objRS.Fields("DAL").Value %></td>
					<td><%=objRS.Fields("DEN").Value %></td>
					<td><%=objRS.Fields("DET").Value %></td>
					<td><%=objRS.Fields("GSW").Value %></td>
				</tr>
				<%
					objRS.MoveNext
						Wend
				%>
				<%
				objRS.Close
				%>

				<tr style="color:black;font-weight:bold;">
					<td width="10%">GM</td>
					<td>ATL</td>
					<td>BKN</td>
					<td>BOS</td>
					<td>CHA</td>
					<td>CHI</td>
					<td>CLE</td>
					<td>DAL</td>
					<td>DEN</td>
					<td>DET</td>
					<td>GSW</td>
				</tr>
				<tr class="warning">
					<td>ToT</td>
					<td><%=objrsATL.RecordCount%></td>
					<td><%=objrsBKN.RecordCount%></td>
					<td><%=objrsBOS.RecordCount%></td>
					<td><%=objrsCHA.RecordCount%></td>
					<td><%=objrsCHI.RecordCount%></td>
					<td><%=objrsCLE.RecordCount%></td>
					<td><%=objrsDAL.RecordCount%></td>
					<td><%=objrsDEN.RecordCount%></td>
					<td><%=objrsDET.RecordCount%></td>
					<td><%=objrsGSW.RecordCount%></td>
				</tr>
			</table>
			</br>
    </div>
		<div class="col-md-4">
			<table class="table table-bordered table-striped table-responsive table-condensed">
				<tr style="color:black;font-weight:bold;">
					<td></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/HOU.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/IND.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/LAC.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/LAL.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MEM.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MIA.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MIL.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/MIN.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/NOP.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/NYK.svg" alt="Atlanta Hawks"></td>
				</tr>
			<%
				objRS.Open "SELECT * FROM tblGameGrid where tblGameGrid.GameDate >= cdate('"&dFinalDate&"') ", objConn,3,3,1
			%>
			<%
				While Not objRS.EOF
			%>
			<tr>
					<% if len(objRS.Fields("GameDate")) = 8 then %>
						<td><%=left(objRS.Fields("GameDate").Value,3) %></td>	
					<% elseif len(objRS.Fields("GameDate")) = 9 then %>
						<td><%=left(objRS.Fields("GameDate").Value,4) %></td>					
					<% else %>
						<td><%=left(objRS.Fields("GameDate").Value,5) %></td>					
					<%end if%>
				<td><%=objRS.Fields("HOU").Value %></td>
				<td><%=objRS.Fields("IND").Value %></td>
				<td><%=objRS.Fields("LAC").Value %></td>
				<td><%=objRS.Fields("LAL").Value %></td>
				<td><%=objRS.Fields("MEM").Value %></td>
				<td><%=objRS.Fields("MIA").Value %></td>
				<td><%=objRS.Fields("MIL").Value %></td>
				<td><%=objRS.Fields("MIN").Value %></td>
				<td><%=objRS.Fields("NOP").Value %></td>
				<td><%=objRS.Fields("NYK").Value %></td>
			</tr>
				<%
					objRS.MoveNext
						Wend
				%>
				<%
				objRS.Close
				%>
				<tr style="color:black;font-weight:bold;">
					<td>GM</td>
					<td>HOU</td>
					<td>IND</td>
					<td>LAC</td>
					<td>LAL</td>
					<td>MEM</td>
					<td>MIA</td>
					<td>MIL</td>
					<td>MIN</td>
					<td>NOP</td>
					<td>NYK</td>
				</tr>
				<tr class="warning">
					<td>ToT</td>
					<td><%=objrsHOU.RecordCount%></td>
					<td><%=objrsIND.RecordCount%></td>
					<td><%=objrsLAC.RecordCount%></td>
					<td><%=objrsLAL.RecordCount%></td>
					<td><%=objrsMEM.RecordCount%></td>
					<td><%=objrsMIA.RecordCount%></td>
					<td><%=objrsMIL.RecordCount%></td>
					<td><%=objrsMIN.RecordCount%></td>
					<td><%=objrsNOP.RecordCount%></td>
					<td><%=objrsNYK.RecordCount%></td>
				</tr>
			</table>
			</br>
    </div>
		<div class="col-md-4">
			<table class="table table-bordered table-striped table-responsive table-condensed">
				<tr style="color:black;font-weight:bold;">
					<td></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/OKC.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/ORL.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/PHI.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/PHX.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/POR.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/SAC.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/SAS.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/TOR.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/UTA.svg" alt="Atlanta Hawks"></td>
					<td><img class="logo" src="//i.cdn.turner.com/nba/nba/assets/logos/teams/primary/web/WAS.svg" alt="Atlanta Hawks"></td>
				</tr>
				<%
					objRS.Open "SELECT * FROM tblGameGrid where tblGameGrid.GameDate >= cdate('"&dFinalDate&"') ", objConn,3,3,1
				%>
				<%
					While Not objRS.EOF
				%>
				<tr>
					<% if len(objRS.Fields("GameDate")) = 8 then %>
						<td><%=left(objRS.Fields("GameDate").Value,3) %></td>	
					<% elseif len(objRS.Fields("GameDate")) = 9 then %>
						<td><%=left(objRS.Fields("GameDate").Value,4) %></td>					
					<% else %>
						<td><%=left(objRS.Fields("GameDate").Value,5) %></td>					
					<%end if%>
					<td><%=objRS.Fields("OKC").Value %></td>
					<td><%=objRS.Fields("ORL").Value %></td>
					<td><%=objRS.Fields("PHI").Value %></td>
					<td><%=objRS.Fields("PHX").Value %></td>
					<td><%=objRS.Fields("POR").Value %></td>
					<td><%=objRS.Fields("SAC").Value %></td>
					<td><%=objRS.Fields("SAS").Value %></td>
					<td><%=objRS.Fields("TOR").Value %></td>
					<td><%=objRS.Fields("UTA").Value %></td>
					<td><%=objRS.Fields("WAS").Value %></td>
				</tr>

				<%
					objRS.MoveNext
						Wend
				%>
				<%
				objRS.Close
				%>
				<tr style="color:black;font-weight:bold;">
					<td width="10%">GM</td>
					<td>OKC</td>
					<td>ORL</td>
					<td>PHI</td>
					<td>PHX</td>
					<td>POR</td>
					<td>SAC</td>
					<td>SAS</td>
					<td>TOR</td>
					<td>UTA</td>
					<td>WAS</td>	
				</tr>
				<tr class="warning">
					<td>ToT</td>
					<td><%=objrsOKC.RecordCount%></td>
					<td><%=objrsORL.RecordCount%></td>
					<td><%=objrsPHI.RecordCount%></td>
					<td><%=objrsPHX.RecordCount%></td>
					<td><%=objrsPOR.RecordCount%></td>
					<td><%=objrsSAC.RecordCount%></td>
					<td><%=objrsSAS.RecordCount%></td>
					<td><%=objrsTOR.RecordCount%></td>
					<td><%=objrsUTA.RecordCount%></td>
					<td><%=objrsWAS.RecordCount%></td>
				</tr>
			</table>
    </div>			
			</div>
			<!--#include virtual="Common/gamegridCloseIO.inc"-->
		</div>
	</div>
</div>		
<%
objrs.close
objrsgames.close
objconn.close
objrsATL.close
objrsBKN.close      
objrsBOS.close      
objrsCHA.close      
objrsCHI.close      
objrsCLE.close      
objrsDAL.close      
objrsDEN.close      
objrsDET.close      
objrsGSW.close      
objrsHOU.close      
objrsIND.close      
objrsLAC.close      
objrsLAL.close      
objrsMEM.close      
objrsMIA.close      
objrsMIL.close      
objrsMIN.close      
objrsNOP.close      
objrsNYK.close      
objrsOKC.close      
objrsORL.close      
objrsPHI.close      
objrsPHX.close      
objrsPOR.close      
objrsSAC.close      
objrsSAS.close      
objrsTOR.close      
objrsUTA.close      
objrsWAS.close      	

Set objrsATL = Nothing
Set objrsBKN = Nothing      
Set objrsBOS = Nothing      
Set objrsCHA = Nothing      
Set objrsCHI = Nothing      
Set objrsCLE = Nothing      
Set objrsDAL = Nothing      
Set objrsDEN = Nothing      
Set objrsDET = Nothing      
Set objrsGSW = Nothing      
Set objrsHOU = Nothing      
Set objrsIND = Nothing      
Set objrsLAC = Nothing      
Set objrsLAL = Nothing      
Set objrsMEM = Nothing      
Set objrsMIA = Nothing      
Set objrsMIL = Nothing      
Set objrsMIN = Nothing      
Set objrsNOP = Nothing      
Set objrsNYK = Nothing      
Set objrsOKC = Nothing      
Set objrsORL = Nothing      
Set objrsPHI = Nothing      
Set objrsPHX = Nothing      
Set objrsPOR = Nothing      
Set objrsSAC = Nothing      
Set objrsSAS = Nothing      
Set objrsTOR = Nothing      
Set objrsUTA = Nothing      
Set objrsWAS = Nothing 
Set objrs = Nothing
Set objrsgames = Nothing
Set objConn = Nothing
Session.CodePage = Session("FP_OldCodePage")
Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>