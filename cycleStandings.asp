<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

 	Dim objRSgames,objRS,objConn, objRSWork, ownerid
	Dim strSQL, iPlayerClaimed,objRSTxns, objRSOwners, objRejectWaivers, iPlayerWaived, iOwner, w_action
	
	Set objConn       = Server.CreateObject("ADODB.Connection")
	Set objRS         = Server.CreateObject("ADODB.RecordSet")
	Set objRSgames    = er.CreateObject("ADODB.RecordSet")
	Set objRSWork     = Server.CreateObject("ADODB.RecordSet")

	objConn.Open Application("lineupstest_ConnectionString")
	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"
	%>
<!--#include virtual="Common/SESSION.inc"-->
<!DOCTYPE HTML>
<html lang="en">
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<link href="css/tabs.css" rel="stylesheet">
<!--#include virtual="Common/headerMain.inc"-->
<style type="text/css">
.bs-example{
		margin: 20px;
}
tr{
	background-color:white;
	text-align:center;
}
.nav-tabs {
    border-bottom: 2px solid black;
}
</style>
</head>
<body>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Cycle Standings</strong>
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="bs-example">
					<ul class="nav nav-tabs">
							<li class="active"><a data-toggle="tab" href="#cycle1">1</a></li>
							<li><a data-toggle="tab" href="#cycle2">2</a></li>
							<li><a data-toggle="tab" href="#cycle3">3</a></li>        
							<li><a data-toggle="tab" href="#cycle4">4</a></li>   
							<li><a data-toggle="tab" href="#cycle5">5</a></li>
							<li><a data-toggle="tab" href="#cycle6">6</a></li>
							<li><a data-toggle="tab" href="#cycle7">7</a></li>				
					</ul>
					<div class="tab-content">
						<div id="cycle1" class="tab-pane fade in active">
							<h3>Cycle 1</h3>
							<table class="table table-custom-black table-responsive table-bordered table-condensed">
								<thead>
									<tr class="big" style="background-color:yellowgreen;color:black;font-weight:bold;">
										<th  class="big"  style="text-align:left" width="10%">Team</th>
										<th  class="big" width="10%">W</th>
										<th  class="big" width="10%">L</th>
										<th  class="big" width="10%">PPG</th>
										<th  class="big" width="10%">OPPG</th>
									</tr>
								</thead>
								<%
									objRS.Open   "SELECT * FROM Standings_Cycle where cycle = 1 order by rank asc ", objConn,3,3,1
								%>
								<%
									While Not objRS.EOF
								%>
								
									<tr class="big" style="text-align:center;vertical-align:middle;">
										<td class="big" class="content" style="text-align:left;color;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
								</tr>
									<%
									objRS.MoveNext
										Wend
									%>
									<%
									objRS.Close
									%>
							</table>	
						</div>
						<div id="cycle2" class="tab-pane fade">
							<h3>Cycle 2</h3>
							<table class="table table-custom-black table-responsive table-bordered table-condensed">
								<thead>
									<tr class="big" style="background-color:yellowgreen;color:black;font-weight:bold;">
										<th  class="big"  style="text-align:left" width="10%">Team</th>
										<th  class="big" width="10%">W</th>
										<th  class="big" width="10%">L</th>
										<th  class="big" width="10%">PPG</th>
										<th  class="big" width="10%">OPPG</th>
									</tr>
								</thead>
								<%
									objRS.Open   "SELECT * FROM Standings_Cycle where cycle = 2 order by rank asc ", objConn,3,3,1
								%>
								<%
									While Not objRS.EOF
								%>
									<tr class="big" style="text-align:center;vertical-align:middle;">
										<td class="big" class="content" style="text-align:left;color;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
									</tr>
									<%
									objRS.MoveNext
										Wend
									%>
									<%
									objRS.Close
									%>
							</table>
						</div>
						<div id="cycle3" class="tab-pane fade">
							<h3>Cycle 3</h3>
							<table class="table table-custom-black table-responsive table-bordered table-condensed">
								<thead>
									<tr class="big" style="background-color:yellowgreen;color:black;font-weight:bold;">
										<th  class="big"  style="text-align:left" width="10%">Team</th>
										<th  class="big" width="10%">W</th>
										<th  class="big" width="10%">L</th>
										<th  class="big" width="10%">PPG</th>
										<th  class="big" width="10%">OPPG</th>
									</tr>
								</thead>
								<%
									objRS.Open   "SELECT * FROM Standings_Cycle where cycle = 3 order by rank asc ", objConn,3,3,1
								%>
								<%
									While Not objRS.EOF
								%>
									<tr class="big" style="text-align:center;vertical-align:middle;">
										<td class="big" class="content" style="text-align:left;color;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
								</tr>
									<%
									objRS.MoveNext
										Wend
									%>
									<%
									objRS.Close
									%>
							</table>			
						</div>
						<div id="cycle4" class="tab-pane fade">
							<h3>Cycle 4</h3>
							<table class="table table-custom-black table-responsive table-bordered table-condensed">
								<thead>
									<tr class="big" style="background-color:yellowgreen;color:black;font-weight:bold;">
										<th  class="big"  style="text-align:left" width="10%">Team</th>
										<th  class="big" width="10%">W</th>
										<th  class="big" width="10%">L</th>
										<th  class="big" width="10%">PPG</th>
										<th  class="big" width="10%">OPPG</th>
									</tr>
								</thead>
								<%
									objRS.Open   "SELECT * FROM Standings_Cycle where cycle = 4 order by rank asc ", objConn,3,3,1
								%>
								<%
									While Not objRS.EOF
								%>
								
									<tr class="big" style="text-align:center;vertical-align:middle;">
										<td class="big" class="content" style="text-align:left;color;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
								</tr>
									<%
									objRS.MoveNext
										Wend
									%>
									<%
									objRS.Close
									%>
							</table>			
						</div>
						<div id="cycle5" class="tab-pane fade">
							<h3>Cycle 5</h3>
							<table class="table table-custom-black table-responsive table-bordered table-condensed">
								<thead>
									<tr class="big" style="background-color:yellowgreen;color:black;font-weight:bold;">
										<th  class="big"  style="text-align:left" width="10%">Team</th>
										<th  class="big" width="10%">W</th>
										<th  class="big" width="10%">L</th>
										<th  class="big" width="10%">PPG</th>
										<th  class="big" width="10%">OPPG</th>
									</tr>
								</thead>
								<%
									objRS.Open   "SELECT * FROM Standings_Cycle where cycle = 5 order by rank asc ", objConn,3,3,1
								%>
								<%
									While Not objRS.EOF
								%>
								
									<tr class="big" style="text-align:center;vertical-align:middle;">
										<td class="big" class="content" style="text-align:left;color;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
								</tr>
									<%
									objRS.MoveNext
										Wend
									%>
									<%
									objRS.Close
									%>
							</table>			
						</div>
						<div id="cycle6" class="tab-pane fade">
							<h3>Cycle 6</h3>
								<table class="table table-custom-black table-responsive table-bordered table-condensed">
								<thead>
									<tr class="big" style="background-color:yellowgreen;color:black;font-weight:bold;">
										<th  class="big"  style="text-align:left" width="10%">Team</th>
										<th  class="big" width="10%">W</th>
										<th  class="big" width="10%">L</th>
										<th  class="big" width="10%">PPG</th>
										<th  class="big" width="10%">OPPG</th>
									</tr>
								</thead>
								<%
									objRS.Open   "SELECT * FROM Standings_Cycle where cycle = 6 order by rank asc ", objConn,3,3,1
								%>
								<%
									While Not objRS.EOF
								%>
								
									<tr class="big" style="text-align:center;vertical-align:middle;">
										<td class="big" class="content" style="text-align:left;color;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
								</tr>
									<%
									objRS.MoveNext
										Wend
									%>
									<%
									objRS.Close
									%>
							</table>		
						</div>
						<div id="cycle7" class="tab-pane fade">
							<h3>Cycle 7</h3>
							<table class="table table-custom-black table-responsive table-bordered table-condensed">
								<thead>
									<tr class="big" style="background-color:yellowgreen;color:black;font-weight:bold;">
										<th  class="big"  style="text-align:left" width="10%">Team</th>
										<th  class="big" width="10%">W</th>
										<th  class="big" width="10%">L</th>
										<th  class="big" width="10%">PPG</th>
										<th  class="big" width="10%">OPPG</th>
									</tr>
								</thead>
								<%
									objRS.Open   "SELECT * FROM Standings_Cycle where cycle = 7 order by rank asc ", objConn,3,3,1
								%>
								<%
									While Not objRS.EOF
								%>
								
									<tr class="big" style="text-align:center;vertical-align:middle;">
										<td class="big" class="content" style="text-align:left;color;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
								</tr>
									<%
									objRS.MoveNext
										Wend
									%>
									<%
									objRS.Close
									%>
							</table>			
						</div>
				</div>
			</div>
		</div>
	</div>
</div>
</body>
</html>                            