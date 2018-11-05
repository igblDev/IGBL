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
	Set objRSgames    = Server.CreateObject("ADODB.RecordSet")
	Set objRSWork     = Server.CreateObject("ADODB.RecordSet")

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
<!--#include virtual="Common/headerMain.inc"-->
<style>
darkorange {
	color: darkorange;
}

th {
    vertical-align: middle;
  	text-align: center;}

td {
    vertical-align: middle;
    text-align: center;
}

td.h2h {
    vertical-align: middle;
    text-align: center;
}
tr.h2hrow {
    vertical-align: middle;
    text-align: center;	
}
black {
	color:black;
		
	text-transform: uppercase;
}

a.blue {
	color: white;
	text-decoration: underline;
	font-weight:bold;	
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
				<strong>Standings</strong>
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<ul class="nav nav-tabs">
				<li class="active"><a data-toggle="tab" href="#standings"><i class="fal fa-sort-numeric-up"></i>&nbsp;Regular</a></li>
				<li><a data-toggle="tab" href="#h2h"><i class="fas fa-user"></i>&nbsp;H2H</a></li>
				
				<%
					objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'PLAYOFF_START_DATE' ",objConn
					wPlayoffStart = objRSWork.Fields("param_date").Value
					objRSWork.Close
					if (date() >= wPlayoffStart) then 
					%>
						<li><a data-toggle="tab" href="#postandings">Play-Offs</a></li>
					<%else%>
						<li><a data-toggle="tab" href="#cycle"><i class="fal fa-sort-numeric-up"></i>&nbsp;Cycle</a></li>					
					<% end if %>
				
				<li><a data-toggle="tab" href="#prs"><i class="fab fa-superpowers"></i>&nbsp;Power Rankings</a></li>
			</ul>
		</div>
	</div>
	</br>
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="tab-content">
				<div id="standings" class="tab-pane fade in active">
					<table class="table table-custom-black table-responsive table-bordered table-condensed">
						<thead>
							<tr class="h2h big" style="background-color:yellowgreen;color:black;font-weight:bold;">
								<th  class="h2h big"  style="text-align:left" width="10%">Team</th>
								<th  class="h2h big" width="10%">W</th>
								<th  class="h2h big" width="10%">L</th>
								<th  class="h2h big" width="10%">GB</th>
								<th  class="h2h big" width="10%">PPG</th>
								<th  class="h2h big" width="10%">OPPG</th>
								<th  class="h2h big" width="10%">DIFF</th>
								<th  class="h2h big" width="10%">STR</th>
								<th  class="h2h big" width="10%">L10</th>
								<th  class="h2h big hidden-xs"  width="10%">P</th>
							</tr>
						</thead>
						<%
							objRS.Open   "SELECT * FROM qryStandings ", objConn,3,3,1
						%>
						<%
							While Not objRS.EOF
						%>
						<% if objRS.Fields("rank")= 1 or objRS.Fields("rank")= 2 then%>
							<%if objRS.Fields("ownerID").Value = ownerid then%>
							<tr class="success big" style="font-weight:bold;text-align:center;vertical-align:middle;">
								<td class="h2h big" class="content" style="text-align:left;color;vertical-align:middle;">
									<small><%=objRS.Fields("rank").Value %></small>&nbsp;<bye><%=objRS.Fields("shortname").Value %></bye>
								</td>
							<%else%>
							<tr style="background-color:white;">
									<td class="h2h big" class="content" style="text-align:left;color;vertical-align:middle;">
									<small><%=objRS.Fields("rank").Value %></small>&nbsp;<bye><%=objRS.Fields("shortname").Value %></bye>							
									</td>
							<%end if%>
						<!--3rd and 6th Seeds  -->
						<%elseif objRS.Fields("rank")= 3 or objRS.Fields("rank")= 4 or objRS.Fields("rank")= 5 or  objRS.Fields("rank")= 6 then%>
							<%if objRS.Fields("ownerID").Value = ownerid then%>
								<tr class="success big" style="font-weight:bold;text-align:center;vertical-align:middle;">					
									<td class="h2h big" class="content" style="text-align:left;vertical-align:middle;">
										<small><%=objRS.Fields("rank").Value %></small>&nbsp;<playoffs><%=objRS.Fields("shortname").Value %></playoffs>
									</td>
							<%else%>
									<tr style="background-color:white;">					
										<td class="h2h big" class="content" style="text-align:left;vertical-align:middle;">
											<small><%=objRS.Fields("rank").Value %></small>&nbsp;<playoffs><%=objRS.Fields("shortname").Value %></playoffs>		
										</td>
							<%end if%>	
						<%else%>
						<%if objRS.Fields("ownerID").Value = ownerid then%>
							<tr class="success big" style="font-weight:bold;text-align:left;vertical-align:middle;">	
							<td class="h2h big" class="content" style="text-align:left;vertical-align:middle;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<nextyear><red><%=objRS.Fields("shortname").Value %></red></nextyear>
				
							</td>
						<%else%>
							<tr style="background-color:white;">	
								<td class="h2h big" class="content" style="text-align:left;vertical-align:middle;">
									<small><%=objRS.Fields("rank").Value %></small>&nbsp;<nextyear><red><%=objRS.Fields("shortname").Value %></red></nextyear>
								</td>
						<%end if%>
					<%end if%>
						<td  class="h2h big" style="text-align:center;vertical-align:middle;"><%=objRS.Fields("won").Value %></td>
						<td  class="h2h big" style="text-align:center;vertical-align:middle;"><%=objRS.Fields("loss").Value %></td>
						<td  class="h2h big" style="text-align:center;vertical-align:middle;"><%=objRS.Fields("gb").Value %></td>
						<td  class="h2h big" style="text-align:center;vertical-align:middle;"><%=round(objRS.Fields("ppg").Value,1) %></td>
						<td  class="h2h big" style="text-align:center;vertical-align:middle;"><%=round(objRS.Fields("oppg").Value,1) %></td>
						<td  class="h2h big" style="text-align:center;vertical-align:middle;"><%=round(objRS.Fields("diff").Value,1) %></td>
						<td  class="h2h big" style="text-align:center;vertical-align:middle;"><%=objRS.Fields("st").Value %></td>
						<td  class="h2h big" style="text-align:center;vertical-align:middle;"><%=objRS.Fields("cycle").Value %></td>

						<%if objRS.Fields("LP").Value > 0 then %>
						<td class="h2h big hidden-xs" style="text-align:center;vertical-align:middle;color:#9a1400;" nowrap><%=objRS.Fields("LP").Value %></td>
						<%else%>
						<td class="h2h big hidden-xs" style="text-align:center;vertical-align:middle;" ><%=objRS.Fields("LP").Value %></td>
						<%end if%>
						</tr>
							<%
							objRS.MoveNext
								Wend
							%>
							<%
							objRS.Close
							%>
							</table>
					<div>
						<h4><span style="color:#9a1400;font-weight:bold;text-decoration:underline;" class="big">Tie Breaker Rules for H2H Standings</span></h4>
						<ol class="big">
							<li><small><strong>The Rules Below Enforced on Position Round!</strong></small></li>
							<li>If multiple teams share identical records, it's determined if an owner has a H2H advantage over all tied teams. If so that team will be placed highest in the standings</li>
							<li>If Step 2 does not render a decision, the team with the highest PPG will receive the highest ranking of all tied teams</li>
							<li>Repeat Step 2 until all tied teams are ranked</li>
						</ol>
					</div>
			</div>



				<div id="h2h" class="tab-pane">
				<table class="table table-custom-black table-responsive table-bordered table-condensed">
				<thead>
				<tr class="h2h big" style="background-color:yellowgreen;color:black;font-weight:bold;">
						<th  class="h2h big" nowrap  style="text-align:left" width="12%">Team</th>
						<th  class="h2h big"  width="8%">GR</th>
						<th  class="h2h big"  width="8%">JP</th>
						<th  class="h2h big"  width="8%">DB</th>
						<th  class="h2h big"  width="8%">CJ</th>
						<th  class="h2h big"  width="8%">JW</th>
						<th  class="h2h big"  width="8%">MJ</th>
						<th  class="h2h big"  width="8%">AW</th>
						<th  class="h2h big"  width="8%">FC</th>
						<th  class="h2h big"  width="8%">TA</th>
						<th  class="h2h big"  width="8%">DM</th>
						<th  class="h2h big hidden-xs"  width="8%">P</th>
						</tr>
				</thead>
				<%
					objRS.Open   "SELECT * FROM qryStandings ", objConn,3,3,1
				%>
				<%
					While Not objRS.EOF
				%>
				<% if objRS.Fields("ownerID").Value = ownerid then %>
				<tr class="success h2hrow" style="font-weight:bold;text-align:center;vertical-align:middle;">	
					<td class="h2h big" style="text-align:left;text-transform: uppercase;"><%=objRS.Fields("shortname").Value %></th></td>
					<td class="h2h big"><%=objRS.Fields("GR").Value %></td>
					<td class="h2h big"><%=objRS.Fields("JP").Value %></td>
					<td class="h2h big"><%=objRS.Fields("DB").Value %></td>
					<td class="h2h big"><%=objRS.Fields("CJ").Value %></td>
					<td class="h2h big"><%=objRS.Fields("JW").Value %></td>
					<td class="h2h big"><%=objRS.Fields("MJ").Value %></td>
					<td class="h2h big"><%=objRS.Fields("AW").Value %></td>
					<td class="h2h big"><%=objRS.Fields("FC").Value %></td>
					<td class="h2h big"><%=objRS.Fields("TA").Value %></td>
					<td class="h2h big"><%=objRS.Fields("DM").Value %></td>
					<%if objRS.Fields("LP").Value > 0 then %>
					<td class="h2h big hidden-xs" style="color:#9a1400;" nowrap><%=objRS.Fields("LP").Value %></td>
					<%else%>
					<td class="h2h big hidden-xs" nowrap><%=objRS.Fields("LP").Value %></td>
					<%end if%>
					</tr>
				<%else%>
					<tr class="h2hrow" style="background-color:white;">
					<td  class="h2h big" style="text-align:left;text-transform: uppercase;"><%=objRS.Fields("shortname").Value %></th></td>
					<td  class="h2h big"><%=objRS.Fields("GR").Value %></td>
					<td  class="h2h big"><%=objRS.Fields("JP").Value %></td>
					<td  class="h2h big"><%=objRS.Fields("DB").Value %></td>
					<td  class="h2h big"><%=objRS.Fields("CJ").Value %></td>
					<td  class="h2h big"><%=objRS.Fields("JW").Value %></td>
					<td  class="h2h big"><%=objRS.Fields("MJ").Value %></td>
					<td  class="h2h big"><%=objRS.Fields("AW").Value %></td>
					<td  class="h2h big"><%=objRS.Fields("FC").Value %></td>
					<td  class="h2h big"><%=objRS.Fields("TA").Value %></td>
					<td  class="h2h big"><%=objRS.Fields("DM").Value %></td>
					<%if objRS.Fields("LP").Value > 0 then %>
					<td class="h2h big hidden-xs" style="color:#9a1400;" nowrap><%=objRS.Fields("LP").Value %></td>
					<%else%>
					<td  class="h2h big hidden-xs"><%=objRS.Fields("LP").Value %></td>
					<%end if%>
					</tr>
				<% end if%>
				<%
					objRS.MoveNext
						Wend
				%>
				<%
				objRS.Close
				%>
			</table>				
				</div>
				<div id="postandings" class="tab-pane">
					<div class="row">
						<div class="col-md-12 col-sm-12 col-xs-12">
							<ul class="nav nav-tabs">
								<li class="active"><a data-toggle="tab" href="#qtrs">Quarters</a></li>
								<li><a data-toggle="tab" href="#semis">Semis</a></li>
								<li><a data-toggle="tab" href="#finals">Finals</a></li>
							</ul>
						</div>
					</div>
					<%
				 	Dim objRSRnd1,objRSRnd2,objRSRnd3,objRSwaivers 
		
					Set objRSRnd1    = Server.CreateObject("ADODB.RecordSet")
					Set objRSRnd2    = Server.CreateObject("ADODB.RecordSet")
					Set objRSRnd3    = Server.CreateObject("ADODB.RecordSet")

					%>
					<div class="row">
						<div class="col-md-12 col-sm-12 col-xs-12">
							<div class="tab-content">
								<div id="qtrs" class="tab-pane fade in active">
								<table class="table table-custom-black table-responsive table-bordered table-condensed">
								<thead>
							<tr style="background-color:yellowgreen;color:black;font-weight:bold;">
								<th  class="big"  style="text-align:left" width="25%">Team</th>
								<th  class="big"  width="15%">W</th>
								<th  class="big"  width="15%">L</th>
								<th  class="big"  width="15%">PPG</th>
								<th  class="big"  width="15%">OPPG</th>
								<th  class="big"  width="15%">PD</th>
								</tr>
								</thead>
								<%
								objRSRnd1.Open   "SELECT x.*, UCASE(y.shortname) as shortname FROM Standings_RD1 x, TBLOwners y where x.id = y.ownerid ", objConn,3,3,1
								%>
								<%
								While Not objRSRnd1.EOF
								%>
								<tr style="background-color:white;">
								<td class="big" class="content" style="text-align:left;color;font-weight:700;"><%=objRSRnd1.Fields("shortname").Value %></td>
								<td  class="big"><%=objRSRnd1.Fields("won").Value %></td>
								<td  class="big"><%=objRSRnd1.Fields("loss").Value %></td>
								<td  class="big"><%=round(objRSRnd1.Fields("ppg").Value,1) %></td>
								<td  class="big"><%=round(objRSRnd1.Fields("oppg").Value,1) %></td>
								<td  class="big"><%=round(objRSRnd1.Fields("diff").Value,1) %></td>
								</tr>
								<%
								objRSRnd1.MoveNext
								Wend
								%>
								<%
								objRSRnd1.Close
								%>
								</table>
								</div>
								<div id="semis" class="tab-pane fade">
								<table class="table table-custom table-bordered table-condensed">
								<thead>
								<tr style="background-color:yellowgreen;color:black;font-weight:bold;">
								<td  class="big"  style="text-align:left" width="25%">Team</td>
								<td  class="big"  width="15%">W</td>
								<td  class="big"  width="15%">L</td>
								<td  class="big"  width="15%">PPG</td>
								<td  class="big"  width="15%">OPPG</td>
								<td  class="big"  width="15%">PD</td>
								</tr>
								</thead>
								<%
								objRSRnd2.Open   "SELECT x.*, UCASE(y.shortname) as shortname FROM Standings_RD2 x, TBLOwners y where x.id = y.ownerid  ", objConn,3,3,1
								%>
								<%
								While Not objRSRnd2.EOF
								%>
								<tr style="background-color:white;">
								<td class="big" class="content" style="text-align:left;color;font-weight:700;"><%=objRSRnd2.Fields("shortname").Value %></td>
								<td  class="big"><%=objRSRnd2.Fields("won").Value %></td>
								<td  class="big"><%=objRSRnd2.Fields("loss").Value %></td>
								<td  class="big"><%=round(objRSRnd2.Fields("ppg").Value,1) %></td>
								<td  class="big"><%=round(objRSRnd2.Fields("oppg").Value,1) %></td>
								<td  class="big"><%=round(objRSRnd2.Fields("diff").Value,1) %></td>
								</tr>
								<%
								objRSRnd2.MoveNext
								Wend
								%>
								<%
								objRSRnd2.Close
								%>
								</table>
								</div>
								<div id="finals" class="tab-pane fade">
								<table class="table table-custom table-bordered table-condensed">
								<thead>
									<tr style="background-color:yellowgreen;color:black;font-weight:bold;">
								<td  class="big"  style="text-align:left" width="25%">Team</td>
								<td  class="big"  width="15%">W</td>
								<td  class="big"  width="15%">L</td>
								<td  class="big"  width="15%">PPG</td>
								<td  class="big"  width="15%">OPPG</td>
								<td  class="big"  width="15%">PD</td>
								</tr>
								</thead>
								<%
								objRSRnd3.Open   "SELECT x.*, UCASE(y.shortname) as shortname FROM Standings_RD3 x, TBLOwners y where x.id = y.ownerid ", objConn,3,3,1
								%>
								<%
								While Not objRSRnd3.EOF
								%>
								<tr style="background-color:white;">
								<td class="big" class="content" style="text-align:left;color;font-weight:700;"><%=objRSRnd3.Fields("shortname").Value %></td>
								<td  class="big"><%=objRSRnd3.Fields("won").Value %></td>
								<td  class="big"><%=objRSRnd3.Fields("loss").Value %></td>
								<td  class="big"><%=round(objRSRnd3.Fields("ppg").Value,1) %></td>
								<td  class="big"><%=round(objRSRnd3.Fields("oppg").Value,1) %></td>
								<td  class="big"><%=round(objRSRnd3.Fields("diff").Value,1) %></td>
								</tr>
								<%
								objRSRnd3.MoveNext
								Wend
								%>
								<%
								objRSRnd3.Close
								%>
								</table>
								</div>
							</div>
						</div>
					</div>				
				</div>
				<div id="prs" class="tab-pane">
						<div class="row">
							<div class="col-md-12 col-sm-12 col-xs-12">
								<div class="panel panel-override">
										<table class="table table-custom-black table-responsive table-bordered table-condensed">
											<tr class="big" style="background-color:yellowgreen;color:black;font-weight:bold;">
												<th class="big" style="text-align:center;vertical-align:middle;" width="14%"><strong>PRS</strong></th>
												<th class="big" style="text-align:left;vertical-align:middle;" width="16%"><strong>Team</strong></th>
												<th class="big" style="text-align:center;vertical-align:middle;" width="14%" ><strong>W</strong></th>
												<th class="big" style="text-align:center;vertical-align:middle;" width="14%" ><strong>L</strong></th>
												<th class="big" style="text-align:center;vertical-align:middle;" width="14%" ><strong>PPG</strong></th>
												<th class="big" style="text-align:center;vertical-align:middle;" width="14%" ><strong>OPPG</strong></th>
												<th class="big" style="text-align:center;vertical-align:middle;" width="14%" ><strong>DIFF</strong></th>
												</tr>
												<%
																	 
												Dim objRSPRS								
												Set objRSPRS    = Server.CreateObject("ADODB.RecordSet")					 
												objRSPRS.Open   "SELECT * FROM tblowners o, standings s where o.ownerId = s.id order by s.prs desc ", objConn,3,3,1
												%>
												<%
													While Not objRSPRS.EOF
												%>
												<%if objRSPRS.Fields("ownerID").Value = ownerid then%>
													<tr class="success big" style="font-weight:bold;text-align:center;vertical-align:middle;">
												<%else%>
													<tr class="big" style="background-color:white;text-align:center;vertical-align:middle;">
												<%end if %>	
														<td class="big" style="text-align:center;vertical-align:middle;"><bluePos><%=objRSPRS.Fields("prs").Value %></bluePos></td>
														<td class="big" style="text-align:left;vertical-align:middle;"><%=objRSPRS.Fields("shortName").Value %></td>
														<td class="big" style="text-align:center;vertical-align:middle;"><%=objRSPRS.Fields("won").Value %></td>
														<td class="big" style="text-align:center;vertical-align:middle;"><%=objRSPRS.Fields("loss").Value %></td>							
														<td class="big" style="text-align:center;vertical-align:middle;"><%=objRSPRS.Fields("ppg").Value %></td>	
														<td class="big" style="text-align:center;vertical-align:middle;"><%=objRSPRS.Fields("oppg").Value %></td>	
														<td class="big" style="text-align:center;vertical-align:middle;"><%=objRSPRS.Fields("diff").Value %></td>	
													</tr>
											<%
											objRSPRS.MoveNext
											Wend
											%>
										 </table>
								</div>
							</div>
						</div>

					<div class="row">		
						<div class="col-md-12 col-sm-12 col-xs-12 big">
								<strong class="big">Power Ranking Score:</strong><small class="big">&nbsp;Cumulative score derived from playing each opponent on game nights. 1 victory awarded for each team you outscore; .5 awarded for ties.</small>
						</div>
					</div>
				</div>
				<div id="cycle" class="tab-pane">
				<%
				Set objRS         = Server.CreateObject("ADODB.RecordSet")
				Set objRSgames    = Server.CreateObject("ADODB.RecordSet")
				Set objRSWork     = Server.CreateObject("ADODB.RecordSet")			
				Set objsrCycle    = Server.CreateObject("ADODB.RecordSet")		

				objsrCycle.Open             "SELECT max (cycle) as currentCycle from Standings_Cycle", objConn,3,3,1		
				w_current_cycle = objsrCycle.Fields("currentCycle").Value
				objsrCycle.Close
				'Response.Write "CYCLE      = "&w_current_cycle&".<br>"
				%>

			<div class="bs-example">
					<ul class="nav nav-tabs">
					
							<%if w_current_cycle = 1 then %>
								<li class="active"><a data-toggle="tab" href="#cycle1">1</a></li>						
							<%else%>
								<li><a data-toggle="tab" href="#cycle1">1</a></li>
							<%end if%>

							<%if w_current_cycle = 2 then %>
								<li class="active"><a data-toggle="tab" href="#cycle2">2</a></li>						
							<%else%>
								<li><a data-toggle="tab" href="#cycle2">2</a></li>
							<%end if%>
							
							<%if w_current_cycle = 3 then %>
								<li class="active"><a data-toggle="tab" href="#cycle3">3</a></li>						
							<%else%>
								<li><a data-toggle="tab" href="#cycle3">3</a></li>
							<%end if%>

							<%if w_current_cycle = 4 then %>
								<li class="active"><a data-toggle="tab" href="#cycle4">4</a></li>						
							<%else%>
								<li><a data-toggle="tab" href="#cycle4">4</a></li>
							<%end if%>							

							<%if w_current_cycle = 5 then %>
								<li class="active"><a data-toggle="tab" href="#cycle5">5</a></li>						
							<%else%>
								<li><a data-toggle="tab" href="#cycle5">5</a></li>
							<%end if%>										

							<%if w_current_cycle = 6 then %>
								<li class="active"><a data-toggle="tab" href="#cycle6">6</a></li>						
							<%else%>
								<li><a data-toggle="tab" href="#cycle6">6</a></li>
							<%end if%>			
	
							<%if w_current_cycle = 7 then %>
								<li class="active"><a data-toggle="tab" href="#cycle7">7</a></li>						
							<%else%>
								<li><a data-toggle="tab" href="#cycle7">7</a></li>
							<%end if%>				
					</ul>
					<div class="tab-content">					  
						<% if w_current_cycle = 1 then %>
						<div id="cycle1" class="tab-pane fade in active">					
						<%else%>
						<div id="cycle1" class="tab-pane fade">
						<%end if%>					
							<h4>Cycle 1</h4>
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
									<%if objRS.Fields("ID").Value = ownerid then%>
									<tr class="success big" style="font-weight:bold;text-align:center;vertical-align:middle;">
										<td class="big" class="content" style="text-align:left;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
									</tr>	
									<%else%>										
									<tr class="big" style="text-align:center;vertical-align:middle;background-color:white;">
										<td class="big" class="content" style="text-align:left;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
									</tr>
									<%end if%>
									<%
									objRS.MoveNext
										Wend
									%>
									<%
									objRS.Close
									%>
							</table>
							<div>
								<h4><span style="color:#9a1400;font-weight:bold;text-decoration:underline;" class="big">Tie Breaker Rules for Cycle Standings</span></h4>
								<ol class="big">
									<li><strong>First place winner receives $25.00</strong></li>
									<li>If multiple teams share identical records at the end of the cycle, it's determined if an owner has a H2H advantage over all tied teams. If so that team will be placed highest in the standings</li>
									<li>If Step 2 does not render a decision, the team with the highest PPG will receive the highest ranking of all tied teams and be awarded $25.00 for winning the cycle.</li>
								</ol>
							</div>							
						</div>
						<% if w_current_cycle = 2 then %>
							<div id="cycle2" class="tab-pane fade in active">					
						<%else%>
							<div id="cycle2" class="tab-pane fade">
						<%end if%>	
							<h4>Cycle 2</h4>
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
									<%if objRS.Fields("ID").Value = ownerid then%>
									<tr class="success big" style="font-weight:bold;text-align:center;vertical-align:middle;">
										<td class="big" class="content" style="text-align:left;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
									</tr>	
									<%else%>										
									<tr class="big" style="text-align:center;vertical-align:middle;background-color:white;">
										<td class="big" class="content" style="text-align:left;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
									</tr>
									<%end if%>
									<%
									objRS.MoveNext
										Wend
									%>
									<%
									objRS.Close
									%>
							</table>
								<div>
								<h4><span style="color:#9a1400;font-weight:bold;text-decoration:underline;" class="big">Tie Breaker Rules for Cycle Standings</span></h4>
								<ol class="big">
									<li><strong>First place winner receives $25.00</strong></li>
									<li>If multiple teams share identical records at the end of the cycle, it's determined if an owner has a H2H advantage over all tied teams. If so that team will be placed highest in the standings</li>
									<li>If Step 2 does not render a decision, the team with the highest PPG will receive the highest ranking of all tied teams and be awarded $25.00 for winning the cycle.</li>
								</ol>
							</div>	
						</div>
						<% if w_current_cycle = 3 then %>
							<div id="cycle3" class="tab-pane fade in active">					
						<%else%>
							<div id="cycle3" class="tab-pane fade">
						<%end if%>	
							<h4>Cycle 3</h4>
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
									<%if objRS.Fields("ID").Value = ownerid then%>
									<tr class="success big" style="font-weight:bold;text-align:center;vertical-align:middle;">
										<td class="big" class="content" style="text-align:left;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
									</tr>	
									<%else%>										
									<tr class="big" style="text-align:center;vertical-align:middle;background-color:white;">
										<td class="big" class="content" style="text-align:left;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
									</tr>
									<%end if%>
									<%
									objRS.MoveNext
										Wend
									%>
									<%
									objRS.Close
									%>
							</table>	
							<div>
								<h4><span style="color:#9a1400;font-weight:bold;text-decoration:underline;" class="big">Tie Breaker Rules for Cycle Standings</span></h4>
								<ol class="big">
									<li><strong>First place winner receives $25.00</strong></li>
									<li>If multiple teams share identical records at the end of the cycle, it's determined if an owner has a H2H advantage over all tied teams. If so that team will be placed highest in the standings</li>
									<li>If Step 2 does not render a decision, the team with the highest PPG will receive the highest ranking of all tied teams and be awarded $25.00 for winning the cycle.</li>
								</ol>
							</div>								
						</div>
						<% if w_current_cycle = 4 then %>
							<div id="cycle4" class="tab-pane fade in active">					
						<%else%>
							<div id="cycle4" class="tab-pane fade">
						<%end if%>	
							<h4>Cycle 4</h4>
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
									<%if objRS.Fields("ID").Value = ownerid then%>
									<tr class="success big" style="font-weight:bold;text-align:center;vertical-align:middle;">
										<td class="big" class="content" style="text-align:left;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
									</tr>	
									<%else%>										
									<tr class="big" style="text-align:center;vertical-align:middle;background-color:white;">
										<td class="big" class="content" style="text-align:left;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
									</tr>
									<%end if%>
									<%
									objRS.MoveNext
										Wend
									%>
									<%
									objRS.Close
									%>
							</table>
							<div>
								<h4><span style="color:#9a1400;font-weight:bold;text-decoration:underline;" class="big">Tie Breaker Rules for Cycle Standings</span></h4>
								<ol class="big">
									<li><strong>First place winner receives $25.00</strong></li>
									<li>If multiple teams share identical records at the end of the cycle, it's determined if an owner has a H2H advantage over all tied teams. If so that team will be placed highest in the standings</li>
									<li>If Step 2 does not render a decision, the team with the highest PPG will receive the highest ranking of all tied teams and be awarded $25.00 for winning the cycle.</li>
								</ol>
							</div>								
						</div>
						<% if w_current_cycle = 5 then %>
							<div id="cycle5" class="tab-pane fade in active">					
						<%else%>
							<div id="cycle5" class="tab-pane fade">
						<%end if%>	
							<h4>Cycle 5</h4>
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
									<%if objRS.Fields("ID").Value = ownerid then%>
									<tr class="success big" style="font-weight:bold;text-align:center;vertical-align:middle;">
										<td class="big" class="content" style="text-align:left;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
									</tr>	
									<%else%>										
									<tr class="big" style="text-align:center;vertical-align:middle;background-color:white;">
										<td class="big" class="content" style="text-align:left;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
									</tr>
									<%end if%>
									<%
									objRS.MoveNext
										Wend
									%>
									<%
									objRS.Close
									%>
							</table>
							<div>
								<h4><span style="color:#9a1400;font-weight:bold;text-decoration:underline;" class="big">Tie Breaker Rules for Cycle Standings</span></h4>
								<ol class="big">
									<li><strong>First place winner receives $25.00</strong></li>
									<li>If multiple teams share identical records at the end of the cycle, it's determined if an owner has a H2H advantage over all tied teams. If so that team will be placed highest in the standings</li>
									<li>If Step 2 does not render a decision, the team with the highest PPG will receive the highest ranking of all tied teams and be awarded $25.00 for winning the cycle.</li>
								</ol>
							</div>								
						</div>
						<% if w_current_cycle = 6 then %>
							<div id="cycle6" class="tab-pane fade in active">					
						<%else%>
							<div id="cycle6" class="tab-pane fade">
						<%end if%>	
							<h4>Cycle 6</h4>
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
									<%if objRS.Fields("ID").Value = ownerid then%>
									<tr class="success big" style="font-weight:bold;text-align:center;vertical-align:middle;">
										<td class="big" class="content" style="text-align:left;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
									</tr>	
									<%else%>										
									<tr class="big" style="text-align:center;vertical-align:middle;background-color:white;">
										<td class="big" class="content" style="text-align:left;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
									</tr>
									<%end if%>
									<%
									objRS.MoveNext
										Wend
									%>
									<%
									objRS.Close
									%>
							</table>
							<div>
								<h4><span style="color:#9a1400;font-weight:bold;text-decoration:underline;" class="big">Tie Breaker Rules for Cycle Standings</span></h4>
								<ol class="big">
									<li><strong>First place winner receives $25.00</strong></li>
									<li>If multiple teams share identical records at the end of the cycle, it's determined if an owner has a H2H advantage over all tied teams. If so that team will be placed highest in the standings</li>
									<li>If Step 2 does not render a decision, the team with the highest PPG will receive the highest ranking of all tied teams and be awarded $25.00 for winning the cycle.</li>
								</ol>
							</div>								
						</div>
						<% if w_current_cycle = 7 then %>
							<div id="cycle7" class="tab-pane fade in active">					
						<%else%>
							<div id="cycle7" class="tab-pane fade">
						<%end if%>	
							<h4>Cycle 7</h4>
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
									<%if objRS.Fields("ID").Value = ownerid then%>
									<tr class="success big" style="font-weight:bold;text-align:center;vertical-align:middle;">
										<td class="big" class="content" style="text-align:left;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
									</tr>	
									<%else%>										
									<tr class="big" style="text-align:center;vertical-align:middle;background-color:white;">
										<td class="big" class="content" style="text-align:left;"><small><%=objRS.Fields("rank").Value %></small>&nbsp;<%=objRS.Fields("team").Value %></td>
										<td class="big"><%=objRS.Fields("won").Value %></td>
										<td class="big"><%=objRS.Fields("loss").Value %></td>
										<td class="big"><%=round(objRS.Fields("ppg").Value,1) %></td>
										<td class="big"><%=round(objRS.Fields("oppg").Value,1) %></td>
									</tr>
									<%end if%>
									<%
									objRS.MoveNext
										Wend
									%>
									<%
									objRS.Close
									%>
							</table>
							<div>
								<h4><span style="color:#9a1400;font-weight:bold;text-decoration:underline;" class="big">Tie Breaker Rules for Cycle Standings</span></h4>
								<ol class="big">
									<li><strong>First place winner receives $25.00</strong></li>
									<li>If multiple teams share identical records at the end of the cycle, it's determined if an owner has a H2H advantage over all tied teams. If so that team will be placed highest in the standings</li>
									<li>If Step 2 does not render a decision, the team with the highest PPG will receive the highest ranking of all tied teams and be awarded $25.00 for winning the cycle.</li>
								</ol>
							</div>								
						</div>
				</div>
			</div>
		</div>

				
				</div>
			</div>	
		</div>		
	</div>
</div>	
<% 
objRSMyGames.Close
objRSgames.Close
objRS.Close        
objRSWork.Close     
objConn.Close
Set objConn = Nothing
 %>
</body>
</html>
